import os
import json
import time
import hmac
import hashlib
import secrets
import threading
from datetime import datetime, timedelta

ROLES = {
    'admin': {
        'label': 'Administrator',
        'description': 'Full access: edit rates, upload PDFs, export reports, and manage users/roles.',
        'permissions': ['view', 'edit_rates', 'upload_docs', 'manage_users'],
        'can_edit': True,
    },
    'editor': {
        'label': 'Accounts Editor',
        'description': 'Operational access: view dashboard, edit discussed rates & surcharges, and upload PDFs.',
        'permissions': ['view', 'edit_rates', 'upload_docs'],
        'can_edit': True,
    },
    'viewer': {
        'label': 'Read-Only Viewer',
        'description': 'Audit access: view dashboard, inspect PDFs, and download reports. Cannot modify data.',
        'permissions': ['view'],
        'can_edit': False,
    },
}


def _hash_password(password: str, salt: str) -> str:
    """Hash password securely using PBKDF2-HMAC-SHA256."""
    dk = hashlib.pbkdf2_hmac('sha256', password.encode('utf-8'), salt.encode('utf-8'), 200_000)
    return dk.hex()


class AuthManager:
    """Thread-safe Authentication, Session, and Role-Based Access Control manager."""

    def __init__(self, data_dir: str):
        self.data_dir = data_dir
        self.users_file = os.path.join(data_dir, 'users.json')
        self.sessions_file = os.path.join(data_dir, 'sessions.json')
        self.lock = threading.Lock()
        self.users = {}
        self.sessions = {}
        self.failed_attempts = {}  # key -> list of failure timestamps
        self.max_failed_attempts = 10
        self.rate_limit_window = 300  # 5 minutes
        self._load_all()

    def _atomic_write(self, path: str, payload: dict):
        os.makedirs(os.path.dirname(path), exist_ok=True)
        tmp = f"{path}.tmp.{os.getpid()}.{threading.get_ident()}.{time.time_ns()}"
        with open(tmp, 'w') as f:
            json.dump(payload, f, indent=2)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp, path)

    def _load_all(self):
        with self.lock:
            if os.path.exists(self.users_file):
                try:
                    with open(self.users_file, 'r') as f:
                        self.users = json.load(f).get('users', {})
                except Exception as e:
                    print(f"Error loading users.json: {e}")
                    self.users = {}
            else:
                self.users = {}

            if not self.users:
                self._seed_default_users()

            if os.path.exists(self.sessions_file):
                try:
                    with open(self.sessions_file, 'r') as f:
                        self.sessions = json.load(f).get('sessions', {})
                except Exception as e:
                    print(f"Error loading sessions.json: {e}")
                    self.sessions = {}
            else:
                self.sessions = {}

            self._prune_expired_sessions_unlocked()

    def _seed_default_users(self):
        now_iso = datetime.now().isoformat()
        defaults = [
            ('admin', 'Mehta Impex Admin', 'admin', 'admin'),
            ('viewer', 'Audit Viewer (Read-Only)', 'viewer', 'Viewer@2026'),
        ]
        for uname, display_name, role, pwd in defaults:
            salt = secrets.token_hex(16)
            self.users[uname] = {
                'username': uname,
                'name': display_name,
                'role': role,
                'salt': salt,
                'password_hash': _hash_password(pwd, salt),
                'active': True,
                'created_at': now_iso,
                'last_login': None,
            }
        self._atomic_write(self.users_file, {'users': self.users, 'updated_at': now_iso})

    def _save_users_unlocked(self):
        self._atomic_write(self.users_file, {
            'users': self.users,
            'updated_at': datetime.now().isoformat()
        })

    def _save_sessions_unlocked(self):
        self._atomic_write(self.sessions_file, {
            'sessions': self.sessions,
            'updated_at': datetime.now().isoformat()
        })

    def _prune_expired_sessions_unlocked(self):
        now_ts = time.time()
        expired = [tok for tok, s in self.sessions.items() if s.get('expires_ts', 0) < now_ts]
        if expired:
            for tok in expired:
                del self.sessions[tok]
            self._save_sessions_unlocked()

    def _public_user_dict(self, u: dict) -> dict:
        role = u.get('role', 'viewer')
        role_meta = ROLES.get(role, ROLES['viewer'])
        return {
            'username': u['username'],
            'name': u.get('name', u['username']),
            'role': role,
            'role_label': role_meta['label'],
            'permissions': list(role_meta['permissions']),
            'can_edit': bool(role_meta['can_edit']),
            'active': bool(u.get('active', True)),
            'created_at': u.get('created_at'),
            'last_login': u.get('last_login'),
        }

    def _check_rate_limit_unlocked(self, key: str):
        now = time.time()
        cutoff = now - self.rate_limit_window
        attempts = [t for t in self.failed_attempts.get(key, []) if t > cutoff]
        if attempts:
            self.failed_attempts[key] = attempts
        elif key in self.failed_attempts:
            del self.failed_attempts[key]
        if len(attempts) >= self.max_failed_attempts:
            retry_after = max(1, int(self.rate_limit_window - (now - attempts[0])))
            return False, retry_after
        return True, 0

    def _record_failure_unlocked(self, key: str):
        now = time.time()
        cutoff = now - self.rate_limit_window
        attempts = [t for t in self.failed_attempts.get(key, []) if t > cutoff]
        attempts.append(now)
        self.failed_attempts[key] = attempts

    def authenticate(self, username: str, password: str, remember: bool = False, client_ip: str = 'local'):
        """Verify credentials with brute-force rate limiting and issue a session token on success."""
        uname = (username or '').strip().lower()
        if not uname or not password:
            return None, 'Please enter both Login ID and Password.'

        ip_key = f"ip:{client_ip or 'unknown'}"
        acct_key = f"acct:{uname}"

        with self.lock:
            ok_ip, wait_ip = self._check_rate_limit_unlocked(ip_key)
            ok_acct, wait_acct = self._check_rate_limit_unlocked(acct_key)
            if not ok_ip or not ok_acct:
                wait_sec = max(wait_ip, wait_acct)
                return {'rate_limited': True, 'retry_after': wait_sec}, f'Too many failed sign-in attempts. Please wait {wait_sec}s before retrying.'

            user = self.users.get(uname)
            if not user:
                self._record_failure_unlocked(ip_key)
                self._record_failure_unlocked(acct_key)
                return None, 'Invalid Login ID or Password.'
            if not user.get('active', True):
                return None, 'This account has been deactivated. Contact an Administrator.'

            expected_hash = user.get('password_hash', '')
            actual_hash = _hash_password(password, user.get('salt', ''))
            if not hmac.compare_digest(expected_hash, actual_hash):
                self._record_failure_unlocked(ip_key)
                self._record_failure_unlocked(acct_key)
                return None, 'Invalid Login ID or Password.'

            self.failed_attempts.pop(ip_key, None)
            self.failed_attempts.pop(acct_key, None)

            ttl_seconds = 7 * 86400 if remember else 86400
            token = secrets.token_urlsafe(32)
            now = datetime.now()
            self.sessions[token] = {
                'username': uname,
                'role': user.get('role', 'viewer'),
                'created_at': now.isoformat(),
                'expires_ts': time.time() + ttl_seconds,
                'expires_at': (now + timedelta(seconds=ttl_seconds)).isoformat(),
            }
            user['last_login'] = now.isoformat()
            self._prune_expired_sessions_unlocked()
            self._save_users_unlocked()
            self._save_sessions_unlocked()
            return {'token': token, 'ttl': ttl_seconds, 'user': self._public_user_dict(user)}, None

    def get_session_user(self, token: str):
        """Resolve a session cookie token to the current active user dict or None."""
        if not token:
            return None
        with self.lock:
            sess = self.sessions.get(token)
            if not sess:
                return None
            if sess.get('expires_ts', 0) < time.time():
                del self.sessions[token]
                self._save_sessions_unlocked()
                return None
            user = self.users.get(sess.get('username'))
            if not user or not user.get('active', True):
                return None
            return self._public_user_dict(user)

    def logout(self, token: str):
        if not token:
            return
        with self.lock:
            if token in self.sessions:
                del self.sessions[token]
                self._save_sessions_unlocked()

    def has_permission(self, user: dict, permission: str) -> bool:
        if not user:
            return False
        return permission in user.get('permissions', [])

    def list_users(self) -> list:
        with self.lock:
            return [self._public_user_dict(u) for u in sorted(self.users.values(), key=lambda x: x['username'])]

    def change_own_password(self, username: str, current_password: str, new_password: str):
        uname = (username or '').strip().lower()
        if not new_password or len(new_password) < 4:
            return False, 'New password must be at least 4 characters long.'
        with self.lock:
            user = self.users.get(uname)
            if not user:
                return False, 'User account not found.'
            if not hmac.compare_digest(user['password_hash'], _hash_password(current_password or '', user['salt'])):
                return False, 'Current password is incorrect.'
            new_salt = secrets.token_hex(16)
            user['salt'] = new_salt
            user['password_hash'] = _hash_password(new_password, new_salt)
            self._save_users_unlocked()
            return True, 'Password updated successfully.'

    def create_user(self, username: str, name: str, role: str, password: str):
        uname = (username or '').strip().lower()
        if not uname or not uname.replace('_', '').replace('-', '').replace('.', '').isalnum():
            return None, 'Login ID must be alphanumeric (letters, numbers, _, -, . allowed).'
        if role not in ROLES:
            return None, f'Invalid role: {role}'
        if not password or len(password) < 4:
            return None, 'Password must be at least 4 characters long.'

        with self.lock:
            if uname in self.users:
                return None, f"Login ID '{uname}' already exists."
            salt = secrets.token_hex(16)
            self.users[uname] = {
                'username': uname,
                'name': (name or uname).strip(),
                'role': role,
                'salt': salt,
                'password_hash': _hash_password(password, salt),
                'active': True,
                'created_at': datetime.now().isoformat(),
                'last_login': None,
            }
            self._save_users_unlocked()
            return self._public_user_dict(self.users[uname]), None

    def update_user(self, acting_username: str, target_username: str, role=None, name=None, active=None, new_password=None):
        t_uname = (target_username or '').strip().lower()
        a_uname = (acting_username or '').strip().lower()
        with self.lock:
            user = self.users.get(t_uname)
            if not user:
                return None, f"User '{t_uname}' not found."

            # Safety check: never allow removing or deactivating the last active admin
            new_role = role if role is not None else user['role']
            new_active = bool(active) if active is not None else user.get('active', True)
            if new_role not in ROLES:
                return None, f"Invalid role: {new_role}"

            if user['role'] == 'admin' and (new_role != 'admin' or not new_active):
                active_admins = [
                    u for u in self.users.values()
                    if u['role'] == 'admin' and u.get('active', True) and u['username'] != t_uname
                ]
                if not active_admins:
                    return None, 'Cannot demote or deactivate the only active Administrator account.'

            if t_uname == a_uname and not new_active:
                return None, 'You cannot deactivate your own currently logged-in account.'

            if name is not None and str(name).strip():
                user['name'] = str(name).strip()
            user['role'] = new_role
            user['active'] = new_active

            if new_password:
                if len(new_password) < 4:
                    return None, 'New password must be at least 4 characters long.'
                new_salt = secrets.token_hex(16)
                user['salt'] = new_salt
                user['password_hash'] = _hash_password(new_password, new_salt)

            # Revoke active sessions if deactivated
            if not new_active:
                to_del = [tok for tok, s in self.sessions.items() if s.get('username') == t_uname]
                for tok in to_del:
                    del self.sessions[tok]
                if to_del:
                    self._save_sessions_unlocked()

            self._save_users_unlocked()
            return self._public_user_dict(user), None

    def delete_user(self, acting_username: str, target_username: str):
        t_uname = (target_username or '').strip().lower()
        a_uname = (acting_username or '').strip().lower()
        if t_uname == a_uname:
            return False, 'You cannot delete your own currently logged-in account.'
        with self.lock:
            user = self.users.get(t_uname)
            if not user:
                return False, f"User '{t_uname}' not found."
            if user['role'] == 'admin' and user.get('active', True):
                other_admins = [
                    u for u in self.users.values()
                    if u['role'] == 'admin' and u.get('active', True) and u['username'] != t_uname
                ]
                if not other_admins:
                    return False, 'Cannot delete the only active Administrator account.'
            del self.users[t_uname]
            to_del = [tok for tok, s in self.sessions.items() if s.get('username') == t_uname]
            for tok in to_del:
                del self.sessions[tok]
            self._save_users_unlocked()
            if to_del:
                self._save_sessions_unlocked()
            return True, f"User '{t_uname}' deleted."
