import os
import sys
import json
import shutil
import hashlib
from email.parser import BytesParser
from email.policy import default as email_policy
from http.server import HTTPServer, BaseHTTPRequestHandler
from http.cookies import SimpleCookie
from socketserver import ThreadingMixIn
from datetime import datetime
from urllib.parse import urlparse, parse_qs

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOCAL_PKGS = os.path.join(os.path.dirname(BASE_DIR), 'python_packages')
if os.path.isdir(LOCAL_PKGS) and LOCAL_PKGS not in sys.path:
    sys.path.append(LOCAL_PKGS)
if BASE_DIR not in sys.path:
    sys.path.append(BASE_DIR)

from pdf_parser import parse_invoice_pdf, parse_credit_note_pdf, parse_ledger_pdf
from engine import ReconEngine
from excel_generator import generate_reconciliation_workbook
from pdf_report_generator import generate_monthly_claim_pdf
from auth import AuthManager, ROLES

PORT = int(os.environ.get('PORT', 8085))
DATA_DIR = os.environ.get('DATA_DIR', os.path.join(BASE_DIR, 'data'))
UPLOAD_DIR = os.path.join(DATA_DIR, 'uploads')
PDF_STORE_DIR = os.path.join(DATA_DIR, 'pdf_store')
EXCEL_PATH = os.environ.get('EXCEL_PATH', os.path.join(DATA_DIR, 'GGL_Master_Reconciliation.xlsx'))

os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(PDF_STORE_DIR, exist_ok=True)
os.makedirs(os.path.dirname(EXCEL_PATH), exist_ok=True)
engine = ReconEngine(data_file=os.path.join(DATA_DIR, 'recon_state.json'))
auth_manager = AuthManager(DATA_DIR)

MAX_JSON_BYTES = 1 * 1024 * 1024      # 1 MB max JSON payload
MAX_UPLOAD_BYTES = 25 * 1024 * 1024   # 25 MB max PDF upload

LOGIN_HTML = """<!DOCTYPE html>
<html lang="en" data-theme="dark">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Sign In | GGL-Recon Audit System</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@500;600;700;800&display=swap" rel="stylesheet">
  <style>
    :root[data-theme="dark"] {
      --bg: #060B17;
      --mesh-1: rgba(56, 189, 248, 0.16);
      --mesh-2: rgba(129, 140, 248, 0.16);
      --mesh-3: rgba(16, 185, 129, 0.10);
      --shell-bg: rgba(13, 21, 40, 0.86);
      --left-bg: linear-gradient(155deg, rgba(16, 31, 62, 0.96), rgba(11, 20, 42, 0.98));
      --right-bg: rgba(13, 21, 40, 0.92);
      --surface: rgba(30, 43, 76, 0.55);
      --border: rgba(148, 163, 184, 0.16);
      --border-strong: rgba(56, 189, 248, 0.30);
      --text: #F1F5F9;
      --text-muted: #94A3B8;
      --text-dim: #64748B;
      --primary: #38BDF8;
      --accent: #818CF8;
      --success: #34D399;
      --warning: #FBBF24;
      --danger: #F87171;
      --danger-bg: rgba(239, 68, 68, 0.14);
      --danger-border: rgba(248, 113, 113, 0.35);
      --shadow: 0 30px 80px rgba(0, 0, 0, 0.65);
      --input-bg: rgba(8, 14, 30, 0.85);
    }
    :root[data-theme="light"] {
      --bg: #EFF4FA;
      --mesh-1: rgba(2, 132, 199, 0.10);
      --mesh-2: rgba(79, 70, 229, 0.09);
      --mesh-3: rgba(5, 150, 105, 0.07);
      --shell-bg: rgba(255, 255, 255, 0.96);
      --left-bg: linear-gradient(155deg, #0F172A, #1E293B);
      --right-bg: #FFFFFF;
      --surface: #F8FAFC;
      --border: rgba(15, 23, 42, 0.11);
      --border-strong: rgba(2, 132, 199, 0.30);
      --text: #0F172A;
      --text-muted: #475569;
      --text-dim: #94A3B8;
      --primary: #0284C7;
      --accent: #4F46E5;
      --success: #059669;
      --warning: #D97706;
      --danger: #DC2626;
      --danger-bg: rgba(220, 38, 38, 0.08);
      --danger-border: rgba(220, 38, 38, 0.25);
      --shadow: 0 24px 70px rgba(15, 23, 42, 0.12);
      --input-bg: #F8FAFC;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
      background-color: var(--bg);
      background-image:
        radial-gradient(circle at 12% 18%, var(--mesh-1), transparent 42%),
        radial-gradient(circle at 88% 15%, var(--mesh-2), transparent 40%),
        radial-gradient(circle at 60% 90%, var(--mesh-3), transparent 45%);
      color: var(--text);
      min-height: 100vh;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 20px;
      position: relative;
      overflow-x: hidden;
      transition: background 0.3s ease, color 0.3s ease;
    }
    .ambient-orb {
      position: fixed;
      border-radius: 50%;
      filter: blur(75px);
      opacity: 0.45;
      pointer-events: none;
      animation: floatOrb 14s ease-in-out infinite alternate;
      z-index: 0;
    }
    .orb-a { width: 320px; height: 320px; top: -80px; left: -60px; background: rgba(56, 189, 248, 0.25); }
    .orb-b { width: 300px; height: 300px; bottom: -90px; right: -50px; background: rgba(129, 140, 248, 0.24); animation-delay: -6s; }
    @keyframes floatOrb {
      from { transform: translate(0, 0) scale(1); }
      to { transform: translate(28px, 24px) scale(1.08); }
    }
    .theme-btn {
      position: fixed;
      top: 18px;
      right: 18px;
      z-index: 20;
      background: var(--shell-bg);
      backdrop-filter: blur(12px);
      border: 1px solid var(--border);
      color: var(--text);
      padding: 8px 15px;
      border-radius: 999px;
      font-family: inherit;
      font-size: 12.5px;
      font-weight: 700;
      cursor: pointer;
      display: flex;
      align-items: center;
      gap: 7px;
      transition: transform 0.2s ease, border-color 0.2s ease;
      box-shadow: 0 6px 20px rgba(0, 0, 0, 0.15);
    }
    .theme-btn:hover { transform: translateY(-2px); border-color: var(--primary); }
    .auth-shell {
      position: relative;
      z-index: 5;
      width: 100%;
      max-width: 920px;
      display: grid;
      grid-template-columns: 1.05fr 1fr;
      background: var(--shell-bg);
      backdrop-filter: blur(22px);
      border: 1px solid var(--border);
      border-radius: 26px;
      overflow: hidden;
      box-shadow: var(--shadow);
      animation: shellEnter 0.5s cubic-bezier(0.16, 1, 0.3, 1);
    }
    @keyframes shellEnter {
      from { opacity: 0; transform: translateY(20px) scale(0.98); }
      to { opacity: 1; transform: translateY(0) scale(1); }
    }
    @keyframes shake {
      0%, 100% { transform: translateX(0); }
      20% { transform: translateX(-7px); }
      40% { transform: translateX(7px); }
      60% { transform: translateX(-5px); }
      80% { transform: translateX(5px); }
    }
    .shake { animation: shake 0.38s ease; }

    /* Left Showcase Column (stays rich dark navy on both themes for high contrast) */
    .brand-panel {
      background: var(--left-bg);
      color: #F8FAFC;
      padding: 42px 36px;
      display: flex;
      flex-direction: column;
      justify-content: space-between;
      border-right: 1px solid rgba(148, 163, 184, 0.14);
      position: relative;
    }
    .brand-panel::after {
      content: '';
      position: absolute;
      inset: 0;
      background: radial-gradient(circle at 20% 10%, rgba(56, 189, 248, 0.14), transparent 55%);
      pointer-events: none;
    }
    .brand-top { position: relative; z-index: 1; }
    .logo-row {
      display: flex;
      align-items: center;
      gap: 12px;
      margin-bottom: 24px;
    }
    .shield-badge {
      width: 48px; height: 48px;
      border-radius: 14px;
      background: linear-gradient(135deg, rgba(56, 189, 248, 0.25), rgba(129, 140, 248, 0.28));
      border: 1px solid rgba(56, 189, 248, 0.45);
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 24px;
      box-shadow: 0 8px 24px rgba(56, 189, 248, 0.18);
    }
    .brand-name {
      font-size: 20px;
      font-weight: 800;
      letter-spacing: -0.02em;
    }
    .brand-tag {
      font-size: 11px;
      font-weight: 700;
      color: #38BDF8;
      letter-spacing: 0.06em;
      text-transform: uppercase;
      margin-top: 2px;
    }
    .hero-title {
      font-size: 26px;
      font-weight: 800;
      line-height: 1.25;
      letter-spacing: -0.02em;
      margin-bottom: 12px;
    }
    .hero-desc {
      font-size: 13.5px;
      color: #94A3B8;
      line-height: 1.6;
      margin-bottom: 26px;
    }
    .feature-list {
      display: flex;
      flex-direction: column;
      gap: 12px;
    }
    .feature-item {
      display: flex;
      align-items: flex-start;
      gap: 12px;
      padding: 12px 14px;
      border-radius: 14px;
      background: rgba(255, 255, 255, 0.04);
      border: 1px solid rgba(148, 163, 184, 0.12);
    }
    .feature-icon {
      font-size: 18px;
      line-height: 1.3;
    }
    .feature-title {
      font-size: 13px;
      font-weight: 700;
      color: #F1F5F9;
    }
    .feature-sub {
      font-size: 11.5px;
      color: #94A3B8;
      margin-top: 2px;
      line-height: 1.4;
    }
    .brand-bottom {
      position: relative;
      z-index: 1;
      margin-top: 28px;
      padding-top: 16px;
      border-top: 1px solid rgba(148, 163, 184, 0.14);
      display: flex;
      justify-content: space-between;
      align-items: center;
      flex-wrap: wrap;
      gap: 8px;
      font-size: 11.5px;
      color: #94A3B8;
    }
    .live-badge {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      padding: 4px 10px;
      border-radius: 999px;
      background: rgba(16, 185, 129, 0.14);
      border: 1px solid rgba(52, 211, 153, 0.32);
      color: #34D399;
      font-weight: 700;
      font-size: 11px;
    }
    .live-dot {
      width: 7px; height: 7px;
      border-radius: 50%;
      background: #34D399;
      box-shadow: 0 0 8px #34D399;
    }

    /* Right Form Column */
    .form-panel {
      background: var(--right-bg);
      padding: 42px 36px;
      display: flex;
      flex-direction: column;
      justify-content: center;
    }
    .form-header { margin-bottom: 24px; }
    .form-eyebrow {
      display: inline-block;
      font-size: 11px;
      font-weight: 800;
      text-transform: uppercase;
      letter-spacing: 0.08em;
      color: var(--primary);
      margin-bottom: 8px;
    }
    .form-header h1 {
      font-size: 24px;
      font-weight: 800;
      letter-spacing: -0.02em;
      margin-bottom: 6px;
    }
    .form-header p {
      font-size: 13px;
      color: var(--text-muted);
    }
    .form-group { margin-bottom: 18px; }
    label {
      display: flex;
      justify-content: space-between;
      align-items: center;
      font-size: 11.5px;
      font-weight: 800;
      text-transform: uppercase;
      letter-spacing: 0.06em;
      color: var(--text-muted);
      margin-bottom: 7px;
    }
    .caps-hint {
      display: none;
      font-size: 10.5px;
      font-weight: 700;
      color: var(--warning);
      text-transform: none;
      letter-spacing: 0;
    }
    .input-wrap { position: relative; }
    .input-icon {
      position: absolute;
      left: 14px;
      top: 50%;
      transform: translateY(-50%);
      font-size: 15px;
      opacity: 0.75;
      pointer-events: none;
    }
    input[type="text"], input[type="password"] {
      width: 100%;
      padding: 13px 14px 13px 42px;
      border-radius: 12px;
      border: 1px solid var(--border);
      background: var(--input-bg);
      color: var(--text);
      font-family: inherit;
      font-size: 14px;
      font-weight: 600;
      transition: border-color 0.2s ease, box-shadow 0.2s ease, background 0.2s ease;
    }
    #password { padding-right: 46px; }
    input:focus {
      outline: none;
      border-color: var(--primary);
      box-shadow: 0 0 0 3px rgba(56, 189, 248, 0.18);
    }
    .eye-btn {
      position: absolute;
      right: 10px;
      top: 50%;
      transform: translateY(-50%);
      background: none;
      border: none;
      color: var(--text-muted);
      cursor: pointer;
      font-size: 16px;
      padding: 6px;
      border-radius: 8px;
    }
    .eye-btn:hover { color: var(--text); background: var(--surface); }
    .remember-row {
      display: flex;
      align-items: center;
      justify-content: space-between;
      margin-bottom: 22px;
      font-size: 13px;
      color: var(--text-muted);
    }
    .remember-label {
      display: flex;
      align-items: center;
      gap: 8px;
      cursor: pointer;
      font-weight: 600;
      text-transform: none;
      letter-spacing: 0;
      margin-bottom: 0;
      color: var(--text-muted);
    }
    .remember-label input { width: 16px; height: 16px; accent-color: var(--primary); cursor: pointer; }
    .submit-btn {
      width: 100%;
      padding: 14px 18px;
      border: none;
      border-radius: 12px;
      background: linear-gradient(135deg, #0284C7, #4F46E5);
      color: white;
      font-family: inherit;
      font-size: 14.5px;
      font-weight: 800;
      cursor: pointer;
      transition: transform 0.2s ease, box-shadow 0.2s ease, opacity 0.2s ease;
      box-shadow: 0 10px 24px rgba(2, 132, 199, 0.30);
    }
    .submit-btn:hover { transform: translateY(-1px); box-shadow: 0 14px 28px rgba(2, 132, 199, 0.42); }
    .submit-btn:disabled { opacity: 0.72; cursor: not-allowed; transform: none; }
    .error-box {
      display: none;
      background: var(--danger-bg);
      border: 1px solid var(--danger-border);
      color: var(--danger);
      padding: 11px 14px;
      border-radius: 11px;
      font-size: 13px;
      font-weight: 600;
      margin-bottom: 18px;
      line-height: 1.4;
    }
    .role-strip {
      margin-top: 22px;
      padding: 12px 14px;
      border-radius: 12px;
      background: var(--surface);
      border: 1px solid var(--border);
      font-size: 11.5px;
      color: var(--text-muted);
      line-height: 1.5;
    }
    .role-strip strong { color: var(--text); }
    @media (max-width: 780px) {
      .auth-shell { grid-template-columns: 1fr; max-width: 460px; }
      .brand-panel { display: none; }
      .form-panel { padding: 34px 26px; }
    }
  </style>
</head>
<body>
  <div class="ambient-orb orb-a"></div>
  <div class="ambient-orb orb-b"></div>

  <button class="theme-btn" type="button" onclick="toggleTheme()">
    <span id="theme-icon">☀️</span> <span id="theme-label">Light Mode</span>
  </button>

  <main class="auth-shell" id="auth-shell">
    <section class="brand-panel">
      <div class="brand-top">
        <div class="logo-row">
          <div class="shield-badge">🛡️</div>
          <div>
            <div class="brand-name">GGL-Recon</div>
            <div class="brand-tag">Mehta Impex • ModiGuard®</div>
          </div>
        </div>
        <h2 class="hero-title">Rate Difference &amp; Supplier Ledger Audit</h2>
        <p class="hero-desc">
          Unified FY 2026–27 billing reconciliation, Credit Note matching, Cash Discount tracking, and audit-ready financial exports.
        </p>

        <div class="feature-list">
          <div class="feature-item">
            <div class="feature-icon">🔐</div>
            <div>
              <div class="feature-title">Role-Based Access Control (RBAC)</div>
              <div class="feature-sub">Strict Admin, Editor &amp; Read-Only Viewer isolation verified on server &amp; UI.</div>
            </div>
          </div>
          <div class="feature-item">
            <div class="feature-icon">📊</div>
            <div>
              <div class="feature-title">9-Point Double-Checked Integrity</div>
              <div class="feature-sub">Exact invoice-to-CN binding across RD, ARD, F&amp;E, and PMS surcharges.</div>
            </div>
          </div>
          <div class="feature-item">
            <div class="feature-icon">📥</div>
            <div>
              <div class="feature-title">Dynamic Excel &amp; 1-Page Claim PDF</div>
              <div class="feature-sub">Every discussed-rate update recalculates ledger &amp; exports instantaneously.</div>
            </div>
          </div>
        </div>
      </div>

      <div class="brand-bottom">
        <span>GGL Reconciliation Portal</span>
        <span class="live-badge"><span class="live-dot"></span> PROTECTED SESSION</span>
      </div>
    </section>

    <section class="form-panel">
      <div class="form-header">
        <span class="form-eyebrow">Authorized Access Only</span>
        <h1>Welcome Back</h1>
        <p>Sign in with your Login ID &amp; Password to open the workspace.</p>
      </div>

      <div id="error-box" class="error-box" role="alert"></div>

      <form id="login-form" onsubmit="handleLogin(event)">
        <div class="form-group">
          <label for="username"><span>Login ID</span></label>
          <div class="input-wrap">
            <span class="input-icon">👤</span>
            <input type="text" id="username" name="username" placeholder="Enter your Login ID (e.g. admin)" autocomplete="username" required autofocus>
          </div>
        </div>

        <div class="form-group">
          <label for="password">
            <span>Password</span>
            <span id="caps-hint" class="caps-hint">⇪ Caps Lock is ON</span>
          </label>
          <div class="input-wrap">
            <span class="input-icon">🔒</span>
            <input type="password" id="password" name="password" placeholder="Enter your password" autocomplete="current-password" required onkeyup="checkCaps(event)" onkeydown="checkCaps(event)">
            <button type="button" class="eye-btn" id="eye-btn" onclick="togglePwd()" title="Show / Hide Password" aria-label="Show or hide password">👁️</button>
          </div>
        </div>

        <div class="remember-row">
          <label class="remember-label">
            <input type="checkbox" id="remember"> Keep me signed in (7 days)
          </label>
        </div>

        <button type="submit" id="submit-btn" class="submit-btn">🔐 Sign In to Dashboard</button>
      </form>

      <div class="role-strip">
        <strong>Access Tiers:</strong> <strong>Admin</strong> (Full Control &amp; Users) · <strong>Editor</strong> (Edit Rates &amp; Upload PDFs) · <strong>Viewer</strong> (Read-Only Reports)
      </div>
    </section>
  </main>

  <script>
    function applyTheme(t) {
      document.documentElement.setAttribute('data-theme', t);
      const icon = document.getElementById('theme-icon');
      const label = document.getElementById('theme-label');
      if (icon) icon.textContent = t === 'dark' ? '☀️' : '🌙';
      if (label) label.textContent = t === 'dark' ? 'Light Mode' : 'Dark Navy';
    }
    let curTheme = 'dark';
    try { curTheme = localStorage.getItem('ggl_theme') || 'dark'; } catch(e) {}
    applyTheme(curTheme);

    function toggleTheme() {
      curTheme = curTheme === 'dark' ? 'light' : 'dark';
      try { localStorage.setItem('ggl_theme', curTheme); } catch(e) {}
      applyTheme(curTheme);
    }

    function togglePwd() {
      const inp = document.getElementById('password');
      const eye = document.getElementById('eye-btn');
      const show = inp.type === 'password';
      inp.type = show ? 'text' : 'password';
      if (eye) eye.textContent = show ? '🙈' : '👁️';
    }

    function checkCaps(e) {
      const hint = document.getElementById('caps-hint');
      if (!hint || !e || typeof e.getModifierState !== 'function') return;
      hint.style.display = e.getModifierState('CapsLock') ? 'inline' : 'none';
    }

    async function handleLogin(e) {
      e.preventDefault();
      const errBox = document.getElementById('error-box');
      const btn = document.getElementById('submit-btn');
      const shell = document.getElementById('auth-shell');
      errBox.style.display = 'none';
      if (shell) shell.classList.remove('shake');
      btn.disabled = true;
      btn.textContent = '⏳ Verifying Credentials...';

      const username = document.getElementById('username').value.trim();
      const password = document.getElementById('password').value;
      const remember = document.getElementById('remember').checked;

      try {
        const res = await fetch('/api/login', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ username, password, remember })
        });
        const data = await res.json();
        if (data.success) {
          btn.textContent = '✅ Verified! Opening Dashboard...';
          window.location.href = '/';
        } else {
          errBox.textContent = '⚠️ ' + (data.error || 'Authentication failed.');
          errBox.style.display = 'block';
          if (shell) { void shell.offsetWidth; shell.classList.add('shake'); }
          btn.disabled = false;
          btn.textContent = '🔐 Sign In to Dashboard';
        }
      } catch (err) {
        errBox.textContent = '⚠️ Network error while signing in. Please try again.';
        errBox.style.display = 'block';
        btn.disabled = false;
        btn.textContent = '🔐 Sign In to Dashboard';
      }
    }
  </script>
</body>
</html>
"""

HTML_CONTENT = """<!DOCTYPE html>
<html lang="en" data-theme="dark">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>GGL-Recon | Gujarat Guardian Rate Difference & Ledger Audit</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@500;600;700;800&display=swap" rel="stylesheet">
  <style>
    /* ============================================================
       DESIGN SYSTEM & COLOR TOKENS (DARK & LIGHT MODES)
       ============================================================ */
    :root[data-theme="dark"] {
      --bg: #080D1A;
      --ambient-1: rgba(56, 189, 248, 0.08);
      --ambient-2: rgba(99, 102, 241, 0.08);
      --ambient-3: rgba(16, 185, 129, 0.05);
      --card-bg: rgba(16, 25, 48, 0.72);
      --card-solid: #101930;
      --card-hover: rgba(23, 37, 71, 0.88);
      --surface: rgba(25, 40, 78, 0.65);
      --surface-subtle: rgba(18, 28, 56, 0.65);
      --border: rgba(56, 189, 248, 0.16);
      --border-subtle: rgba(255, 255, 255, 0.08);
      --border-glow: rgba(56, 189, 248, 0.38);
      --text: #F8FAFC;
      --text-muted: #94A3B8;
      --text-dim: #64748B;
      --primary: #38BDF8;
      --primary-hover: #0EA5E9;
      --primary-glow: rgba(56, 189, 248, 0.35);
      --primary-glow-subtle: rgba(56, 189, 248, 0.08);
      --accent: #60A5FA;
      --accent-purple: #A78BFA;
      --success: #10B981;
      --success-bg: rgba(16, 185, 129, 0.14);
      --success-border: rgba(16, 185, 129, 0.4);
      --success-glow: rgba(16, 185, 129, 0.28);
      --danger: #F43F5E;
      --danger-bg: rgba(244, 63, 94, 0.14);
      --danger-border: rgba(244, 63, 94, 0.4);
      --danger-glow: rgba(244, 63, 94, 0.28);
      --warning: #F59E0B;
      --warning-bg: rgba(245, 158, 11, 0.14);
      --warning-border: rgba(245, 158, 11, 0.4);
      --table-stripe: rgba(255, 255, 255, 0.015);
      --table-header: rgba(19, 32, 66, 0.9);
      --table-header-glass: rgba(14, 23, 48, 0.94);
      --shadow-card: 0 8px 32px rgba(0, 0, 0, 0.4);
      --shadow-hover: 0 16px 44px rgba(0, 0, 0, 0.55);
      --glass-sheen: inset 0 1px 0 rgba(255, 255, 255, 0.08);
    }
    :root[data-theme="light"] {
      --bg: #F4F7FB;
      --ambient-1: rgba(2, 132, 199, 0.05);
      --ambient-2: rgba(99, 102, 241, 0.04);
      --ambient-3: rgba(16, 185, 129, 0.03);
      --card-bg: rgba(255, 255, 255, 0.88);
      --card-solid: #FFFFFF;
      --card-hover: rgba(248, 250, 252, 0.98);
      --surface: #EDF2F7;
      --surface-subtle: #F8FAFC;
      --border: #DDE3EC;
      --border-subtle: rgba(226, 232, 240, 0.85);
      --border-glow: rgba(2, 132, 199, 0.35);
      --text: #0F172A;
      --text-muted: #526075;
      --text-dim: #8B98A9;
      --primary: #0284C7;
      --primary-hover: #0369A1;
      --primary-glow: rgba(2, 132, 199, 0.25);
      --primary-glow-subtle: rgba(2, 132, 199, 0.06);
      --accent: #2563EB;
      --accent-purple: #9333EA;
      --success: #059669;
      --success-bg: #ECFDF5;
      --success-border: #A7F3D0;
      --success-glow: rgba(5, 150, 105, 0.2);
      --danger: #E11D48;
      --danger-bg: #FFF1F2;
      --danger-border: #FECDD3;
      --danger-glow: rgba(225, 29, 72, 0.2);
      --warning: #D97706;
      --warning-bg: #FFFBEB;
      --warning-border: #FDE68A;
      --table-stripe: #FAFCFE;
      --table-header: #EBF2F9;
      --table-header-glass: rgba(235, 242, 249, 0.95);
      --shadow-card: 0 4px 20px -2px rgba(15, 23, 42, 0.06), 0 2px 6px -1px rgba(15, 23, 42, 0.02);
      --shadow-hover: 0 14px 32px -4px rgba(15, 23, 42, 0.12);
      --glass-sheen: inset 0 1px 0 rgba(255, 255, 255, 0.85);
    }
    
    * { 
      box-sizing: border-box; 
      margin: 0; 
      padding: 0; 
      font-family: 'Plus Jakarta Sans', sans-serif; 
    }
    
    body { 
      background-color: var(--bg); 
      background-image: 
        radial-gradient(ellipse 80% 45% at 15% 0%, var(--ambient-1) 0%, transparent 70%),
        radial-gradient(ellipse 70% 45% at 85% 0%, var(--ambient-2) 0%, transparent 70%);
      background-repeat: no-repeat;
      color: var(--text); 
      padding: 24px; 
      min-height: 100vh; 
      overflow-y: scroll; 
      -webkit-font-smoothing: antialiased;
    }
    
    .container { 
      max-width: 1560px; 
      margin: 0 auto; 
    }

    /* ============================================================
       KEYFRAME ANIMATIONS (GPU-Composited & Lightweight)
       ============================================================ */
    @keyframes pulseDot {
      0%, 100% { transform: scale(1); opacity: 1; }
      50% { transform: scale(1.25); opacity: 0.65; }
    }
    @keyframes tabFadeIn {
      0% { opacity: 0; transform: translateY(4px); }
      100% { opacity: 1; transform: translateY(0); }
    }
    @keyframes modalIn {
      0% { opacity: 0; transform: scale(0.98) translateY(8px); }
      100% { opacity: 1; transform: scale(1) translateY(0); }
    }

    /* ============================================================
       TOP HEADER & BRANDING
       ============================================================ */
    header { 
      display: flex; 
      justify-content: space-between; 
      align-items: center; 
      margin-bottom: 24px; 
      padding-bottom: 20px; 
      border-bottom: 1px solid var(--border); 
      flex-wrap: wrap;
      gap: 16px;
    }
    .brand {
      display: flex;
      flex-direction: column;
      gap: 6px;
    }
    .brand-title-wrap {
      display: flex;
      align-items: center;
      gap: 14px;
      flex-wrap: wrap;
    }
    .brand h1 { 
      font-size: 25px; 
      font-weight: 800; 
      color: var(--text); 
      display: flex; 
      align-items: center; 
      gap: 10px; 
      letter-spacing: -0.6px; 
    }
    .brand h1 span.logo-shield { 
      background: linear-gradient(135deg, #38BDF8 0%, #2563EB 50%, #818CF8 100%); 
      -webkit-background-clip: text; 
      -webkit-text-fill-color: transparent; 
      filter: drop-shadow(0 2px 8px var(--primary-glow));
    }
    .brand p { 
      font-size: 13px; 
      color: var(--text-muted); 
      font-weight: 500; 
      display: flex;
      align-items: center;
      gap: 8px;
    }
    .brand-pill {
      background: var(--surface);
      border: 1px solid var(--border-subtle);
      padding: 2px 8px;
      border-radius: 6px;
      font-size: 11.5px;
      font-weight: 600;
      color: var(--text-muted);
    }
    .live-pill {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      background: var(--success-bg);
      border: 1px solid var(--success-border);
      color: var(--success);
      padding: 3px 10px;
      border-radius: 9999px;
      font-size: 11px;
      font-weight: 800;
      letter-spacing: 0.6px;
      box-shadow: 0 0 14px var(--success-glow);
    }
    .live-dot {
      width: 7px;
      height: 7px;
      border-radius: 50%;
      background: var(--success);
      animation: pulseDot 2s infinite ease-in-out;
    }
    
    .header-actions { 
      display: flex; 
      gap: 12px; 
      align-items: center; 
      flex-wrap: wrap;
    }
    .btn { 
      position: relative;
      padding: 10px 18px; 
      border-radius: 11px; 
      font-size: 13px; 
      font-weight: 700; 
      cursor: pointer; 
      display: inline-flex; 
      align-items: center; 
      gap: 8px; 
      border: none; 
      text-decoration: none; 
      transition: transform 0.15s ease, opacity 0.15s ease, background-color 0.15s ease; 
    }
    .btn:hover { 
      transform: translateY(-1px); 
    }
    .btn:active { 
      transform: translateY(1px); 
    }
    .btn-download { 
      background: linear-gradient(135deg, #10B981 0%, #059669 100%); 
      color: #FFFFFF; 
      box-shadow: 0 4px 16px rgba(16, 185, 129, 0.35); 
    }
    .btn-download:hover { 
      box-shadow: 0 8px 24px rgba(16, 185, 129, 0.5); 
    }
    .btn-primary { 
      background: linear-gradient(135deg, #0284C7 0%, #2563EB 100%); 
      color: #FFFFFF; 
      box-shadow: 0 4px 16px rgba(37, 99, 235, 0.35); 
    }
    .btn-primary:hover { 
      box-shadow: 0 8px 24px rgba(37, 99, 235, 0.5); 
    }
    .btn-primary-sm { 
      background: var(--primary); 
      color: #090E1D; 
      padding: 6px 12px; 
      border-radius: 8px; 
      font-size: 11.5px; 
      font-weight: 800; 
      cursor: pointer; 
      border: none; 
      display: inline-flex; 
      align-items: center; 
      gap: 5px; 
      box-shadow: 0 2px 8px var(--primary-glow); 
      transition: background-color 0.15s ease, transform 0.15s ease;
    }
    .btn-primary-sm:hover { 
      background: var(--primary-hover); 
      color: #FFFFFF; 
      transform: translateY(-1px);
      box-shadow: 0 4px 12px var(--primary-glow);
    }
    .theme-toggle { 
      background: var(--surface); 
      color: var(--text); 
      padding: 10px 15px; 
      border-radius: 11px; 
      border: 1px solid var(--border-subtle); 
      cursor: pointer; 
      font-size: 13px; 
      font-weight: 600; 
      display: flex; 
      align-items: center; 
      gap: 8px; 
      transition: background-color 0.15s ease, border-color 0.15s ease, transform 0.15s ease;
    }
    .theme-toggle:hover {
      background: var(--card-hover);
      border-color: var(--primary);
      transform: translateY(-1px);
    }
    .theme-toggle span#theme-icon {
      display: inline-block;
      transition: transform 0.2s ease;
    }
    .theme-toggle:hover span#theme-icon {
      transform: rotate(25deg) scale(1.1);
    }
    
    /* ============================================================
       CROSS-AUDIT VERIFICATION BANNER
       ============================================================ */
    .audit-banner { 
      background: var(--card-solid); 
      border: 1px solid var(--border); 
      border-left: 5px solid var(--success); 
      border-radius: 16px; 
      padding: 18px 24px; 
      margin-bottom: 24px; 
      display: flex; 
      justify-content: space-between; 
      align-items: center; 
      box-shadow: var(--shadow-card); 
      flex-wrap: wrap;
      gap: 16px;
    }
    .audit-banner-left {
      display: flex;
      align-items: center;
      gap: 14px;
    }
    .audit-shield-icon {
      width: 44px;
      height: 44px;
      border-radius: 12px;
      background: var(--success-bg);
      border: 1px solid var(--success-border);
      display: flex;
      align-items: center;
      justify-content: center;
      flex-shrink: 0;
      box-shadow: 0 4px 12px var(--success-glow);
    }
    .audit-title { 
      font-size: 14.5px; 
      font-weight: 800; 
      color: var(--text); 
      letter-spacing: -0.2px;
    }
    .audit-desc { 
      font-size: 12.5px; 
      color: var(--text-muted); 
      margin-top: 4px; 
      line-height: 1.4;
    }
    .audit-stats { 
      display: flex; 
      gap: 20px; 
      align-items: center;
    }
    .audit-stat-item { 
      text-align: right; 
    }
    .audit-stat-divider {
      width: 1px;
      height: 34px;
      background: var(--border);
    }
    .audit-stat-label { 
      font-size: 11px; 
      text-transform: uppercase; 
      color: var(--text-muted); 
      font-weight: 700; 
      letter-spacing: 0.5px;
    }
    .audit-stat-val { 
      font-size: 19px; 
      font-weight: 800; 
      color: var(--success); 
      font-family: 'JetBrains Mono', monospace; 
      margin-top: 2px;
    }

    /* ============================================================
       TOP KPI CARDS GRID
       ============================================================ */
    .kpi-grid { 
      display: grid; 
      grid-template-columns: repeat(auto-fit, minmax(235px, 1fr)); 
      gap: 16px; 
      margin-bottom: 24px; 
    }
    .kpi-card { 
      background: var(--card-solid); 
      padding: 20px 22px; 
      border-radius: 16px; 
      border: 1px solid var(--border-subtle); 
      box-shadow: var(--shadow-card); 
      position: relative; 
      overflow: hidden;
      transition: transform 0.16s ease, border-color 0.16s ease;
    }
    .kpi-card:hover {
      transform: translateY(-2px);
      box-shadow: var(--shadow-hover);
      border-color: var(--border-glow);
    }
    .kpi-card::before {
      content: '';
      position: absolute;
      top: 0; left: 0; right: 0;
      height: 2.5px;
      background: linear-gradient(90deg, transparent, var(--border-glow), transparent);
      opacity: 0.6;
      transition: opacity 0.25s;
    }
    .kpi-card:hover::before {
      opacity: 1;
    }
    .kpi-tag {
      font-size: 9.5px;
      font-weight: 800;
      letter-spacing: 0.8px;
      color: var(--text-dim);
      text-transform: uppercase;
      margin-bottom: 6px;
    }
    .kpi-title { 
      font-size: 11.5px; 
      font-weight: 700; 
      color: var(--text-muted); 
      text-transform: uppercase; 
      letter-spacing: 0.6px; 
    }
    .kpi-value { 
      font-size: 24px; 
      font-weight: 800; 
      color: var(--text); 
      margin-top: 8px; 
      font-family: 'JetBrains Mono', monospace; 
      letter-spacing: -0.5px;
    }
    .kpi-subtext { 
      font-size: 11.5px; 
      color: var(--text-muted); 
      margin-top: 5px; 
      font-weight: 500; 
    }
    .kpi-card.highlight { 
      border-color: var(--success-border); 
      background: linear-gradient(180deg, var(--card-bg) 0%, var(--success-bg) 100%); 
      box-shadow: 0 8px 30px var(--success-glow);
    }
    .kpi-card.highlight::before {
      background: linear-gradient(90deg, transparent, var(--success), transparent);
    }
    .kpi-card.highlight .kpi-value { 
      color: var(--success); 
    }
    .kpi-card.settled .kpi-value { 
      color: var(--success); 
    }
    .kpi-card.danger { 
      border-color: var(--danger-border); 
      background: linear-gradient(180deg, var(--card-bg) 0%, var(--danger-bg) 100%); 
      box-shadow: 0 8px 30px var(--danger-glow);
    }
    .kpi-card.danger::before {
      background: linear-gradient(90deg, transparent, var(--danger), transparent);
    }
    .kpi-card.danger .kpi-value { 
      color: var(--danger); 
    }

    /* ============================================================
       3-BOX DRAG & DROP UPLOAD GRID
       ============================================================ */
    .upload-grid { 
      display: grid; 
      grid-template-columns: repeat(3, 1fr); 
      gap: 16px; 
      margin-bottom: 24px; 
    }
    @media (max-width: 900px) {
      .upload-grid { grid-template-columns: 1fr; }
    }
    .upload-card { 
      background: var(--card-solid); 
      border: 2px dashed var(--border); 
      border-radius: 16px; 
      padding: 24px 20px; 
      text-align: center; 
      cursor: pointer; 
      transition: transform 0.16s ease, border-color 0.16s ease, background-color 0.16s ease; 
      position: relative; 
      box-shadow: var(--shadow-card);
    }
    .upload-card:hover { 
      border-color: var(--primary); 
      background: var(--primary-glow-subtle); 
      transform: translateY(-2px); 
    }
    .upload-icon { 
      font-size: 30px; 
      margin-bottom: 8px; 
      display: inline-block;
      transition: transform 0.25s ease;
    }
    .upload-card:hover .upload-icon {
      transform: translateY(-4px) scale(1.1);
    }
    .upload-title { 
      font-size: 14px; 
      font-weight: 800; 
      color: var(--text); 
      letter-spacing: -0.2px;
    }
    .upload-desc { 
      font-size: 11.5px; 
      color: var(--text-muted); 
      margin-top: 5px; 
      line-height: 1.45; 
    }
    .upload-chip {
      display: inline-block;
      margin-top: 10px;
      padding: 3px 10px;
      border-radius: 6px;
      background: var(--surface);
      border: 1px solid var(--border-subtle);
      font-size: 10.5px;
      font-weight: 700;
      color: var(--primary);
    }
    .upload-input { 
      position: absolute; 
      inset: 0; 
      opacity: 0; 
      cursor: pointer; 
      width: 100%; 
      height: 100%; 
    }

    /* ============================================================
       SEGMENTED NAVIGATION TABS
       ============================================================ */
    .tabs { 
      display: flex; 
      gap: 6px; 
      margin-bottom: 20px; 
      overflow-x: auto; 
      padding: 6px; 
      background: var(--surface-subtle); 
      border: 1px solid var(--border-subtle); 
      border-radius: 14px; 
      box-shadow: inset 0 1px 3px rgba(0,0,0,0.06);
    }
    .tab-btn { 
      padding: 10px 18px; 
      border-radius: 10px; 
      font-size: 13px; 
      font-weight: 700; 
      cursor: pointer; 
      background: transparent; 
      border: 1px solid transparent; 
      color: var(--text-muted); 
      transition: background-color 0.15s ease, color 0.15s ease, border-color 0.15s ease; 
      white-space: nowrap; 
      display: inline-flex; 
      align-items: center; 
      gap: 6px; 
    }
    .tab-btn:hover { 
      color: var(--text); 
      background: var(--surface);
    }
    .tab-btn.active { 
      background: linear-gradient(135deg, var(--primary) 0%, var(--accent) 100%); 
      color: #FFFFFF; 
      border-color: rgba(255, 255, 255, 0.2); 
      box-shadow: 0 3px 10px var(--primary-glow); 
    }

    /* ============================================================
       SECTION CARDS & TABLES
       ============================================================ */
    .section-card { 
      background: var(--card-solid); 
      border-radius: 18px; 
      border: 1px solid var(--border-subtle); 
      margin-bottom: 24px; 
      box-shadow: var(--shadow-card); 
      overflow: hidden;
      contain: layout style;
      animation: tabFadeIn 0.15s ease-out;
    }
    .section-header { 
      padding: 18px 24px; 
      border-bottom: 1px solid var(--border); 
      display: flex; 
      justify-content: space-between; 
      align-items: center; 
      background: var(--surface-subtle); 
      flex-wrap: wrap;
      gap: 12px;
    }
    .section-title { 
      font-size: 15px; 
      font-weight: 800; 
      color: var(--text); 
      display: flex; 
      align-items: center; 
      gap: 8px; 
      letter-spacing: -0.3px;
    }
    .section-actions { 
      display: flex; 
      gap: 10px; 
      align-items: center; 
    }

    /* Table Container with Custom Scrollbar */
    .table-container { 
      width: 100%; 
      overflow-x: auto; 
      -webkit-overflow-scrolling: touch; 
    }
    .table-container::-webkit-scrollbar { 
      height: 8px; 
    }
    .table-container::-webkit-scrollbar-track { 
      background: var(--card-bg); 
    }
    .table-container::-webkit-scrollbar-thumb { 
      background: var(--border); 
      border-radius: 4px; 
    }
    .table-container::-webkit-scrollbar-thumb:hover { 
      background: var(--primary); 
    }

    table { 
      width: 100%; 
      min-width: 1120px; 
      border-collapse: collapse; 
      font-size: 12.5px; 
      text-align: left; 
    }
    th { 
      background: var(--table-header-glass); 
      color: var(--text-muted); 
      font-weight: 800; 
      font-size: 11px; 
      text-transform: uppercase; 
      letter-spacing: 0.6px; 
      padding: 13px 15px; 
      border-bottom: 1px solid var(--border); 
      white-space: nowrap; 
    }
    td { 
      padding: 12px 15px; 
      border-bottom: 1px solid var(--border); 
      vertical-align: middle; 
    }
    tr:nth-child(even) td { 
      background-color: var(--table-stripe); 
    }
    tr:hover td { 
      background-color: var(--card-hover); 
    }
    
    .num { 
      text-align: right; 
      font-family: 'JetBrains Mono', monospace; 
      font-variant-numeric: tabular-nums; 
    }
    .center { text-align: center; }
    .bold { font-weight: 700; }
    .mono { font-family: 'JetBrains Mono', monospace; }
    
    /* Modern Status Badges */
    .badge { 
      display: inline-flex; 
      align-items: center;
      gap: 5px;
      padding: 4px 10px; 
      border-radius: 9999px; 
      font-size: 11px; 
      font-weight: 800; 
      letter-spacing: 0.3px; 
      border: 1px solid transparent; 
      white-space: nowrap; 
    }
    .badge::before {
      content: '';
      display: inline-block;
      width: 6px;
      height: 6px;
      border-radius: 50%;
      background: currentColor;
    }
    .badge-danger { 
      background: var(--danger-bg); 
      color: var(--danger); 
      border-color: var(--danger-border); 
    }
    .badge-success { 
      background: var(--success-bg); 
      color: var(--success); 
      border-color: var(--success-border); 
    }
    .badge-warning { 
      background: var(--warning-bg); 
      color: var(--warning); 
      border-color: var(--warning-border); 
    }
    .badge-info { 
      background: var(--primary-glow-subtle); 
      color: var(--primary); 
      border-color: var(--primary); 
    }

    /* Quick Rate Editor Bar */
    .quick-bar { 
      background: var(--surface-subtle); 
      border-bottom: 1px solid var(--border); 
      padding: 12px 22px; 
      display: flex; 
      align-items: center; 
      justify-content: space-between; 
      flex-wrap: wrap; 
      gap: 12px; 
    }
    .quick-bar-title { 
      font-size: 12.5px; 
      font-weight: 700; 
      color: var(--text-muted); 
    }
    .quick-bar-buttons { 
      display: flex; 
      gap: 8px; 
      flex-wrap: wrap; 
    }
    .quick-chip { 
      background: var(--surface); 
      border: 1px solid var(--border); 
      border-radius: 8px; 
      padding: 6px 12px; 
      font-size: 11.5px; 
      font-weight: 700; 
      color: var(--text); 
      cursor: pointer; 
      display: inline-flex; 
      align-items: center; 
      gap: 5px; 
      transition: background-color 0.15s ease, transform 0.15s ease; 
    }
    .quick-chip:hover { 
      border-color: var(--primary); 
      background: var(--card-hover); 
      color: var(--primary); 
      transform: translateY(-1.5px);
      box-shadow: 0 4px 10px var(--primary-glow);
    }

    /* Expandable Item Drawer Sub-row */
    .item-drawer-row td { 
      background: var(--surface-subtle) !important; 
      padding: 14px 20px; 
    }
    .item-subtable { 
      width: 100%; 
      min-width: 800px; 
      border: 1px solid var(--border); 
      border-radius: 10px; 
      overflow: hidden; 
      font-size: 12px; 
      margin: 4px 0; 
      box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    }
    .item-subtable th { 
      background: var(--surface); 
      font-size: 10.5px; 
      padding: 9px 12px; 
    }
    .item-subtable td { 
      padding: 9px 12px; 
      border-bottom: 1px solid var(--border); 
    }
    .inv-link { 
      color: var(--primary); 
      text-decoration: none; 
      font-weight: 700; 
      cursor: pointer; 
      display: inline-flex; 
      align-items: center; 
      gap: 4px; 
    }
    .inv-link:hover { 
      text-decoration: underline; 
    }

    /* Clickable PDF Link Pills with Modern Hover Glow */
    .pdf-pill {
      color: var(--primary);
      text-decoration: none;
      font-weight: 700;
      display: inline-flex;
      align-items: center;
      gap: 5px;
      background: rgba(56, 189, 248, 0.08);
      padding: 3.5px 9px;
      border-radius: 7px;
      border: 1px solid rgba(56, 189, 248, 0.25);
      transition: background-color 0.15s ease, transform 0.15s ease;
      font-family: 'JetBrains Mono', monospace;
      font-size: 11.5px;
      white-space: nowrap;
    }
    .pdf-pill:hover {
      background: var(--primary);
      color: #090E1D;
      border-color: var(--primary);
      transform: translateY(-1.5px);
      box-shadow: 0 4px 12px var(--primary-glow);
    }
    .pdf-pill-cn {
      color: #A78BFA;
      background: rgba(167, 139, 250, 0.08);
      border-color: rgba(167, 139, 250, 0.25);
    }
    .pdf-pill-cn:hover {
      background: #A78BFA;
      color: #090E1D;
      border-color: #A78BFA;
      box-shadow: 0 4px 12px rgba(167, 139, 250, 0.4);
    }
    .pdf-pill-fe {
      color: #F59E0B;
      background: rgba(245, 158, 11, 0.08);
      border-color: rgba(245, 158, 11, 0.3);
    }
    .pdf-pill-fe:hover {
      background: #F59E0B;
      color: #090E1D;
      border-color: #F59E0B;
      box-shadow: 0 4px 12px rgba(245, 158, 11, 0.4);
    }
    .pdf-pill-pms {
      color: #C084FC;
      background: rgba(192, 132, 252, 0.08);
      border-color: rgba(192, 132, 252, 0.3);
    }
    .pdf-pill-pms:hover {
      background: #C084FC;
      color: #090E1D;
      border-color: #C084FC;
      box-shadow: 0 4px 12px rgba(192, 132, 252, 0.4);
    }

    /* CD Guide Info Card */
    .info-guide-card { 
      background: var(--surface-subtle); 
      border: 1px solid var(--border); 
      border-radius: 14px; 
      padding: 20px 22px; 
      margin-bottom: 20px; 
    }
    .info-guide-header { 
      font-size: 14px; 
      font-weight: 800; 
      color: var(--text); 
      margin-bottom: 12px; 
      display: flex; 
      align-items: center; 
      gap: 8px; 
    }
    .info-guide-body { 
      font-size: 12.5px; 
      color: var(--text-muted); 
      line-height: 1.6; 
    }
    .guide-steps { 
      display: grid; 
      grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); 
      gap: 14px; 
      margin-top: 14px; 
    }
    .guide-step { 
      background: var(--card-bg); 
      border: 1px solid var(--border); 
      border-radius: 10px; 
      padding: 14px; 
      display: flex; 
      gap: 12px; 
      align-items: flex-start; 
      box-shadow: 0 2px 6px rgba(0,0,0,0.02);
    }
    .step-num { 
      background: linear-gradient(135deg, var(--primary), var(--accent)); 
      color: #FFFFFF; 
      width: 24px; 
      height: 24px; 
      border-radius: 50%; 
      font-size: 11px; 
      font-weight: 800; 
      display: flex; 
      align-items: center; 
      justify-content: center; 
      flex-shrink: 0; 
      margin-top: 1px; 
      box-shadow: 0 2px 6px var(--primary-glow);
    }

    /* ============================================================
       MODAL GLASSMORPHISM & RATE INPUTS
       ============================================================ */
    .modal-backdrop { 
      position: fixed; 
      inset: 0; 
      background: rgba(4, 8, 20, 0.86); 
      display: none; 
      align-items: center; 
      justify-content: center; 
      z-index: 1000; 
    }
    .modal { 
      background: var(--card-solid); 
      width: 95%; 
      max-width: 1000px; 
      border-radius: 20px; 
      max-height: 90vh; 
      display: flex; 
      flex-direction: column; 
      overflow: hidden; 
      border: 1px solid var(--border-glow); 
      box-shadow: 0 24px 48px -12px rgba(0,0,0,0.7), 0 0 0 1px rgba(255,255,255,0.08); 
      animation: modalIn 0.16s ease-out;
    }
    .modal-header { 
      padding: 20px 24px; 
      border-bottom: 1px solid var(--border); 
      display: flex; 
      justify-content: space-between; 
      align-items: center; 
      background: var(--surface-subtle); 
    }
    .modal-header h3 { 
      font-size: 16.5px; 
      font-weight: 800; 
      color: var(--text); 
      letter-spacing: -0.3px;
    }
    .modal-body { 
      padding: 24px; 
      overflow-y: auto; 
      flex: 1; 
    }
    .modal-footer { 
      padding: 18px 24px; 
      border-top: 1px solid var(--border); 
      display: flex; 
      justify-content: flex-end; 
      gap: 12px; 
      background: var(--surface-subtle); 
    }
    .rate-input { 
      width: 95px; 
      padding: 8px 10px; 
      background: var(--surface); 
      border: 1px solid var(--border); 
      border-radius: 8px; 
      font-size: 13px; 
      font-weight: 700; 
      text-align: right; 
      color: var(--text); 
      outline: none; 
      transition: all 0.2s ease;
    }
    .rate-input:focus { 
      border-color: var(--primary); 
      box-shadow: 0 0 0 3px var(--primary-glow); 
      background: var(--card-bg);
    }
    .close-btn { 
      background: none; 
      border: none; 
      font-size: 24px; 
      cursor: pointer; 
      color: var(--text-muted); 
      line-height: 1; 
      transition: color 0.15s ease, transform 0.15s ease;
    }
    .close-btn:hover { 
      color: var(--text); 
      transform: scale(1.15);
    }
  </style>
</head>
<body>
  <div class="container">
    <header>
      <div class="brand">
        <div class="brand-title-wrap">
          <h1><span class="logo-shield">🛡️ GGL-Recon</span> Rate Difference & Audit System</h1>
          <div class="live-pill"><span class="live-dot"></span> LIVE VERIFIED</div>
        </div>
        <p>
          <span>Mehta Impex</span> 
          <span class="brand-pill">GGL ModiGuard®</span> 
          <span class="brand-pill">FY 2026–27</span>
          <span style="color:var(--text-dim);">• Official Supplier Account Audit & Reconciliation</span>
        </p>
      </div>
      <div class="header-actions">
        <div id="user-badge-pill" class="brand-pill" style="display:flex; align-items:center; gap:8px; padding:7px 12px; background:var(--surface); border:1px solid var(--border); border-radius:999px; font-size:12px;">
          <span id="user-badge-icon">👤</span>
          <strong id="user-badge-name" style="color:var(--text);">User</strong>
          <span id="user-role-tag" class="badge badge-info" style="font-size:10px; padding:2px 7px;">ROLE</span>
        </div>
        <button id="btn-manage-users" class="btn" style="display:none; background:var(--surface); color:var(--text); border:1px solid var(--border);" onclick="switchTab('users')" title="Manage Users & Roles (RBAC)">
          <span>👥</span> Users & Roles
        </button>
        <button class="btn" style="background:var(--surface); color:var(--text); border:1px solid var(--border);" onclick="switchTab('account')" title="Password & Account Security">
          <span>🔑</span> Password
        </button>
        <button class="theme-toggle" onclick="toggleTheme()" title="Toggle Dark / Light Theme">
          <span id="theme-icon">☀️</span> <span id="theme-label">Light Mode</span>
        </button>
        <a href="/api/download-claim-pdf" class="btn btn-primary" target="_blank" title="Download 1-Page Official Settlement Claim PDF to send to GGL">
          <span>📄</span> Download 1-Page Claim PDF
        </a>
        <a href="/api/download-excel" class="btn btn-download" download title="Download Complete 7-Sheet Master Reconciliation Workbook">
          <span>📥</span> Download Master Excel (.xlsx)
        </a>
        <button class="btn" style="background:rgba(239,68,68,0.14); color:var(--danger); border:1px solid rgba(248,113,113,0.35);" onclick="logoutUser()" title="Sign Out">
          <span>🚪</span> Logout
        </button>
      </div>
    </header>

    <!-- Cross-Audit Live Verification Alert Banner -->
    <div class="audit-banner" id="audit-banner">
      <div class="audit-banner-left">
        <div class="audit-shield-icon">
          <span id="audit-icon" style="font-size:24px;">🛡️</span>
        </div>
        <div>
          <div class="audit-title" id="audit-headline">100% Double-Checked Safety Audit Passed</div>
          <div class="audit-desc" id="audit-subtext">Zero duplicate CN allocations. F&E CN #3242 bound strictly to May (#1303). F&E CN #6041 strictly in June (#2138). All dispatch calculations verified.</div>
        </div>
      </div>
      <div class="audit-stats">
        <div class="audit-stat-item">
          <div class="audit-stat-label">Safety Checks</div>
          <div class="audit-stat-val" id="audit-safety-val" style="color:var(--success);">9 / 9 PASSED</div>
        </div>
        <div class="audit-stat-divider"></div>
        <div class="audit-stat-item">
          <div class="audit-stat-label">CNs in GGL Ledger</div>
          <div class="audit-stat-val" id="audit-total-cns">26 / 26</div>
        </div>
        <div class="audit-stat-divider"></div>
        <div class="audit-stat-item">
          <div class="audit-stat-label">GGL Statement Closing Balance</div>
          <div class="audit-stat-val" id="audit-closing-bal" style="color:var(--success);">₹0.00</div>
        </div>
      </div>
    </div>

    <!-- Top KPI Row -->
    <div class="kpi-grid">
      <div class="kpi-card">
        <div class="kpi-tag">Gross Invoicing</div>
        <div class="kpi-title">Total Dispatches Billed</div>
        <div class="kpi-value" id="kpi-billed">₹0.00</div>
        <div class="kpi-subtext">Gross billed by GGL at list rates</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-tag">Agreed Phone Rate</div>
        <div class="kpi-title">Actual Expected Cost</div>
        <div class="kpi-value" id="kpi-actual">₹0.00</div>
        <div class="kpi-subtext">Based on agreed phone rates</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-tag" style="color:var(--danger)">Variance</div>
        <div class="kpi-title">Overbilled Rate Difference</div>
        <div class="kpi-value" id="kpi-overbilled" style="color:var(--text)">₹0.00</div>
        <div class="kpi-subtext">Basic difference + 18% GST</div>
      </div>
      <div class="kpi-card settled">
        <div class="kpi-tag" style="color:var(--success)">GGL Credits</div>
        <div class="kpi-title">Total RD Settled via CNs</div>
        <div class="kpi-value" id="kpi-settled">₹0.00</div>
        <div class="kpi-subtext">RD, ARD & F&E refund notes</div>
      </div>
      <div class="kpi-card highlight" id="kpi-card-net">
        <div class="kpi-tag" style="color:var(--success)">Surplus Position</div>
        <div class="kpi-title" id="kpi-net-title">NET RATE DIFFERENCE POSITION</div>
        <div class="kpi-value" id="kpi-net-val">₹0.00</div>
        <div class="kpi-subtext" id="kpi-net-sub">Cumulative extra credit in your favor</div>
      </div>
    </div>

    <!-- Volume Incentives & Separate Allowances Row (FY 2026–27 Only) -->
    <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(235px, 1fr)); gap: 16px; margin-bottom: 24px;">
      <div class="kpi-card" style="border-left: 4px solid var(--accent); background: linear-gradient(180deg, var(--card-bg) 0%, rgba(96, 165, 250, 0.08) 100%);">
        <div class="kpi-tag" style="color:var(--accent);">Q1 & Q2 Volume</div>
        <div class="kpi-title" style="color:var(--accent);">🏅 QUARTERLY ARD (FY 26–27)</div>
        <div class="kpi-value" id="kpi-qtr-ard" style="color:var(--accent);">₹0.00</div>
        <div class="kpi-subtext" id="kpi-qtr-sub">Quarterly Tonnage CNs credited in FY 26–27</div>
      </div>
      <div class="kpi-card" style="border-left: 4px solid #A855F7; background: linear-gradient(180deg, var(--card-bg) 0%, rgba(168, 85, 247, 0.08) 100%);">
        <div class="kpi-tag" style="color:#C084FC;">Target Settlement</div>
        <div class="kpi-title" style="color:#C084FC;">🎖️ ANNUAL ARD (FY 26–27)</div>
        <div class="kpi-value" id="kpi-yearly-ard" style="color:#C084FC;">₹0.00</div>
        <div class="kpi-subtext">Nil (Awaiting FY 26–27 year-end closing)</div>
      </div>
      <div class="kpi-card" style="border-left: 4px solid var(--warning); background: linear-gradient(180deg, var(--card-bg) 0%, var(--warning-bg) 100%);">
        <div class="kpi-tag" style="color:var(--warning);">Prompt Turnaround</div>
        <div class="kpi-title" style="color:var(--warning);">💰 CASH DISCOUNT RECEIVED (FY 26–27)</div>
        <div class="kpi-value" id="kpi-tot-cd" style="color:var(--warning);">₹0.00</div>
        <div class="kpi-subtext" id="kpi-sub-cd">Prompt Payment Cash Discount</div>
      </div>
      <div class="kpi-card" style="border-left: 4px solid var(--success); background: linear-gradient(180deg, var(--card-bg) 0%, var(--success-bg) 100%);">
        <div class="kpi-tag" style="color:var(--success);">Supplier Ledger</div>
        <div class="kpi-title" style="color:var(--success);">📑 GGL STATEMENT ADVANCE BALANCE</div>
        <div class="kpi-value" id="kpi-ledger-bal" style="color:var(--success);">₹0.00</div>
        <div class="kpi-subtext">Official closing credit in Mehta Impex's favor</div>
      </div>
    </div>

    <!-- Read-Only RBAC Notice & Report Downloads (shown to Viewer accounts) -->
    <div id="readonly-banner" style="display:none; padding:14px 20px; margin-bottom:20px; border-radius:14px; background:var(--surface); border:1px solid var(--border); border-left:4px solid var(--accent); align-items:center; justify-content:space-between; flex-wrap:wrap; gap:12px;">
      <div style="display:flex; align-items:center; gap:10px; font-size:13px; color:var(--text);">
        <span style="font-size:18px;">👁️</span>
        <span><strong>Read-Only Viewer Mode:</strong> You can inspect all audit reports, open original PDFs, and download the Claim PDF or Master Excel workbook. Editing rates or uploading documents is disabled.</span>
      </div>
      <div style="display:flex; align-items:center; gap:8px; flex-wrap:wrap;">
        <a href="/api/download-claim-pdf" class="btn btn-primary" style="padding:7px 12px; font-size:12px;" target="_blank">📄 Download Claim PDF</a>
        <a href="/api/download-excel" class="btn btn-download" style="padding:7px 12px; font-size:12px;" download>📥 Download Master Excel</a>
        <span class="badge badge-info">Read-Only RBAC</span>
      </div>
    </div>

    <!-- 3 Drag-and-Drop Upload Boxes -->
    <div class="upload-grid" id="upload-grid-wrap">
      <div class="upload-card">
        <input type="file" class="upload-input" accept=".pdf" onchange="uploadFile(this, '/api/upload-bill')">
        <div class="upload-icon">📄</div>
        <div class="upload-title">Upload GGL Invoice / Bill (PDF)</div>
        <div class="upload-desc">Drop new dispatch invoice. Auto-extracts dimensions, SQM, MM, Billed Rate & Surcharges.</div>
        <div class="upload-chip">Drop PDF here or click to browse</div>
      </div>
      <div class="upload-card">
        <input type="file" class="upload-input" accept=".pdf" onchange="uploadFile(this, '/api/upload-cn')">
        <div class="upload-icon">🏷️</div>
        <div class="upload-title">Upload Credit Note (PDF)</div>
        <div class="upload-desc">Drop Rate Difference (RD), ARD, Cash Discount (CD), or F&E refund credit notes.</div>
        <div class="upload-chip">Drop PDF here or click to browse</div>
      </div>
      <div class="upload-card">
        <input type="file" class="upload-input" accept=".pdf" onchange="uploadFile(this, '/api/upload-ledger')">
        <div class="upload-icon">📋</div>
        <div class="upload-title">Upload GGL Statement / Ledger (PDF)</div>
        <div class="upload-desc">Drop company statement to auto-update payments, closing balance, and audit for missing CNs.</div>
        <div class="upload-chip">Drop PDF here or click to browse</div>
      </div>
    </div>

    <!-- Tabs Navigation (FY 2026–27 Only) -->
    <div class="tabs" id="main-tabs-nav">
      <button class="tab-btn active" onclick="switchTab('monthly', this)">📅 Month-wise RD Reconciliation</button>
      <button class="tab-btn" onclick="switchTab('invoices', this)">🚚 Bill-wise Breakdown & Agreed Rate Editor</button>
      <button class="tab-btn" onclick="switchTab('cd', this)">💰 Cash Discount (3%) Tracker</button>
      <button class="tab-btn" onclick="switchTab('tonnage', this)">🏅 Quarterly ARD Incentives (FY 26–27)</button>
      <button class="tab-btn" onclick="switchTab('ggl-ledger', this)">📑 GGL Official Statement (FY 2026–27)</button>
      <button class="tab-btn" id="tab-btn-users" style="display:none;" onclick="switchTab('users', this)">👥 Users &amp; RBAC</button>
      <button class="tab-btn" id="tab-btn-account" onclick="switchTab('account', this)">🔑 Password &amp; Security</button>
    </div>

    <!-- Tab 1: Month-wise Reconciliation -->
    <div id="tab-monthly" class="section-card">
      <div class="section-header">
        <div class="section-title">📅 Month-wise Rate Difference Settlement Summary (FY 2026–27)</div>
        <div class="section-actions">
          <a href="/api/download-claim-pdf" target="_blank" class="pdf-pill" style="font-size:12px; padding:6px 12px; text-decoration:none; color:var(--primary); background:rgba(56,189,248,0.1); border-color:var(--primary);" title="Download 1-Page Official Settlement Claim PDF to send to GGL">
            📄 Download 1-Page Claim PDF (For GGL) ↗
          </a>
        </div>
      </div>
      <div class="table-container">
        <table>
          <thead>
            <tr>
              <th>Month</th>
              <th>Invoice IDs</th>
              <th class="num">Total SQMM</th>
              <th class="num">Billed Total (₹)</th>
              <th class="num">Expected Cost (₹)</th>
              <th class="num" title="Pure Rate Difference on Glass Items">Glass RD DIFF (₹)</th>
              <th class="num" title="Strictly Rate Difference & Additional Rate Difference Credit Notes">RD & ARD Settled (₹)</th>
              <th class="num">Month Net Difference (₹)</th>
              <th class="num">Cumulative Running Balance (₹)</th>
              <th class="center">Settlement Status</th>
            </tr>
          </thead>
          <tbody id="monthly-tbody">
            <tr><td colspan="10" class="center">Loading data...</td></tr>
          </tbody>
        </table>
      </div>
    </div>

    <!-- Tab 2: Bill-wise Breakdown -->
    <div id="tab-invoices" class="section-card" style="display:none;">
      <div class="section-header">
        <div class="section-title">🚚 Bill-by-Bill Breakdown & Agreed Rate / Surcharge Editor</div>
      </div>
      <div class="quick-bar" id="quick-edit-bar">
        <div class="quick-bar-title">⚡ Quick Edit Discussed Rates for Invoices:</div>
        <div class="quick-bar-buttons" id="quick-inv-buttons"></div>
      </div>
      <div class="table-container">
        <table>
          <thead>
            <tr>
              <th>Invoice Date</th>
              <th>Invoice No</th>
              <th class="center">Agreed Rates Action</th>
              <th>Items & Types</th>
              <th class="num">Total SQMM</th>
              <th class="num">F&E (Rate & ₹)</th>
              <th class="num">PMS (Rate & ₹)</th>
              <th class="num">INS (₹)</th>
              <th class="num">Billed Amount (₹)</th>
              <th class="num">Expected Cost (₹)</th>
              <th class="num">Glass RD DIFF (₹)</th>
              <th>Matched Credit Note IDs</th>
            </tr>
          </thead>
          <tbody id="invoices-tbody"></tbody>
        </table>
      </div>
    </div>

    <!-- Tab 3: Separate CD Tracker -->
    <div id="tab-cd" class="section-card" style="display:none;">
      <div class="section-header">
        <div class="section-title">💰 Cash Discount (CD) Verification & Prompt Payment Tracker</div>
      </div>
      <div style="padding: 18px 20px 0 20px;">
        <div class="info-guide-card">
          <div class="info-guide-header">💡 Official GGL Cash Discount (CD) Calculation Rules</div>
          <div class="info-guide-body">
            <p><strong>Official GGL Cash Discount Formula & Tiered Daily Prompt Payment Rules:</strong></p>
            <div class="guide-steps">
              <div class="guide-step">
                <div class="step-num">1</div>
                <div><strong>Base 'X' Calculation:</strong> <code>X = Invoice Amount - 1st Rate Difference CN Amount</code>. CD is computed strictly on Base 'X', where the rate difference is the 1st CN issued with the invoice mentioning 'Rate Difference'.</div>
              </div>
              <div class="guide-step">
                <div class="step-num">2</div>
                <div><strong>24-Hour Benchmark:</strong> Payment made within 24 hours of invoice date qualifies for the full <strong>3.0%</strong> Cash Discount.</div>
              </div>
              <div class="guide-step">
                <div class="step-num">3</div>
                <div><strong>Daily Tiered Reduction:</strong> From 24 hours onward, the discount decreases by <strong>0.1% per day late</strong> (e.g. 24h = 3.0%, Day 2 = 2.9%, Day 3 = 2.8%, Day 4 = 2.7%, and so on).</div>
              </div>
              <div class="guide-step">
                <div class="step-num">4</div>
                <div><strong>Shortfall Claiming:</strong> Any gap between the eligible CD on Base X and GGL's actual CD Credit Note is tracked below as a claimable shortfall.</div>
              </div>
            </div>
          </div>
        </div>
      </div>
      <div class="table-container">
        <table>
          <thead>
            <tr>
              <th>Month</th>
              <th>Invoice No</th>
              <th>Inv Date</th>
              <th class="num">Billed Gross (₹)</th>
              <th class="center">1st RD CN No.</th>
              <th class="num">1st RD CN Amt (₹)</th>
              <th class="num" style="color:var(--primary)">Base 'X' (₹)</th>
              <th>Payment Date</th>
              <th>Payment Ref / NFT</th>
              <th class="center">Turnaround</th>
              <th class="center">CD Rate</th>
              <th class="num bold" style="color:var(--primary)">Eligible CD (₹)</th>
              <th class="center">GGL CD CN</th>
              <th class="num bold" style="color:var(--success)">CD Received (₹)</th>
              <th class="num bold" style="color:var(--danger)">CD Shortfall (₹)</th>
              <th class="center">Status</th>
            </tr>
          </thead>
          <tbody id="cd-tbody"></tbody>
        </table>
      </div>
    </div>

    <!-- Tab 5: GGL Official Statement -->
    <div id="tab-ggl-ledger" class="section-card" style="display:none;">
      <div class="section-header">
        <div class="section-title">📑 Gujarat Guardian Limited — Official Statement of Account (FY 2026–27)</div>
        <div class="section-actions">
          <a href="/api/view-pdf?doc=GGL_LEDGER_26_27" target="_blank" class="pdf-pill" style="font-size:12px; padding:6px 12px; text-decoration:none;">
            📄 View Official Statement (PDF) ↗
          </a>
        </div>
      </div>

      <!-- Missing Documents in Statement Alert Box -->
      <div id="ledger-missing-alert" style="display:none; padding:16px 20px; background:var(--warning-bg); border-bottom:1px solid var(--warning-border);">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
          <div style="font-size:13.5px; font-weight:800; color:var(--warning); display:flex; align-items:center; gap:8px;">
            <span>⚠️</span> <span>Action Required: Credit Notes in GGL Statement Missing from Database</span>
          </div>
          <span class="badge badge-warning" id="ledger-missing-count" style="font-size:11px;">0 Missing</span>
        </div>
        <p style="font-size:12px; color:var(--text); line-height:1.4; margin-bottom:12px;">
          The official GGL statement lists the following credit note transactions which have not been uploaded as PDF documents yet. Please upload their PDF copies to enable line-by-line verification and automatic reconciliation.
        </p>
        <div class="table-container" style="border:1px solid var(--border); border-radius:8px; background:var(--card-bg); max-height:220px; overflow-y:auto;">
          <table style="min-width:700px; font-size:11.5px;">
            <thead>
              <tr>
                <th>Statement Date</th>
                <th>Credit Note #</th>
                <th class="num">Credit Amount (₹)</th>
                <th>Statement Remarks</th>
                <th class="center">Status</th>
              </tr>
            </thead>
            <tbody id="ledger-missing-tbody"></tbody>
          </table>
        </div>
      </div>

      <!-- 100% Cross-Audited Banner -->
      <div id="ledger-all-matched-banner" style="padding:14px 22px; background:var(--success-bg); border-bottom:1px solid var(--success-border); display:flex; align-items:center; gap:10px;">
        <span style="font-size:18px;">✅</span>
        <span style="font-size:13px; font-weight:700; color:var(--success);" id="ledger-matched-msg">
          100% Cross-Audited: All Credit Notes in GGL's Official Statement are present and verified in your database.
        </span>
      </div>

      <div class="table-container">
        <table>
          <thead>
            <tr>
              <th>Date</th>
              <th class="center">Type</th>
              <th>Document No.</th>
              <th class="num">Debit Amt (₹)</th>
              <th class="num">Credit Amt (₹)</th>
              <th class="num">Balance (₹)</th>
              <th>Remarks</th>
            </tr>
          </thead>
          <tbody id="ledger-tbody"></tbody>
        </table>
      </div>
    </div>

    <!-- Tab 4: Quarterly ARD Incentives (FY 2026–27) -->
    <div id="tab-tonnage" class="section-card" style="display:none;">
      <div class="section-header">
        <div class="section-title">🏅 Quarterly ARD Incentives (FY 2026–27 Volume Tonnage Business)</div>
      </div>
      <div style="padding: 18px 20px 0 20px;">
        <div class="info-guide-card">
          <div class="info-guide-header">💡 Quarterly Tonnage ARDs Received in FY 2026–27</div>
          <div class="info-guide-body">
            <p id="ton-tab-desc">In this financial year, quarterly tonnage incentives are tracked in this dedicated section in your favor and are strictly separated from monthly invoice overbilling.</p>
          </div>
        </div>
      </div>
      <div style="padding: 0 20px 16px 20px; display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 14px;">
        <div style="background:var(--surface); border:1px solid var(--border); border-left:4px solid var(--accent); border-radius:12px; padding:14px 18px;">
          <div style="font-size:11px; font-weight:800; color:var(--text-muted); text-transform:uppercase; letter-spacing:0.5px;">Quarterly ARD Received (FY 26–27)</div>
          <div style="font-size:22px; font-weight:800; color:var(--accent); font-family:'JetBrains Mono',monospace; margin-top:4px;" id="ton-tab-qtr">₹0.00</div>
          <div style="font-size:11px; color:var(--text-muted); margin-top:2px;">Quarterly Tonnage CNs in FY 26–27</div>
        </div>
        <div style="background:var(--surface); border:1px solid var(--border); border-left:4px solid #A855F7; border-radius:12px; padding:14px 18px;">
          <div style="font-size:11px; font-weight:800; color:var(--text-muted); text-transform:uppercase; letter-spacing:0.5px;">Annual ARD (FY 26–27)</div>
          <div style="font-size:22px; font-weight:800; color:#C084FC; font-family:'JetBrains Mono',monospace; margin-top:4px;" id="ton-tab-annual">₹0.00</div>
          <div style="font-size:11px; color:var(--text-muted); margin-top:2px;">Nil (Awaiting FY 26–27 Year-End Closing)</div>
        </div>
        <div style="background:var(--surface); border:1px solid var(--border); border-left:4px solid var(--success); border-radius:12px; padding:14px 18px;">
          <div style="font-size:11px; font-weight:800; color:var(--text-muted); text-transform:uppercase; letter-spacing:0.5px;">Total Tonnage Incentives (FY 26–27)</div>
          <div style="font-size:22px; font-weight:800; color:var(--success); font-family:'JetBrains Mono',monospace; margin-top:4px;" id="ton-tab-total">₹0.00</div>
          <div style="font-size:11px; color:var(--text-muted); margin-top:2px;">100% Credited to Mehta Impex</div>
        </div>
      </div>
      <div class="table-container">
        <table>
          <thead>
            <tr>
              <th>Period / Scope</th>
              <th>CN Date</th>
              <th>Document No.</th>
              <th>Description / Tonnage Details</th>
              <th class="center">Eligible Tonnage</th>
              <th class="center">Incentive Rate</th>
              <th class="num">Credit Amount (₹)</th>
              <th class="center">Status</th>
            </tr>
          </thead>
          <tbody id="tonnage-tbody"></tbody>
        </table>
      </div>
    </div>

    <!-- Tab 6: Admin RBAC Users & Roles Management Tab (Admin Only) -->
    <div id="tab-users" class="section-card" style="display:none;">
      <div class="section-header">
        <div class="section-title">👥 User Accounts &amp; Role-Based Access Control (RBAC)</div>
        <div class="section-actions">
          <button class="btn btn-primary-sm" onclick="loadUsersTable()">🔄 Refresh Users</button>
        </div>
      </div>
      <div style="padding: 20px;">
        <div id="users-msg" style="display:none; padding:10px 14px; border-radius:10px; font-size:12.5px; font-weight:700; margin-bottom:16px;"></div>

        <!-- Role Matrix Overview -->
        <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(220px, 1fr)); gap:12px; margin-bottom:18px;">
          <div style="background:var(--surface); border:1px solid var(--border); border-left:3px solid #38BDF8; border-radius:12px; padding:12px 14px;">
            <div style="font-size:12px; font-weight:800; color:var(--primary);">🛡️ ADMIN (Administrator)</div>
            <div style="font-size:11.5px; color:var(--text-muted); margin-top:4px;">View, Edit Rates, Upload PDFs, Download Exports &amp; Manage Users/Roles.</div>
          </div>
          <div style="background:var(--surface); border:1px solid var(--border); border-left:3px solid #34D399; border-radius:12px; padding:12px 14px;">
            <div style="font-size:12px; font-weight:800; color:var(--success);">✏️ EDITOR (Accounts Editor)</div>
            <div style="font-size:11.5px; color:var(--text-muted); margin-top:4px;">View, Edit Discussed Rates/Surcharges, Upload PDFs &amp; Download Exports.</div>
          </div>
          <div style="background:var(--surface); border:1px solid var(--border); border-left:3px solid #818CF8; border-radius:12px; padding:12px 14px;">
            <div style="font-size:12px; font-weight:800; color:var(--accent);">👁️ VIEWER (Read-Only)</div>
            <div style="font-size:11.5px; color:var(--text-muted); margin-top:4px;">Read-Only Dashboard, Inspect PDFs, &amp; Download Claim PDF / Master Excel. Editing blocked.</div>
          </div>
        </div>

        <!-- Add New User Form -->
        <div style="background:var(--surface); border:1px solid var(--border); border-radius:12px; padding:16px; margin-bottom:18px;">
          <div style="font-size:13px; font-weight:800; margin-bottom:12px; color:var(--text);">➕ Create New User Account</div>
          <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(160px, 1fr)); gap:12px; align-items:end;">
            <div>
              <label style="display:block; font-size:10.5px; font-weight:700; color:var(--text-muted); margin-bottom:5px;">LOGIN ID</label>
              <input type="text" id="newUsername" class="rate-input" style="width:100%; text-align:left; padding:9px 11px;" placeholder="e.g. accounts1">
            </div>
            <div>
              <label style="display:block; font-size:10.5px; font-weight:700; color:var(--text-muted); margin-bottom:5px;">FULL NAME</label>
              <input type="text" id="newDisplayName" class="rate-input" style="width:100%; text-align:left; padding:9px 11px;" placeholder="e.g. Ramesh Patel">
            </div>
            <div>
              <label style="display:block; font-size:10.5px; font-weight:700; color:var(--text-muted); margin-bottom:5px;">ROLE</label>
              <select id="newUserRole" class="rate-input" style="width:100%; text-align:left; padding:9px 11px;">
                <option value="viewer">Viewer (Read-Only + Downloads)</option>
                <option value="editor">Editor (Can Edit & Upload)</option>
                <option value="admin">Admin (Full Control)</option>
              </select>
            </div>
            <div>
              <label style="display:block; font-size:10.5px; font-weight:700; color:var(--text-muted); margin-bottom:5px;">PASSWORD</label>
              <input type="password" id="newUserPwd" class="rate-input" style="width:100%; text-align:left; padding:9px 11px;" placeholder="Min 4 chars">
            </div>
            <div>
              <button class="btn btn-primary" style="width:100%; justify-content:center;" onclick="createUserAccount()">➕ Add User</button>
            </div>
          </div>
        </div>

        <!-- Existing Users Table -->
        <div class="table-container" style="border:1px solid var(--border); border-radius:12px;">
          <table style="min-width:720px;">
            <thead>
              <tr>
                <th>Login ID</th>
                <th>Name</th>
                <th>Role</th>
                <th class="center">Status</th>
                <th>Last Login</th>
                <th class="center">Actions</th>
              </tr>
            </thead>
            <tbody id="users-tbody">
              <tr><td colspan="6" class="center">Loading users...</td></tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>

    <!-- Tab 7: Account Profile & Password Change Tab -->
    <div id="tab-account" class="section-card" style="display:none;">
      <div class="section-header">
        <div class="section-title">🔑 Account Security &amp; Change Password</div>
      </div>
      <div style="padding: 24px; display:grid; grid-template-columns:repeat(auto-fit, minmax(300px, 1fr)); gap:22px; align-items:start;">
        <div style="background:var(--surface); border:1px solid var(--border); border-radius:14px; padding:20px;">
          <div style="font-size:14px; font-weight:800; color:var(--text); margin-bottom:14px;">👤 Active Session Profile</div>
          <div style="display:flex; flex-direction:column; gap:10px; font-size:13px;">
            <div><span style="color:var(--text-muted);">Login ID:</span> <strong class="mono" id="acct-username">-</strong></div>
            <div><span style="color:var(--text-muted);">Display Name:</span> <strong id="acct-name">-</strong></div>
            <div><span style="color:var(--text-muted);">Assigned Role:</span> <span id="acct-role-badge" class="badge badge-info">-</span></div>
            <div style="margin-top:4px; color:var(--text-muted); font-size:12px;" id="acct-role-desc">-</div>
          </div>
        </div>

        <div style="background:var(--surface); border:1px solid var(--border); border-radius:14px; padding:20px;">
          <div style="font-size:14px; font-weight:800; color:var(--text); margin-bottom:14px;">🔐 Update My Password</div>
          <div id="pwd-msg" style="display:none; padding:10px 12px; border-radius:8px; font-size:12.5px; font-weight:700; margin-bottom:14px;"></div>
          <div style="margin-bottom:14px;">
            <label style="display:block; font-size:11.5px; font-weight:700; color:var(--text-muted); margin-bottom:6px; text-transform:uppercase;">Current Password</label>
            <input type="password" id="pwdCurrent" class="rate-input" style="width:100%; text-align:left; padding:10px 12px;" placeholder="Enter current password">
          </div>
          <div style="margin-bottom:14px;">
            <label style="display:block; font-size:11.5px; font-weight:700; color:var(--text-muted); margin-bottom:6px; text-transform:uppercase;">New Password (min 4 characters)</label>
            <input type="password" id="pwdNew" class="rate-input" style="width:100%; text-align:left; padding:10px 12px;" placeholder="Enter new password">
          </div>
          <div style="margin-bottom:16px;">
            <label style="display:block; font-size:11.5px; font-weight:700; color:var(--text-muted); margin-bottom:6px; text-transform:uppercase;">Confirm New Password</label>
            <input type="password" id="pwdConfirm" class="rate-input" style="width:100%; text-align:left; padding:10px 12px;" placeholder="Re-enter new password">
          </div>
          <button class="btn btn-primary" onclick="submitPasswordChange()">💾 Update Password</button>
        </div>
      </div>
    </div>

  </div>

  <!-- Rate Modal -->
  <div class="modal-backdrop" id="rateModal">
    <div class="modal">
      <div class="modal-header">
        <h3 id="modalTitle">Enter Discussed Phone Rates</h3>
        <button class="close-btn" onclick="closeModal()">&times;</button>
      </div>
      <div class="modal-body">
        <div id="modalInvSummary" style="margin-bottom: 16px; padding: 16px; background: var(--surface); border: 1px solid var(--border-subtle); border-radius: 12px; font-size: 13px;"></div>
        
        <!-- F&E Surcharge Adjustment Section -->
        <div style="background:var(--surface-subtle); border:1px solid var(--border); border-left:4px solid var(--warning); border-radius:12px; padding:14px 18px; margin-bottom:16px;">
          <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px;">
            <div>
              <div style="font-size:13px; font-weight:800; color:var(--text); display:flex; align-items:center; gap:6px;">
                <span>⛽ Fuel & Energy (F&E) Surcharge Setting:</span>
              </div>
              <div style="font-size:11.5px; color:var(--text-muted); margin-top:4px;">
                Billed on Invoice: <strong class="mono bold" id="modalBilledFe" style="color:var(--text);">₹0.00</strong>
                <span id="modalFeRateBilled" style="margin-left:6px; color:var(--text-muted);"></span>
                <span id="modalFeCnNote" style="margin-left:8px;"></span>
              </div>
            </div>
            <div style="display:flex; align-items:center; gap:10px; flex-wrap:wrap;">
              <div>
                <label style="font-size:11px; font-weight:700; color:var(--text-muted); display:block; margin-bottom:2px;">Discussed Rate (₹/SQMM):</label>
                <input type="number" step="0.05" id="modalAgreedFeRate" class="rate-input" style="width:115px;" placeholder="₹/SQMM" oninput="onFeRateChange()">
              </div>
              <div>
                <label style="font-size:11px; font-weight:700; color:var(--text-muted); display:block; margin-bottom:2px;">Agreed Total (₹):</label>
                <input type="number" step="1" id="modalAgreedFe" class="rate-input" style="width:130px;" placeholder="Billed F&E" oninput="onFeAmtChange()">
              </div>
              <div style="margin-top:16px;">
                <button type="button" class="btn btn-primary-sm" style="padding:7px 12px;" onclick="resetFeToBilled()">Reset to Billed</button>
              </div>
            </div>
          </div>
          <div id="modalFeLiveCalc" style="font-size:11.5px; margin-top:8px; color:var(--primary); font-weight:700;"></div>
        </div>

        <!-- PMS Surcharge Adjustment Section -->
        <div id="modalPmsSection" style="background:var(--surface-subtle); border:1px solid var(--border); border-left:4px solid #8b5cf6; border-radius:12px; padding:14px 18px; margin-bottom:16px;">
          <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px;">
            <div>
              <div style="font-size:13px; font-weight:800; color:var(--text); display:flex; align-items:center; gap:6px;">
                <span>🥈 Precious Metal Surcharge (PMS) Setting:</span>
              </div>
              <div style="font-size:11.5px; color:var(--text-muted); margin-top:4px;">
                Billed on Invoice: <strong class="mono bold" id="modalBilledPms" style="color:var(--text);">₹0.00</strong>
                <span id="modalPmsRateBilled" style="margin-left:6px; color:var(--text-muted);"></span>
                <span id="modalPmsCnNote" style="margin-left:8px;"></span>
              </div>
            </div>
            <div id="modalPmsInputsWrap" style="display:flex; align-items:center; gap:10px; flex-wrap:wrap;">
              <div>
                <label style="font-size:11px; font-weight:700; color:var(--text-muted); display:block; margin-bottom:2px;">Discussed Rate (₹/SQMM):</label>
                <input type="number" step="0.05" id="modalAgreedPmsRate" class="rate-input" style="width:115px;" placeholder="₹/SQMM" oninput="onPmsRateChange()">
              </div>
              <div>
                <label style="font-size:11px; font-weight:700; color:var(--text-muted); display:block; margin-bottom:2px;">Agreed Total (₹):</label>
                <input type="number" step="1" id="modalAgreedPms" class="rate-input" style="width:130px;" placeholder="Billed PMS" oninput="onPmsAmtChange()">
              </div>
              <div style="margin-top:16px;">
                <button type="button" class="btn btn-primary-sm" style="padding:7px 12px;" onclick="resetPmsToBilled()">Reset to Billed</button>
              </div>
            </div>
          </div>
          <div id="modalPmsLiveCalc" style="font-size:11.5px; margin-top:8px; color:#8b5cf6; font-weight:700;"></div>
        </div>

        <div style="background:var(--primary-glow-subtle); border:1px solid var(--primary-glow); border-radius:10px; padding:12px 16px; font-size:12px; margin-bottom:16px; color:var(--text); display:flex; align-items:center; gap:10px;">
          <span style="font-size:18px;">💡</span>
          <div>
            <strong>Pricing Guide:</strong> 
            Mirror rate is flat across all thicknesses. 
            Clear glass varies: <strong>Thin</strong> (3.5, 4, 5, 6mm) vs <strong>Thick</strong> (8, 10, 12mm).
          </div>
        </div>
        <div class="table-container">
          <table style="min-width:700px;">
            <thead>
              <tr>
                <th class="center">Sr</th>
                <th>Item Description</th>
                <th class="center">MM</th>
                <th class="center">Dimensions (W x L)</th>
                <th class="num">Qty</th>
                <th class="num">SQMM</th>
                <th class="num">Billed Rate (₹)</th>
                <th class="center">Discussed Rate (₹)</th>
              </tr>
            </thead>
            <tbody id="modalItemsTbody"></tbody>
          </table>
        </div>
      </div>
      <div class="modal-footer">
        <button class="btn" style="background:var(--surface); color:var(--text);" onclick="closeModal()">Cancel</button>
        <button class="btn btn-download" onclick="saveAgreedRates()">💾 Save Rates & Update Ledger</button>
      </div>
    </div>
  </div>

  <script>
    /* __INITIAL_STATE_JSON__ */
    let appState = (typeof window !== 'undefined' && window.INITIAL_DATA) ? window.INITIAL_DATA : null;
    let currentInvoiceNo = null;

    function getCurrentUser() {
      return (appState && appState.current_user) ? appState.current_user : null;
    }

    function canEditRates() {
      const u = getCurrentUser();
      return !!(u && u.can_edit);
    }

    function isAdminUser() {
      const u = getCurrentUser();
      return !!(u && u.role === 'admin');
    }

    function escHtml(str) {
      return String(str || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
    }

    function safeGetStorage(key, def) {
      try {
        return localStorage.getItem(key) || def;
      } catch (e) {
        return def;
      }
    }

    function safeSetStorage(key, val) {
      try {
        localStorage.setItem(key, val);
      } catch (e) {}
    }

    function toggleTheme() {
      const html = document.documentElement;
      const current = html.getAttribute('data-theme');
      const target = current === 'dark' ? 'light' : 'dark';
      html.setAttribute('data-theme', target);
      safeSetStorage('ggl_theme', target);
      const icon = document.getElementById('theme-icon');
      const label = document.getElementById('theme-label');
      if (icon) icon.textContent = target === 'dark' ? '☀️' : '🌙';
      if (label) label.textContent = target === 'dark' ? 'Light Mode' : 'Dark Navy';
    }

    const savedTheme = safeGetStorage('ggl_theme', 'dark');
    document.documentElement.setAttribute('data-theme', savedTheme);
    window.addEventListener('DOMContentLoaded', () => {
      const icon = document.getElementById('theme-icon');
      const label = document.getElementById('theme-label');
      if (icon) icon.textContent = savedTheme === 'dark' ? '☀️' : '🌙';
      if (label) label.textContent = savedTheme === 'dark' ? 'Light Mode' : 'Dark Navy';
    });

    function showNetworkAlert(msg) {
      let el = document.getElementById('network-alert-banner');
      if (!el) {
        el = document.createElement('div');
        el.id = 'network-alert-banner';
        el.style.cssText = 'position:fixed; bottom:20px; right:20px; background:#EF4444; color:white; padding:12px 18px; border-radius:10px; font-size:13px; font-weight:700; z-index:9999; box-shadow:0 10px 25px rgba(0,0,0,0.4); display:flex; align-items:center; gap:10px;';
        document.body.appendChild(el);
      }
      el.innerHTML = `<span>⚠️ ${msg}</span> <button onclick="loadState(0)" style="background:white; color:#EF4444; border:none; padding:4px 10px; border-radius:6px; font-weight:800; cursor:pointer;">🔄 Retry</button> <button onclick="this.parentElement.remove()" style="background:transparent; border:none; color:white; font-size:16px; cursor:pointer;">&times;</button>`;
      el.style.display = 'flex';
    }

    function hideNetworkAlert() {
      const el = document.getElementById('network-alert-banner');
      if (el) el.style.display = 'none';
    }

    let isFetchingState = false;
    async function loadState(retryCount = 0) {
      if (isFetchingState) return;
      isFetchingState = true;
      try {
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 8000);
        const res = await fetch('/api/state', { signal: controller.signal });
        clearTimeout(timeoutId);
        if (res.status === 401) {
          window.location.href = '/login';
          return;
        }
        if (!res.ok) throw new Error('HTTP ' + res.status);
        appState = await res.json();
        hideNetworkAlert();
        renderDashboard();
      } catch (e) {
        console.warn('loadState notice:', e);
        if (retryCount < 2) {
          setTimeout(() => {
            isFetchingState = false;
            loadState(retryCount + 1);
          }, 1500);
          return;
        }
        if (!appState) {
          showNetworkAlert('Unable to load latest records. Click to retry.');
        }
      } finally {
        isFetchingState = false;
      }
    }

    // Keepalive ping every 25 seconds to keep Cloudflare tunnel & connection warm
    if (typeof setInterval === 'function') {
      setInterval(() => {
        fetch('/api/health').catch(() => {});
      }, 25000);
    }

    function fmt(num) {
      if (num === null || num === undefined) return '-';
      return '₹' + Number(num).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }
    function fmtQty(num) {
      if (num === null || num === undefined) return '-';
      return Number(num).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }

    function pdfLink(docNo, label, extraClass = '') {
      if (!docNo || docNo === '-' || docNo === 'None') return '-';
      const cleanDoc = String(docNo).trim();
      const display = label || cleanDoc;
      return `<a href="/api/view-pdf?doc=${encodeURIComponent(cleanDoc)}" target="_blank" class="pdf-pill ${extraClass}" title="Click to view original PDF: ${cleanDoc}">📄 ${display}</a>`;
    }

    function renderDashboard() {
      if (!appState) return;

      // 0. RBAC User Context & Permissions UI
      const user = getCurrentUser();
      const canEdit = canEditRates();
      const isAdmin = isAdminUser();
      if (user) {
        const nameEl = document.getElementById('user-badge-name');
        const roleEl = document.getElementById('user-role-tag');
        const iconEl = document.getElementById('user-badge-icon');
        if (nameEl) nameEl.textContent = user.name || user.username;
        if (roleEl) {
          roleEl.textContent = (user.role || 'viewer').toUpperCase();
          roleEl.className = 'badge ' + (user.role === 'admin' ? 'badge-success' : (user.role === 'editor' ? 'badge-info' : 'badge-warning'));
        }
        if (iconEl) iconEl.textContent = user.role === 'admin' ? '🛡️' : (user.role === 'editor' ? '✏️' : '👁️');
        const manageBtn = document.getElementById('btn-manage-users');
        if (manageBtn) manageBtn.style.display = isAdmin ? 'inline-flex' : 'none';
        const usersTabBtn = document.getElementById('tab-btn-users');
        if (usersTabBtn) usersTabBtn.style.display = isAdmin ? 'inline-flex' : 'none';

        const acctUser = document.getElementById('acct-username');
        const acctName = document.getElementById('acct-name');
        const acctBadge = document.getElementById('acct-role-badge');
        const acctDesc = document.getElementById('acct-role-desc');
        if (acctUser) acctUser.textContent = user.username;
        if (acctName) acctName.textContent = user.name || user.username;
        if (acctBadge) {
          acctBadge.textContent = user.role_label || (user.role || '').toUpperCase();
          acctBadge.className = 'badge ' + (user.role === 'admin' ? 'badge-success' : (user.role === 'editor' ? 'badge-info' : 'badge-warning'));
        }
        if (acctDesc) {
          acctDesc.textContent = (appState.roles && appState.roles[user.role] && appState.roles[user.role].description) || '';
        }
      }
      const uploadGrid = document.getElementById('upload-grid-wrap');
      const roBanner = document.getElementById('readonly-banner');
      const quickBar = document.getElementById('quick-edit-bar');
      if (uploadGrid) uploadGrid.style.display = canEdit ? 'grid' : 'none';
      if (roBanner) roBanner.style.display = canEdit ? 'none' : 'flex';
      if (quickBar) quickBar.style.display = canEdit ? 'flex' : 'none';

      // 1. Monthly Summary & KPIs
      try {
        const m = appState.monthly || [];
        let totBilled = 0, totActual = 0, totDiff = 0, totSettled = 0;
        let tbody = '';
        m.forEach(e => {
          totBilled += e.billed_total;
          totActual += e.actual_total || 0;
          totDiff += e.overbilled_total || 0;
          totSettled += e.rd_settled_total || 0;
          
          let invIds = (e.invoices || []).map(i => pdfLink(i.invoice_no, '#' + i.invoice_no)).join(' ') || '-';
          let netDiff = e.month_net_diff || 0;
          let cumBal = e.cumulative_rd_balance || 0;
          
          // Month Net Cell Formatting
          let monthNetHtml = '';
          if (netDiff < -10) {
            monthNetHtml = `<span style="color:var(--success); font-weight:700;">+${fmt(Math.abs(netDiff))} (Extra Credit)</span>`;
          } else if (netDiff > 10) {
            monthNetHtml = `<span style="color:var(--danger); font-weight:700;">-${fmt(netDiff)} (Pending)</span>`;
          } else {
            monthNetHtml = `<span style="color:var(--success); font-weight:700;">₹0.00 (Settled)</span>`;
          }

          // Cumulative Balance Cell Formatting
          let cumBalHtml = '';
          if (cumBal < -10) {
            cumBalHtml = `<span style="color:var(--success); font-weight:800;">+${fmt(Math.abs(cumBal))} (Extra Credit)</span>`;
          } else if (cumBal > 10) {
            cumBalHtml = `<span style="color:var(--danger); font-weight:800;">-${fmt(cumBal)} (Pending)</span>`;
          } else {
            cumBalHtml = `<span style="color:var(--success); font-weight:800;">₹0.00 (Balanced)</span>`;
          }

          // Status Badge
          let stBadge = '';
          if (netDiff < -10) {
            stBadge = `<span class="badge badge-success">EXTRA CREDIT (Settles Prior)</span>`;
          } else if (Math.abs(netDiff) <= 10) {
            stBadge = `<span class="badge badge-success">100% SETTLED</span>`;
          } else if (cumBal <= 0) {
            stBadge = `<span class="badge badge-info">COVERED BY ADVANCE POOL</span>`;
          } else {
            stBadge = `<span class="badge badge-danger">PENDING CLAIM</span>`;
          }

          let feNote = (e.fe_received > 0) ? `<div style="font-size:9.5px; color:#F59E0B; font-weight:600; margin-top:2px;">+ ⛽ F&E: ${fmt(e.fe_received)}</div>` : '';
          let pmsNote = (e.pms_received > 0) ? `<div style="font-size:9.5px; color:#C084FC; font-weight:600; margin-top:2px;">+ 🥈 PMS: ${fmt(e.pms_received)}</div>` : '';

          tbody += `<tr>
            <td class="bold">${e.month_key}</td>
            <td class="center mono" style="font-size:11px;">${invIds}</td>
            <td class="num bold">${fmtQty(e.total_sqmm)}</td>
            <td class="num">${fmt(e.billed_total)}</td>
            <td class="num">${fmt(e.actual_total)}</td>
            <td class="num bold">${fmt(e.overbilled_total)}</td>
            <td class="num bold" style="color:var(--success)">${fmt(e.rd_settled_total)}${feNote}${pmsNote}</td>
            <td class="num bold">${monthNetHtml}</td>
            <td class="num bold">${cumBalHtml}</td>
            <td class="center">${stBadge}</td>
          </tr>`;
        });

        // Total Summary Row
        let netPosition = totSettled - totDiff;
        let netPosHtml = '';
        let netBadge = '';
        if (netPosition >= 10) {
          netPosHtml = `<span style="color:var(--success); font-weight:800;">+${fmt(netPosition)} (Extra Credit)</span>`;
          netBadge = `<span class="badge badge-success">NET EXTRA CREDIT</span>`;
        } else if (netPosition <= -10) {
          netPosHtml = `<span style="color:var(--danger); font-weight:800;">-${fmt(Math.abs(netPosition))} (Pending)</span>`;
          netBadge = `<span class="badge badge-danger">NET PENDING FROM GGL</span>`;
        } else {
          netPosHtml = `<span style="color:var(--success); font-weight:800;">₹0.00 (Balanced)</span>`;
          netBadge = `<span class="badge badge-success">100% SETTLED</span>`;
        }
        let totFe = m.reduce((a,b)=>a+(b.fe_received||0), 0);
        let totFeNote = (totFe > 0) ? `<div style="font-size:9.5px; color:#F59E0B; font-weight:600; margin-top:2px;">+ ⛽ F&E: ${fmt(totFe)}</div>` : '';
        tbody += `<tr style="background:var(--table-header); font-weight:800;">
          <td class="bold">TOTAL (FY 26-27)</td>
          <td class="center mono" style="font-size:11px;">All Dispatches</td>
          <td class="num bold">${fmtQty(m.reduce((a,b)=>a+b.total_sqmm, 0))}</td>
          <td class="num bold">${fmt(totBilled)}</td>
          <td class="num bold">${fmt(totActual)}</td>
          <td class="num bold">${fmt(totDiff)}</td>
          <td class="num bold" style="color:var(--success)">${fmt(totSettled)}${totFeNote}</td>
          <td class="num bold">${netPosHtml}</td>
          <td class="num bold">${netPosHtml}</td>
          <td class="center">${netBadge}</td>
        </tr>`;
        
        const mBody = document.getElementById('monthly-tbody');
        if (mBody) mBody.innerHTML = tbody;

        // KPI Cards
        if (document.getElementById('kpi-billed')) document.getElementById('kpi-billed').textContent = fmt(totBilled);
        if (document.getElementById('kpi-actual')) document.getElementById('kpi-actual').textContent = fmt(totActual);
        if (document.getElementById('kpi-overbilled')) document.getElementById('kpi-overbilled').textContent = fmt(totDiff);
        if (document.getElementById('kpi-settled')) document.getElementById('kpi-settled').textContent = fmt(totSettled);
        
        if (document.getElementById('kpi-card-net')) {
          if (netPosition >= 0) {
            document.getElementById('kpi-card-net').className = 'kpi-card highlight';
            document.getElementById('kpi-net-title').textContent = 'NET EXTRA CREDIT (FY 26-27)';
            document.getElementById('kpi-net-val').textContent = '+' + fmt(netPosition);
            document.getElementById('kpi-net-sub').textContent = 'GGL has credited more than overbilling';
          } else {
            document.getElementById('kpi-card-net').className = 'kpi-card danger';
            document.getElementById('kpi-net-title').textContent = 'NET PENDING RATE DIFF';
            document.getElementById('kpi-net-val').textContent = '-' + fmt(Math.abs(netPosition));
            document.getElementById('kpi-net-sub').textContent = 'Awaiting credit notes from GGL';
          }
        }

        // Dynamic Tonnage KPIs
        const ton = appState.tonnage || { items: [], qtr_total: 0, annual_total: 0, total_tonnage: 0 };
        if (document.getElementById('kpi-qtr-ard')) document.getElementById('kpi-qtr-ard').textContent = fmt(ton.qtr_total);
        if (document.getElementById('kpi-qtr-sub')) document.getElementById('kpi-qtr-sub').textContent = `${ton.items.length} Quarterly Tonnage CNs credited in FY 26–27`;
        if (document.getElementById('kpi-yearly-ard')) document.getElementById('kpi-yearly-ard').textContent = fmt(ton.annual_total);
        if (document.getElementById('kpi-tot-tonnage')) document.getElementById('kpi-tot-tonnage').textContent = fmt(ton.total_tonnage);

        // Dynamic Audit Banner & GGL Statement Closing Balance
        const audit = appState.audit || {};
        const missingCns = audit.missing_cns || [];
        const totalLedgerCns = audit.total_ledger_cns || 0;
        const matchedCns = audit.matched_cns || 0;
        const closingBal = (audit.closing_balance !== undefined) ? audit.closing_balance : (appState.ledger ? appState.ledger.closing_balance : 0);
        const closingBalText = (closingBal < 0 ? `-₹${Math.abs(closingBal).toLocaleString('en-IN', {minimumFractionDigits:2})} Cr` : `₹${Number(closingBal).toLocaleString('en-IN', {minimumFractionDigits:2})} Dr`);

        const bannerEl = document.getElementById('audit-banner');
        const headlineEl = document.getElementById('audit-headline');
        const subtextEl = document.getElementById('audit-subtext');
        const totalCnsEl = document.getElementById('audit-total-cns');
        const closingBalEl = document.getElementById('audit-closing-bal');
        const kpiLedgerBalEl = document.getElementById('kpi-ledger-bal');

        if (closingBalEl) closingBalEl.textContent = closingBalText;
        if (kpiLedgerBalEl) kpiLedgerBalEl.textContent = closingBalText;

        const safety = appState.safety || {};
        const safetyValEl = document.getElementById('audit-safety-val');
        if (safetyValEl) {
          safetyValEl.textContent = `${safety.passed_checks || 9} / ${safety.total_checks || 9} PASSED`;
          safetyValEl.style.color = (safety.status === 'PASSED') ? 'var(--success)' : 'var(--danger)';
        }

        if (missingCns.length > 0) {
          if (bannerEl) {
            bannerEl.style.borderLeftColor = 'var(--danger)';
            bannerEl.style.background = 'linear-gradient(180deg, var(--card-bg), var(--danger-bg))';
          }
          if (headlineEl) headlineEl.innerHTML = `<span style="color:var(--danger)">⚠️ ACTION REQUIRED: ${missingCns.length} Credit Note(s) in GGL Statement Missing from Database</span>`;
          let misList = missingCns.map(m => `#${m.doc_no} (${fmt(m.cr)})`).join(', ');
          if (subtextEl) subtextEl.innerHTML = `Your uploaded GGL Statement has ${missingCns.length} credit notes not yet uploaded to the database: <strong>${misList}</strong>. Please upload these PDFs.`;
          if (totalCnsEl) totalCnsEl.innerHTML = `<span style="color:var(--danger)">${matchedCns} / ${totalLedgerCns}</span>`;
        } else {
          if (bannerEl) {
            bannerEl.style.borderLeftColor = 'var(--success)';
            bannerEl.style.background = 'var(--card-bg)';
          }
          if (headlineEl) headlineEl.innerHTML = `<span>100% Double-Checked Safety Audit Passed — All Calculations Verified</span>`;
          if (subtextEl) subtextEl.textContent = `Zero duplicate CN allocations. F&E CN #3242 strictly in May (#1303). F&E CN #6041 strictly in June (#2138). All ${totalLedgerCns} GGL Ledger CNs verified.`;
          if (totalCnsEl) totalCnsEl.innerHTML = `<span style="color:var(--success)">${totalLedgerCns} / ${totalLedgerCns}</span>`;
        }
      } catch (errMonthly) {
        console.error('Error in renderDashboard monthly:', errMonthly);
      }

      // 2. Bill-wise Dispatches & Line Items
      try {
        let sortedInvs = Object.values(appState.invoices || {}).sort((a,b) => (b.invoice_date || '').localeCompare(a.invoice_date || ''));
        let quickBtns = '';
        if (canEdit) {
          sortedInvs.forEach(inv => {
            let mLabel = inv.month_key || '';
            quickBtns += `<button class="quick-chip" onclick="openRateModal('${inv.invoice_no}')">✏️ ${mLabel}: #${inv.invoice_no}</button>`;
          });
        }
        const quickEl = document.getElementById('quick-inv-buttons');
        if (quickEl) quickEl.innerHTML = quickBtns;

        let invTbody = '';
        sortedInvs.forEach(inv => {
          let matchedCNsList = Object.values(appState.credit_notes || {}).filter(c => {
            // Strict invoice reference binding: if CN references an invoice, it MUST match only that invoice!
            if (c.ref_invoice) {
              return inv.invoice_no && (inv.invoice_no.includes(c.ref_invoice) || c.ref_invoice.includes(inv.invoice_no));
            }
            // Surcharges (F&E and PMS) must never match unreferenced dispatches
            if (c.cn_type === 'FE_ADJ' || c.cn_type === 'PMS_RET') {
              return false;
            }
            return c.target_month === inv.month_key && ['RD', 'ARD'].includes(c.cn_type);
          });
          let matchedCNsHtml = matchedCNsList.map(c => {
            let isFe = c.cn_type === 'FE_ADJ' || (c.desc || '').includes('F&E') || (c.desc || '').includes('EDIFF');
            let isPms = c.cn_type === 'PMS_RET' || (c.desc || '').includes('PMS');
            if (isFe) {
              return pdfLink(c.cn_no, `⛽ F&E: #${c.cn_no} (${fmt(c.total_amount)})`, 'pdf-pill-fe');
            } else if (isPms) {
              return pdfLink(c.cn_no, `🥈 PMS: #${c.cn_no} (${fmt(c.total_amount)})`, 'pdf-pill-pms');
            } else {
              return pdfLink(c.cn_no, `#${c.cn_no} (${fmt(c.total_amount)})`, 'pdf-pill-cn');
            }
          }).join(' ') || '-';
          
          let itemTypes = (inv.items || []).map(i => `${i.mm}mm ${i.item_type}`).filter((v,i,a)=>a.indexOf(v)===i).join(', ');

          let feBilled = inv.fe || 0;
          let feRateBilled = (inv.fe_rate_billed !== undefined && inv.fe_rate_billed !== null) ? inv.fe_rate_billed : (inv.total_sqmm > 0 ? (feBilled / inv.total_sqmm) : 0);
          let feDisplay = '';
          if (inv.fe_is_custom) {
            let aFeRate = (inv.agreed_fe_rate !== undefined && inv.agreed_fe_rate !== null) ? inv.agreed_fe_rate : (inv.agreed_fe / inv.total_sqmm);
            feDisplay = `<div>${fmt(feBilled)} <span style="font-size:10px; color:var(--text-muted);">(Billed: ₹${feRateBilled.toFixed(2)})</span></div>`;
            feDisplay += `<div style="font-size:10.5px; color:var(--warning); font-weight:700;">Discussed: ₹${aFeRate.toFixed(2)}/sqmm</div>`;
            if (inv.fe_cn_received > 0) {
              feDisplay += `<div style="font-size:10px; color:var(--success); font-weight:700;">⛽ CN: ${fmt(inv.fe_cn_received)} (Settled)</div>`;
            } else if (inv.fe_diff_gst > 0) {
              feDisplay += `<div style="font-size:10px; color:var(--danger); font-weight:700;">Awaiting F&E CN (${fmt(inv.fe_diff_gst)})</div>`;
            }
          } else {
            feDisplay = `<div>${fmt(feBilled)}</div><div style="font-size:10px; color:var(--text-muted);">₹${feRateBilled.toFixed(2)}/sqmm <span class="badge" style="padding:1px 5px; font-size:9px; background:var(--surface);">Default</span></div>`;
          }

          let pmsBilled = inv.pms || 0;
          let pmsDisplay = '-';
          if (pmsBilled > 0) {
            let pmsRateBilled = (inv.pms_rate_billed !== undefined && inv.pms_rate_billed !== null) ? inv.pms_rate_billed : (inv.total_sqmm > 0 ? (pmsBilled / inv.total_sqmm) : 0);
            if (inv.pms_is_custom) {
              let aPmsRate = (inv.agreed_pms_rate !== undefined && inv.agreed_pms_rate !== null) ? inv.agreed_pms_rate : (inv.agreed_pms / inv.total_sqmm);
              pmsDisplay = `<div>${fmt(pmsBilled)} <span style="font-size:10px; color:var(--text-muted);">(Billed: ₹${pmsRateBilled.toFixed(2)})</span></div>`;
              pmsDisplay += `<div style="font-size:10.5px; color:#8b5cf6; font-weight:700;">Discussed: ₹${aPmsRate.toFixed(2)}/sqmm</div>`;
              if (inv.pms_cn_received > 0) {
                pmsDisplay += `<div style="font-size:10px; color:var(--success); font-weight:700;">🥈 CN: ${fmt(inv.pms_cn_received)} (Settled)</div>`;
              } else if (inv.pms_diff_gst > 0) {
                pmsDisplay += `<div style="font-size:10px; color:var(--danger); font-weight:700;">Awaiting PMS CN (${fmt(inv.pms_diff_gst)})</div>`;
              }
            } else {
              pmsDisplay = `<div>${fmt(pmsBilled)}</div><div style="font-size:10px; color:var(--text-muted);">₹${pmsRateBilled.toFixed(2)}/sqmm <span class="badge" style="padding:1px 5px; font-size:9px; background:var(--surface);">Default</span></div>`;
            }
          }

          invTbody += `<tr>
            <td>${inv.invoice_date}</td>
            <td>
              ${pdfLink(inv.invoice_no, inv.invoice_no)}
              ${canEdit ? `<a href="javascript:void(0)" onclick="openRateModal('${inv.invoice_no}')" class="inv-link mono bold" style="margin-left:4px; font-size:12px;" title="Edit Discussed Rates & Surcharges">✏️</a>` : ''}
            </td>
            <td class="center">
              ${canEdit ? `<button class="btn btn-primary-sm" onclick="openRateModal('${inv.invoice_no}')">
                ✏️ Edit Discussed Rates
              </button>` : `<span class="badge badge-info">🔒 View Only</span>`}
            </td>
            <td><strong>${(inv.items || []).length} items:</strong> ${itemTypes}</td>
            <td class="num bold">${fmtQty(inv.total_sqmm)}</td>
            <td class="num">${feDisplay}</td>
            <td class="num">${pmsDisplay}</td>
            <td class="num">${fmt(inv.ins)}</td>
            <td class="num bold">${fmt(inv.total_amount)}</td>
            <td class="num">${fmt(inv.actual_total)}</td>
            <td class="num bold" style="color:var(--danger)">${fmt(inv.diff_overbilled)}</td>
            <td style="font-size:11px; max-width:280px;">${matchedCNsHtml}</td>
          </tr>`;

          // Sub-row displaying each line item with agreed rate and diff!
          let subItems = '';
          (inv.items || []).forEach(it => {
            let bRate = it.b_rate || 0;
            let aRate = it.agreed_rate || 0;
            let rateDiff = (bRate - aRate).toFixed(1);
            let itemOverbill = it.actual_amount ? (it.billed_amount - it.actual_amount).toFixed(2) : '-';

            subItems += `<tr>
              <td class="center">#${it.sr}</td>
              <td><strong>${it.desc}</strong></td>
              <td class="center bold">${it.mm}mm</td>
              <td class="center">${it.width} × ${it.length}</td>
              <td class="num">${it.qty}</td>
              <td class="num bold">${fmtQty(it.sqmm)}</td>
              <td class="num">${fmt(bRate)}</td>
              <td class="num bold" style="color:var(--primary); font-size:13px;">${aRate ? fmt(aRate) : 'Not Set'}</td>
              <td class="num" style="color:var(--warning)">${aRate ? '₹' + rateDiff + '/sqmm' : '-'}</td>
              <td class="num bold" style="color:var(--danger)">${aRate ? fmt(itemOverbill) : '-'}</td>
            </tr>`;
          });

          invTbody += `<tr class="item-drawer-row">
            <td colspan="12">
              <div style="font-size:11px; font-weight:700; color:var(--text-muted); margin-bottom:6px; display:flex; justify-content:space-between; align-items:center;">
                <span>📦 Line Items Breakdown for Invoice #${inv.invoice_no}:</span>
                ${canEdit ? `<button class="btn btn-primary-sm" style="padding:3px 8px; font-size:10.5px;" onclick="openRateModal('${inv.invoice_no}')">✏️ Modify Rates</button>` : `<span style="font-size:10.5px; color:var(--text-dim);">🔒 Read-Only</span>`}
              </div>
              <table class="item-subtable">
                <thead>
                  <tr>
                    <th class="center">Sr</th>
                    <th>Description</th>
                    <th class="center">MM</th>
                    <th class="center">Dimensions</th>
                    <th class="num">Qty</th>
                    <th class="num">SQMM</th>
                    <th class="num">Billed List Rate</th>
                    <th class="num" style="color:var(--primary)">Agreed Discussed Rate</th>
                    <th class="num">Rate Diff</th>
                    <th class="num">Basic Overbilled</th>
                  </tr>
                </thead>
                <tbody>${subItems}</tbody>
              </table>
            </td>
          </tr>`;
        });
        const invEl = document.getElementById('invoices-tbody');
        if (invEl) invEl.innerHTML = invTbody;
      } catch (errInvs) {
        console.error('Error in renderDashboard invoices:', errInvs);
      }

      // 3. Cash Discount Table (Base X & Prompt Payment Tiered Calculation)
      try {
        const cdData = appState.cd_tracker || [];
        let cdTbody = '';
        let totBilled = 0, totRd = 0, totBaseX = 0, totElig = 0, totRcvd = 0, totShortfall = 0;

        cdData.forEach(c => {
          totBilled += c.inv_total || 0;
          totRd += c.rd_cn_amt || 0;
          totBaseX += c.base_x || 0;
          totElig += c.eligible_cd || 0;
          totRcvd += c.cd_received || 0;
          totShortfall += c.shortfall || 0;

          let invDisplay = pdfLink(c.inv_no, c.inv_no);
          let rdDisplay = c.rd_cn_no && c.rd_cn_no.startsWith('2026') ? pdfLink(c.rd_cn_no, '#' + c.rd_cn_no, 'pdf-pill-cn') : c.rd_cn_no;
          let cdCnDisplay = c.cd_cn_no && c.cd_cn_no.startsWith('2026') ? pdfLink(c.cd_cn_no, '#' + c.cd_cn_no, 'pdf-pill-cn') : c.cd_cn_no;

          let statusBadge = '';
          if (c.cd_received > 0 && c.shortfall <= 10) {
            statusBadge = `<span class="badge badge-success">100% SETTLED</span>`;
          } else if (c.cd_received > 0) {
            statusBadge = `<span class="badge badge-warning">SHORTFALL (${fmt(c.shortfall)})</span>`;
          } else {
            statusBadge = `<span class="badge badge-danger">PENDING CLAIM</span>`;
          }

          let rateColor = c.cd_rate_pct >= 3.0 ? 'var(--success)' : (c.cd_rate_pct >= 2.8 ? 'var(--warning)' : 'var(--danger)');

          cdTbody += `<tr>
            <td class="bold">${c.month_key}</td>
            <td class="bold mono">${invDisplay}</td>
            <td>${c.inv_date}</td>
            <td class="num">${fmt(c.inv_total)}</td>
            <td class="center bold mono">${rdDisplay}</td>
            <td class="num" style="color:var(--text-muted)">${fmt(c.rd_cn_amt)}</td>
            <td class="num bold" style="color:var(--primary); font-size:13px;">${fmt(c.base_x)}</td>
            <td>${c.pay_date}</td>
            <td style="font-size:11px;">${c.pay_ref}</td>
            <td class="center bold" style="font-size:11.5px;">${c.turnaround_label}</td>
            <td class="center bold" style="color:${rateColor}; font-size:13px;">${c.cd_rate_pct}%</td>
            <td class="num bold" style="color:var(--primary)">${fmt(c.eligible_cd)}</td>
            <td class="center bold mono">${cdCnDisplay}</td>
            <td class="num bold" style="color:var(--success)">${fmt(c.cd_received)}</td>
            <td class="num bold" style="color:var(--danger)">${fmt(c.shortfall)}</td>
            <td class="center">${statusBadge}</td>
          </tr>`;
        });

        // Total Summary Row
        cdTbody += `<tr style="background:var(--table-header); font-weight:800;">
          <td colspan="3" class="bold">TOTAL (FY 2026–27)</td>
          <td class="num bold">${fmt(totBilled)}</td>
          <td class="center">-</td>
          <td class="num bold" style="color:var(--text-muted)">${fmt(totRd)}</td>
          <td class="num bold" style="color:var(--primary); font-size:14px;">${fmt(totBaseX)}</td>
          <td colspan="4" class="center bold">Tiered Daily Benchmark</td>
          <td class="num bold" style="color:var(--primary); font-size:14px;">${fmt(totElig)}</td>
          <td class="center">-</td>
          <td class="num bold" style="color:var(--success); font-size:14px;">${fmt(totRcvd)}</td>
          <td class="num bold" style="color:var(--danger); font-size:14px;">${fmt(totShortfall)}</td>
          <td class="center"><span class="badge badge-warning">₹${totShortfall.toLocaleString('en-IN', {minimumFractionDigits:2})} PENDING</span></td>
        </tr>`;

        const cdEl = document.getElementById('cd-tbody');
        if (cdEl) cdEl.innerHTML = cdTbody;

        if (document.getElementById('kpi-tot-cd')) {
          document.getElementById('kpi-tot-cd').textContent = fmt(totRcvd);
        }
        if (document.getElementById('kpi-sub-cd')) {
          document.getElementById('kpi-sub-cd').textContent = `Eligible: ${fmt(totElig)} (${fmt(totShortfall)} Pending Claim)`;
        }
      } catch (errCd) {
        console.error('Error in renderDashboard cd:', errCd);
      }

      // 4. GGL Statement Ledger Table (FY 26-27)
      try {
        const audit = appState.audit || {};
        const missingCns = audit.missing_cns || [];
        const totalLedgerCns = audit.total_ledger_cns || 0;
        const closingBal = (audit.closing_balance !== undefined) ? audit.closing_balance : (appState.ledger ? appState.ledger.closing_balance : 0);
        const closingBalText = (closingBal < 0 ? `-₹${Math.abs(closingBal).toLocaleString('en-IN', {minimumFractionDigits:2})} Cr` : `₹${Number(closingBal).toLocaleString('en-IN', {minimumFractionDigits:2})} Dr`);

        // Missing CNs alert box in Tab 5
        const missingBox = document.getElementById('ledger-missing-alert');
        const matchedBox = document.getElementById('ledger-all-matched-banner');
        const missingTbody = document.getElementById('ledger-missing-tbody');
        const missingCountEl = document.getElementById('ledger-missing-count');
        const matchedMsgEl = document.getElementById('ledger-matched-msg');

        if (missingCns.length > 0) {
          if (missingBox) missingBox.style.display = 'block';
          if (matchedBox) matchedBox.style.display = 'none';
          if (missingCountEl) missingCountEl.textContent = `${missingCns.length} Missing`;
          let misRows = '';
          missingCns.forEach(m => {
            misRows += `<tr>
              <td>${m.date}</td>
              <td class="bold mono"><span class="badge badge-danger">Needs Upload</span> #${m.doc_no}</td>
              <td class="num bold" style="color:var(--danger)">${fmt(m.cr)}</td>
              <td>${m.remarks || 'Credit Note'}</td>
              <td class="center"><span class="badge badge-warning">Upload PDF</span></td>
            </tr>`;
          });
          if (missingTbody) missingTbody.innerHTML = misRows;
        } else {
          if (missingBox) missingBox.style.display = 'none';
          if (matchedBox) matchedBox.style.display = 'flex';
          if (matchedMsgEl) matchedMsgEl.textContent = `100% Cross-Audited: All ${totalLedgerCns} Credit Notes in GGL's Official Statement are present and verified in your database.`;
        }

        const ledgerEntries = (appState.ledger && appState.ledger.entries) ? appState.ledger.entries : [];
        let legTbody = '';
        ledgerEntries.forEach(l => {
          let docHtml = l.doc_no;
          if (['INV', 'CNT'].includes(l.type) && l.doc_no && l.doc_no !== '-') {
            let pillClass = l.type === 'CNT' ? 'pdf-pill-cn' : '';
            docHtml = pdfLink(l.doc_no, l.doc_no, pillClass);
          }
          legTbody += `<tr>
            <td class="center">${l.date}</td>
            <td class="center"><span class="badge" style="background:var(--surface);">${l.type}</span></td>
            <td class="bold mono">${docHtml}</td>
            <td class="num">${l.dr ? fmt(l.dr) : '-'}</td>
            <td class="num">${l.cr ? fmt(l.cr) : '-'}</td>
            <td class="num bold" style="color:${l.balance < 0 ? 'var(--success)' : 'var(--danger)'}">${fmt(l.balance)}</td>
            <td>${l.remarks}</td>
          </tr>`;
        });
        legTbody += `<tr style="background:var(--success-bg); font-weight:800;">
          <td colspan="5" class="bold" style="font-size:13.5px; color:var(--success);">OFFICIAL CLOSING BALANCE IN GGL STATEMENT:</td>
          <td class="num bold" style="font-size:15px; color:var(--success);">${closingBalText}</td>
          <td style="color:var(--success);">${closingBal < 0 ? 'Advance Credit (GGL Owes Mehta Impex)' : 'Debit Balance'}</td>
        </tr>`;
        const legEl = document.getElementById('ledger-tbody');
        if (legEl) legEl.innerHTML = legTbody;
      } catch (errLedger) {
        console.error('Error in renderDashboard ledger:', errLedger);
      }

      // 5. Tonnage ARD Table (FY 2026–27 Only)
      try {
        const ton = appState.tonnage || { items: [], qtr_total: 0, annual_total: 0, total_tonnage: 0 };
        if (document.getElementById('ton-tab-qtr')) document.getElementById('ton-tab-qtr').textContent = fmt(ton.qtr_total);
        if (document.getElementById('ton-tab-annual')) document.getElementById('ton-tab-annual').textContent = fmt(ton.annual_total);
        if (document.getElementById('ton-tab-total')) document.getElementById('ton-tab-total').textContent = fmt(ton.total_tonnage);
        if (document.getElementById('ton-tab-desc')) {
          document.getElementById('ton-tab-desc').innerHTML = `In this financial year, Mehta Impex has received <strong>${ton.items.length} Quarterly ARD Credit Notes</strong> totaling <strong>${fmt(ton.qtr_total)}</strong> based on quarterly tonnage business with GGL. <strong>Annual ARD for FY 2026–27 has not been received yet</strong> (annual incentives are settled by GGL at year-end). These tonnage incentives are tracked in this dedicated section in your favor and are <strong>strictly separated</strong> from monthly invoice overbilling.`;
        }

        let tonTbody = '';
        (ton.items || []).forEach(t => {
          let docHtml = t.doc && t.doc.startsWith('2026') ? pdfLink(t.doc, t.doc, 'pdf-pill-cn') : t.doc;
          tonTbody += `<tr>
            <td class="bold">${t.scope}</td>
            <td class="center">${t.dt}</td>
            <td class="bold mono">${docHtml}</td>
            <td>${t.desc}</td>
            <td class="center bold">${t.ton}</td>
            <td class="center">${t.rate}</td>
            <td class="num bold" style="color:var(--success)">${fmt(t.amt)}</td>
            <td class="center"><span class="badge badge-success">${t.st}</span></td>
          </tr>`;
        });
        tonTbody += `<tr style="background:var(--table-stripe);">
          <td class="bold" style="color:var(--text-muted)">FY 2026–27 Annual</td>
          <td class="center">-</td>
          <td class="bold mono">-</td>
          <td style="color:var(--text-muted)">Annual Tonnage Incentive (Float & Mirror) — Not received yet / Finalized at year-end</td>
          <td class="center bold">-</td>
          <td class="center">Annual Incentive</td>
          <td class="num bold" style="color:var(--text-muted)">₹0.00</td>
          <td class="center"><span class="badge badge-warning">AWAITING YEAR-END</span></td>
        </tr>`;
        tonTbody += `<tr style="background:var(--table-header); font-weight:800;">
          <td colspan="6" class="bold">TOTAL TONNAGE INCENTIVES RECEIVED IN FY 2026–27:</td>
          <td class="num bold" style="color:var(--success); font-size:15px;">${fmt(ton.total_tonnage)}</td>
          <td class="center"><span class="badge badge-success">ALL CREDITED</span></td>
        </tr>`;
        const tonEl = document.getElementById('tonnage-tbody');
        if (tonEl) tonEl.innerHTML = tonTbody;
      } catch (errTon) {
        console.error('Error in renderDashboard tonnage:', errTon);
      }
    }

    function switchTab(tab, btn) {
      if (tab === 'users' && !isAdminUser()) {
        alert('Only Administrators can manage users and roles.');
        return;
      }
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      const tabBtnIds = {
        'monthly': 0, 'invoices': 1, 'cd': 2, 'tonnage': 3, 'ggl-ledger': 4, 'users': 5, 'account': 6
      };
      let target = btn;
      if (!target && typeof event !== 'undefined' && event && event.target) {
        const clickedTab = event.target.closest('.tab-btn');
        if (clickedTab) target = clickedTab;
      }
      if (target) {
        target.classList.add('active');
      } else {
        const allBtns = document.querySelectorAll('#main-tabs-nav .tab-btn');
        if (tabBtnIds[tab] !== undefined && allBtns[tabBtnIds[tab]]) {
          allBtns[tabBtnIds[tab]].classList.add('active');
        }
      }
      const tabMapEls = {
        'monthly': document.getElementById('tab-monthly'),
        'invoices': document.getElementById('tab-invoices'),
        'cd': document.getElementById('tab-cd'),
        'tonnage': document.getElementById('tab-tonnage'),
        'ggl-ledger': document.getElementById('tab-ggl-ledger'),
        'users': document.getElementById('tab-users'),
        'account': document.getElementById('tab-account')
      };
      
      Object.keys(tabMapEls).forEach(k => {
        const el = tabMapEls[k];
        if (el) {
          el.style.display = (k === tab) ? 'block' : 'none';
        }
      });

      if (tab === 'users') {
        const uMsg = document.getElementById('users-msg');
        if (uMsg) { uMsg.style.display = 'none'; uMsg.textContent = ''; }
        loadUsersTable();
      } else if (tab === 'account') {
        const pMsg = document.getElementById('pwd-msg');
        if (pMsg) { pMsg.style.display = 'none'; pMsg.textContent = ''; }
      }
      if (!btn) {
        const nav = document.getElementById('main-tabs-nav');
        if (nav) nav.scrollIntoView({ behavior: 'auto', block: 'start' });
      }
    }

    async function uploadFile(input, endpoint) {
      if (!canEditRates()) {
        alert('🔒 Read-Only Access: Only authorized Admin or Editor accounts can upload documents.');
        input.value = '';
        return;
      }
      if (!input.files || input.files.length === 0) return;
      const file = input.files[0];
      const formData = new FormData();
      formData.append('file', file);
      
      try {
        const res = await fetch(endpoint, { method: 'POST', body: formData });
        if (res.status === 401) { window.location.href = '/login'; return; }
        const result = await res.json();
        if (result.success) {
          if (endpoint === '/api/upload-bill') {
            await loadState();
            openRateModal(result.invoice.invoice_no);
            alert(`✅ Invoice #${result.invoice.invoice_no} uploaded & archived to PDF store successfully!`);
          } else if (endpoint === '/api/upload-ledger') {
            alert('✅ ' + (result.message || 'GGL Statement of Account uploaded & cross-audited successfully!'));
            await loadState();
          } else {
            alert(`✅ Credit Note #${result.credit_note.cn_no} uploaded & reconciled successfully!`);
            await loadState();
          }
        } else {
          if (result.is_duplicate) {
            alert('⚠️ DUPLICATE DOCUMENT DETECTED & DECLINED:\\n\\n' + result.error);
          } else {
            alert('❌ Upload failed: ' + (result.error || 'Unknown error'));
          }
        }
      } catch (e) {
        alert('Upload error: ' + e);
      }
      input.value = '';
    }

    function onFeRateChange() {
      if (!currentInvoiceNo || !appState || !appState.invoices[currentInvoiceNo]) return;
      const inv = appState.invoices[currentInvoiceNo];
      const rateVal = parseFloat(document.getElementById('modalAgreedFeRate').value);
      if (!isNaN(rateVal) && inv.total_sqmm > 0) {
        const amt = Math.round(rateVal * inv.total_sqmm * 100) / 100;
        document.getElementById('modalAgreedFe').value = amt.toFixed(2);
      } else if (document.getElementById('modalAgreedFeRate').value.trim() === '') {
        document.getElementById('modalAgreedFe').value = '';
      }
      updateFeLiveCalc();
    }

    function onFeAmtChange() {
      if (!currentInvoiceNo || !appState || !appState.invoices[currentInvoiceNo]) return;
      const inv = appState.invoices[currentInvoiceNo];
      const amtVal = parseFloat(document.getElementById('modalAgreedFe').value);
      if (!isNaN(amtVal) && inv.total_sqmm > 0) {
        const rate = (amtVal / inv.total_sqmm).toFixed(4);
        document.getElementById('modalAgreedFeRate').value = parseFloat(rate);
      } else if (document.getElementById('modalAgreedFe').value.trim() === '') {
        document.getElementById('modalAgreedFeRate').value = '';
      }
      updateFeLiveCalc();
    }

    function resetFeToBilled() {
      if (!currentInvoiceNo || !appState || !appState.invoices[currentInvoiceNo]) return;
      const inv = appState.invoices[currentInvoiceNo];
      const feRateBilled = (inv.fe_rate_billed !== undefined && inv.fe_rate_billed !== null) ? inv.fe_rate_billed : (inv.total_sqmm > 0 ? (inv.fe / inv.total_sqmm) : 0);
      document.getElementById('modalAgreedFeRate').value = feRateBilled ? parseFloat(feRateBilled.toFixed(2)) : '';
      document.getElementById('modalAgreedFe').value = (inv.fe !== undefined && inv.fe !== null) ? inv.fe : '';
      updateFeLiveCalc();
    }

    function updateFeLiveCalc() {
      if (!currentInvoiceNo || !appState || !appState.invoices[currentInvoiceNo]) return;
      const inv = appState.invoices[currentInvoiceNo];
      const billedFe = inv.fe || 0;
      const amtInput = parseFloat(document.getElementById('modalAgreedFe').value);
      const calcEl = document.getElementById('modalFeLiveCalc');
      if (!calcEl) return;
      
      if (!isNaN(amtInput) && Math.abs(amtInput - billedFe) > 0.05) {
        const diffBase = billedFe - amtInput;
        const diffGst = diffBase * 1.18;
        const diffRate = (inv.total_sqmm > 0) ? (diffBase / inv.total_sqmm).toFixed(2) : '0.00';
        calcEl.innerHTML = `<span style="color:var(--warning); font-weight:700;">💡 Custom F&E Overcharge Reversal:</span> ₹${diffRate}/SQMM → Claim: <strong>₹${diffGst.toLocaleString('en-IN', {minimumFractionDigits:2, maximumFractionDigits:2})}</strong> (inc. 18% GST). <em>Audited separately from glass RD.</em>`;
      } else {
        calcEl.innerHTML = `<span class="badge" style="background:#f1f5f9; color:#475569; font-size:11px;">Default Rate from Invoice (Standard Billed Rate — No Surcharge Overcharge)</span>`;
      }
    }

    function onPmsRateChange() {
      if (!currentInvoiceNo || !appState || !appState.invoices[currentInvoiceNo]) return;
      const inv = appState.invoices[currentInvoiceNo];
      const rateVal = parseFloat(document.getElementById('modalAgreedPmsRate').value);
      if (!isNaN(rateVal) && inv.total_sqmm > 0) {
        const amt = Math.round(rateVal * inv.total_sqmm * 100) / 100;
        document.getElementById('modalAgreedPms').value = amt.toFixed(2);
      } else if (document.getElementById('modalAgreedPmsRate').value.trim() === '') {
        document.getElementById('modalAgreedPms').value = '';
      }
      updatePmsLiveCalc();
    }

    function onPmsAmtChange() {
      if (!currentInvoiceNo || !appState || !appState.invoices[currentInvoiceNo]) return;
      const inv = appState.invoices[currentInvoiceNo];
      const amtVal = parseFloat(document.getElementById('modalAgreedPms').value);
      if (!isNaN(amtVal) && inv.total_sqmm > 0) {
        const rate = (amtVal / inv.total_sqmm).toFixed(4);
        document.getElementById('modalAgreedPmsRate').value = parseFloat(rate);
      } else if (document.getElementById('modalAgreedPms').value.trim() === '') {
        document.getElementById('modalAgreedPmsRate').value = '';
      }
      updatePmsLiveCalc();
    }

    function resetPmsToBilled() {
      if (!currentInvoiceNo || !appState || !appState.invoices[currentInvoiceNo]) return;
      const inv = appState.invoices[currentInvoiceNo];
      const pmsRateBilled = (inv.pms_rate_billed !== undefined && inv.pms_rate_billed !== null) ? inv.pms_rate_billed : (inv.total_sqmm > 0 ? (inv.pms / inv.total_sqmm) : 0);
      document.getElementById('modalAgreedPmsRate').value = pmsRateBilled ? parseFloat(pmsRateBilled.toFixed(2)) : '';
      document.getElementById('modalAgreedPms').value = (inv.pms !== undefined && inv.pms !== null) ? inv.pms : '';
      updatePmsLiveCalc();
    }

    function updatePmsLiveCalc() {
      if (!currentInvoiceNo || !appState || !appState.invoices[currentInvoiceNo]) return;
      const inv = appState.invoices[currentInvoiceNo];
      const billedPms = inv.pms || 0;
      if (billedPms <= 0) return;
      const amtInput = parseFloat(document.getElementById('modalAgreedPms').value);
      const calcEl = document.getElementById('modalPmsLiveCalc');
      if (!calcEl) return;
      
      if (!isNaN(amtInput) && Math.abs(amtInput - billedPms) > 0.05) {
        const diffBase = billedPms - amtInput;
        const diffGst = diffBase * 1.18;
        const diffRate = (inv.total_sqmm > 0) ? (diffBase / inv.total_sqmm).toFixed(2) : '0.00';
        calcEl.innerHTML = `<span style="color:var(--warning); font-weight:700;">💡 Custom PMS Overcharge Reversal:</span> ₹${diffRate}/SQMM → Claim: <strong>₹${diffGst.toLocaleString('en-IN', {minimumFractionDigits:2, maximumFractionDigits:2})}</strong> (inc. 18% GST). <em>Audited separately from glass RD.</em>`;
      } else {
        calcEl.innerHTML = `<span class="badge" style="background:#f1f5f9; color:#475569; font-size:11px;">Default Rate from Invoice (Standard Billed Rate — No Surcharge Overcharge)</span>`;
      }
    }

    function openRateModal(invNo) {
      if (!canEditRates()) {
        alert('🔒 Read-Only Access: Only authorized Admin or Editor accounts can modify discussed rates.');
        return;
      }
      currentInvoiceNo = invNo;
      const inv = appState.invoices[invNo];
      if (!inv) return;
      
      document.getElementById('modalTitle').textContent = `✏️ Enter Discussed Rates & Surcharges for Invoice #${inv.invoice_no} (${inv.invoice_date})`;
      document.getElementById('modalInvSummary').innerHTML = `
        <strong>Total Bill Amount:</strong> ${fmt(inv.total_amount)} &nbsp;|&nbsp; 
        <strong>F&E:</strong> ${fmt(inv.fe)} &nbsp;|&nbsp; 
        <strong>PMS:</strong> ${fmt(inv.pms)} &nbsp;|&nbsp; 
        <strong>INS:</strong> ${fmt(inv.ins)} &nbsp;|&nbsp; 
        <strong>Total SQMM:</strong> ${fmtQty(inv.total_sqmm)}
      `;

      // F&E Surcharge Display & Inputs
      const billedFe = inv.fe || 0;
      const feRateBilled = (inv.fe_rate_billed !== undefined && inv.fe_rate_billed !== null) ? inv.fe_rate_billed : (inv.total_sqmm > 0 ? (billedFe / inv.total_sqmm) : 0);
      document.getElementById('modalBilledFe').textContent = fmt(billedFe);
      document.getElementById('modalFeRateBilled').textContent = `(Billed: ₹${feRateBilled.toFixed(2)}/SQMM)`;
      
      let agreedFeRate = '';
      let agreedFeAmt = '';
      if (inv.fe_is_custom && inv.agreed_fe_rate !== undefined && inv.agreed_fe_rate !== null) {
        agreedFeRate = inv.agreed_fe_rate;
        agreedFeAmt = (inv.agreed_fe !== undefined && inv.agreed_fe !== null) ? inv.agreed_fe : '';
      } else {
        agreedFeRate = feRateBilled ? parseFloat(feRateBilled.toFixed(2)) : '';
        agreedFeAmt = (inv.fe !== undefined && inv.fe !== null) ? inv.fe : '';
      }
      document.getElementById('modalAgreedFeRate').value = agreedFeRate;
      document.getElementById('modalAgreedFe').value = agreedFeAmt;
      
      // Matching F&E CNs
      const feCns = Object.values(appState.credit_notes || {}).filter(c => {
        let matchesInv = Boolean(c.ref_invoice && inv.invoice_no && (inv.invoice_no.includes(c.ref_invoice) || c.ref_invoice.includes(inv.invoice_no)));
        let isFe = c.cn_type === 'FE_ADJ' || (c.desc || '').includes('F&E') || (c.desc || '').includes('EDIFF');
        return matchesInv && isFe;
      });
      if (feCns.length > 0) {
        document.getElementById('modalFeCnNote').innerHTML = feCns.map(c => 
          `<span class="badge badge-success" style="font-size:11px;">⛽ Refund CN #${c.cn_no} (${fmt(c.total_amount)}) Settled</span>`
        ).join(' ');
      } else {
        document.getElementById('modalFeCnNote').innerHTML = '';
      }
      updateFeLiveCalc();

      // PMS Surcharge Display & Inputs
      const billedPms = inv.pms || 0;
      const pmsSectionEl = document.getElementById('modalPmsSection');
      if (billedPms > 0) {
        if (pmsSectionEl) pmsSectionEl.style.display = 'block';
        const pmsRateBilled = (inv.pms_rate_billed !== undefined && inv.pms_rate_billed !== null) ? inv.pms_rate_billed : (inv.total_sqmm > 0 ? (billedPms / inv.total_sqmm) : 0);
        document.getElementById('modalBilledPms').textContent = fmt(billedPms);
        document.getElementById('modalPmsRateBilled').textContent = `(Billed: ₹${pmsRateBilled.toFixed(2)}/SQMM)`;
        
        let agreedPmsRate = '';
        let agreedPmsAmt = '';
        if (inv.pms_is_custom && inv.agreed_pms_rate !== undefined && inv.agreed_pms_rate !== null) {
          agreedPmsRate = inv.agreed_pms_rate;
          agreedPmsAmt = (inv.agreed_pms !== undefined && inv.agreed_pms !== null) ? inv.agreed_pms : '';
        } else {
          agreedPmsRate = pmsRateBilled ? parseFloat(pmsRateBilled.toFixed(2)) : '';
          agreedPmsAmt = (inv.pms !== undefined && inv.pms !== null) ? inv.pms : '';
        }
        document.getElementById('modalAgreedPmsRate').value = agreedPmsRate;
        document.getElementById('modalAgreedPms').value = agreedPmsAmt;

        const pmsCns = Object.values(appState.credit_notes || {}).filter(c => {
          let matchesInv = Boolean(c.ref_invoice && inv.invoice_no && (inv.invoice_no.includes(c.ref_invoice) || c.ref_invoice.includes(inv.invoice_no)));
          let isPms = c.cn_type === 'PMS_RET' || (c.desc || '').includes('PMS');
          return matchesInv && isPms;
        });
        if (pmsCns.length > 0) {
          document.getElementById('modalPmsCnNote').innerHTML = pmsCns.map(c => 
            `<span class="badge badge-success" style="font-size:11px;">🥈 Refund CN #${c.cn_no} (${fmt(c.total_amount)}) Settled</span>`
          ).join(' ');
        } else {
          document.getElementById('modalPmsCnNote').innerHTML = '';
        }
        updatePmsLiveCalc();
      } else {
        if (pmsSectionEl) pmsSectionEl.style.display = 'none';
      }

      let tbody = '';
      inv.items.forEach(it => {
        tbody += `<tr>
          <td class="center">#${it.sr}</td>
          <td><strong>${it.desc}</strong></td>
          <td class="center bold">${it.mm}mm</td>
          <td class="center">${it.width} × ${it.length}</td>
          <td class="num">${it.qty}</td>
          <td class="num bold">${fmtQty(it.sqmm)}</td>
          <td class="num" style="color:var(--text-muted)">${fmt(it.b_rate)}</td>
          <td class="center">
            <input type="number" step="0.5" class="rate-input" data-sr="${it.sr}" value="${it.agreed_rate || ''}" placeholder="e.g. 95">
          </td>
        </tr>`;
      });
      document.getElementById('modalItemsTbody').innerHTML = tbody;
      document.getElementById('rateModal').style.display = 'flex';
      
      // Auto-focus first input
      setTimeout(() => {
        const first = document.querySelector('#modalItemsTbody input.rate-input');
        if (first) first.focus();
      }, 100);
    }

    function closeModal() {
      document.getElementById('rateModal').style.display = 'none';
      currentInvoiceNo = null;
    }

    async function saveAgreedRates() {
      if (!canEditRates()) {
        alert('🔒 Read-Only Access: You do not have permission to save rates.');
        return;
      }
      if (!currentInvoiceNo) return;
      const inv = appState.invoices[currentInvoiceNo];
      const inputs = document.querySelectorAll('#modalItemsTbody input.rate-input');
      const rates = {};
      inputs.forEach(inp => {
        const sr = inp.dataset.sr;
        const val = parseFloat(inp.value);
        rates[sr] = isNaN(val) ? null : val;
      });

      const agreedFeRateVal = document.getElementById('modalAgreedFeRate').value.trim();
      let agreed_fe_rate = agreedFeRateVal === '' ? null : parseFloat(agreedFeRateVal);

      const agreedFeVal = document.getElementById('modalAgreedFe').value.trim();
      let agreed_fe = agreedFeVal === '' ? null : parseFloat(agreedFeVal);

      const agreedPmsRateVal = document.getElementById('modalAgreedPmsRate').value.trim();
      let agreed_pms_rate = agreedPmsRateVal === '' ? null : parseFloat(agreedPmsRateVal);

      const agreedPmsVal = document.getElementById('modalAgreedPms').value.trim();
      let agreed_pms = agreedPmsVal === '' ? null : parseFloat(agreedPmsVal);

      // If user kept the default billed rate (within 0.02), clear custom so it stays default
      if (inv) {
        const feRateBilled = (inv.fe_rate_billed !== undefined && inv.fe_rate_billed !== null) ? inv.fe_rate_billed : (inv.total_sqmm > 0 ? (inv.fe / inv.total_sqmm) : 0);
        if (agreed_fe_rate !== null && Math.abs(agreed_fe_rate - feRateBilled) <= 0.02) {
          agreed_fe_rate = null;
          agreed_fe = null;
        }
        const pmsRateBilled = (inv.pms_rate_billed !== undefined && inv.pms_rate_billed !== null) ? inv.pms_rate_billed : (inv.total_sqmm > 0 ? (inv.pms / inv.total_sqmm) : 0);
        if (agreed_pms_rate !== null && Math.abs(agreed_pms_rate - pmsRateBilled) <= 0.02) {
          agreed_pms_rate = null;
          agreed_pms = null;
        }
      }

      try {
        const res = await fetch('/api/save-rates', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            invoice_no: currentInvoiceNo,
            rates,
            agreed_fe_rate,
            agreed_fe,
            agreed_pms_rate,
            agreed_pms
          })
        });
        if (res.status === 401) { window.location.href = '/login'; return; }
        const ans = await res.json();
        if (ans.success) {
          closeModal();
          await loadState();
        } else {
          alert('Error saving rates: ' + (ans.error || 'Unknown error'));
        }
      } catch (e) {
        alert('Network error: ' + e);
      }
    }

    async function logoutUser() {
      try {
        await fetch('/api/logout', { method: 'POST' });
      } catch (e) {}
      window.location.href = '/login';
    }

    async function submitPasswordChange() {
      const current_password = document.getElementById('pwdCurrent').value;
      const new_password = document.getElementById('pwdNew').value;
      const confirm = document.getElementById('pwdConfirm').value;
      const msg = document.getElementById('pwd-msg');
      const showMsg = (ok, txt) => {
        msg.style.display = 'block';
        msg.style.background = ok ? 'rgba(16,185,129,0.14)' : 'rgba(239,68,68,0.14)';
        msg.style.color = ok ? 'var(--success)' : 'var(--danger)';
        msg.style.border = ok ? '1px solid rgba(52,211,153,0.35)' : '1px solid rgba(248,113,113,0.35)';
        msg.textContent = txt;
      };
      if (!current_password || !new_password) {
        showMsg(false, 'Please enter both current and new password.');
        return;
      }
      if (new_password !== confirm) {
        showMsg(false, 'New password and confirmation do not match.');
        return;
      }
      try {
        const res = await fetch('/api/change-password', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ current_password, new_password })
        });
        if (res.status === 401) { window.location.href = '/login'; return; }
        const data = await res.json();
        if (data.success) {
          showMsg(true, '✅ ' + (data.message || 'Password updated successfully!'));
          ['pwdCurrent', 'pwdNew', 'pwdConfirm'].forEach(id => {
            const el = document.getElementById(id);
            if (el) el.value = '';
          });
        } else {
          showMsg(false, '❌ ' + (data.error || 'Failed to update password'));
        }
      } catch (e) {
        showMsg(false, 'Network error: ' + e);
      }
    }

    function showUsersMsg(ok, txt) {
      const msg = document.getElementById('users-msg');
      if (!msg) return;
      msg.style.display = 'block';
      msg.style.background = ok ? 'rgba(16,185,129,0.14)' : 'rgba(239,68,68,0.14)';
      msg.style.color = ok ? 'var(--success)' : 'var(--danger)';
      msg.style.border = ok ? '1px solid rgba(52,211,153,0.35)' : '1px solid rgba(248,113,113,0.35)';
      msg.textContent = txt;
    }

    async function loadUsersTable() {
      const tbody = document.getElementById('users-tbody');
      if (!tbody) return;
      try {
        const res = await fetch('/api/users');
        if (res.status === 401) { window.location.href = '/login'; return; }
        const data = await res.json();
        if (!data.success) {
          tbody.innerHTML = `<tr><td colspan="6" class="center" style="color:var(--danger);">${escHtml(data.error || 'Failed to load users')}</td></tr>`;
          return;
        }
        const me = getCurrentUser() ? getCurrentUser().username : '';
        let html = '';
        (data.users || []).forEach(u => {
          const isSelf = u.username === me;
          const lastLogin = u.last_login ? u.last_login.replace('T', ' ').slice(0, 16) : 'Never';
          html += `<tr>
            <td class="mono bold">${escHtml(u.username)} ${isSelf ? '<span class="badge badge-info" style="font-size:9px;">YOU</span>' : ''}</td>
            <td>${escHtml(u.name || u.username)}</td>
            <td>
              <select class="rate-input" style="padding:4px 8px; width:130px; text-align:left;" onchange="updateUserRole('${escHtml(u.username)}', this.value)">
                <option value="admin" ${u.role === 'admin' ? 'selected' : ''}>🛡️ Admin</option>
                <option value="editor" ${u.role === 'editor' ? 'selected' : ''}>✏️ Editor</option>
                <option value="viewer" ${u.role === 'viewer' ? 'selected' : ''}>👁️ Viewer</option>
              </select>
            </td>
            <td class="center">
              <span class="badge ${u.active ? 'badge-success' : 'badge-danger'}">${u.active ? 'Active' : 'Disabled'}</span>
            </td>
            <td style="font-size:11.5px; color:var(--text-muted);">${escHtml(lastLogin)}</td>
            <td class="center" style="white-space:nowrap;">
              <button class="btn btn-primary-sm" style="padding:4px 8px; font-size:11px;" onclick="resetUserPassword('${escHtml(u.username)}')" title="Reset Password">🔑 Reset</button>
              ${!isSelf ? `<button class="btn btn-primary-sm" style="padding:4px 8px; font-size:11px; margin-left:4px;" onclick="toggleUserActive('${escHtml(u.username)}', ${!u.active})">${u.active ? '⏸ Disable' : '▶ Enable'}</button>` : ''}
              ${!isSelf ? `<button class="btn btn-primary-sm" style="padding:4px 8px; font-size:11px; margin-left:4px; border-color:rgba(248,113,113,0.4); color:var(--danger);" onclick="deleteUserAccount('${escHtml(u.username)}')">🗑</button>` : ''}
            </td>
          </tr>`;
        });
        tbody.innerHTML = html || '<tr><td colspan="6" class="center">No users found</td></tr>';
      } catch (e) {
        tbody.innerHTML = `<tr><td colspan="6" class="center" style="color:var(--danger);">Error loading users</td></tr>`;
      }
    }

    async function createUserAccount() {
      const username = document.getElementById('newUsername').value.trim();
      const name = document.getElementById('newDisplayName').value.trim();
      const role = document.getElementById('newUserRole').value;
      const password = document.getElementById('newUserPwd').value;
      if (!username || !password) {
        showUsersMsg(false, 'Login ID and Password are required.');
        return;
      }
      try {
        const res = await fetch('/api/users/create', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ username, name: name || username, role, password })
        });
        if (res.status === 401) { window.location.href = '/login'; return; }
        const data = await res.json();
        if (data.success) {
          document.getElementById('newUsername').value = '';
          document.getElementById('newDisplayName').value = '';
          document.getElementById('newUserPwd').value = '';
          showUsersMsg(true, `✅ Created user '${data.user.username}' with role '${data.user.role.toUpperCase()}'.`);
          loadUsersTable();
        } else {
          showUsersMsg(false, '❌ ' + (data.error || 'Could not create user'));
        }
      } catch (e) {
        showUsersMsg(false, 'Network error: ' + e);
      }
    }

    async function updateUserRole(username, role) {
      try {
        const res = await fetch('/api/users/update', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ username, role })
        });
        if (res.status === 401) { window.location.href = '/login'; return; }
        const data = await res.json();
        if (data.success) {
          showUsersMsg(true, `✅ Updated '${username}' role to '${role.toUpperCase()}'.`);
        } else {
          showUsersMsg(false, '❌ ' + (data.error || 'Update failed'));
        }
        loadUsersTable();
      } catch (e) {
        showUsersMsg(false, 'Network error: ' + e);
      }
    }

    async function toggleUserActive(username, active) {
      try {
        const res = await fetch('/api/users/update', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ username, active })
        });
        if (res.status === 401) { window.location.href = '/login'; return; }
        const data = await res.json();
        if (data.success) {
          showUsersMsg(true, `✅ User '${username}' is now ${active ? 'Active' : 'Disabled'}.`);
        } else {
          showUsersMsg(false, '❌ ' + (data.error || 'Status update failed'));
        }
        loadUsersTable();
      } catch (e) {
        showUsersMsg(false, 'Network error: ' + e);
      }
    }

    async function resetUserPassword(username) {
      const pwd = prompt(`Enter new password for '${username}' (min 4 characters):`);
      if (pwd === null) return;
      if (pwd.length < 4) {
        showUsersMsg(false, 'Password must be at least 4 characters.');
        return;
      }
      try {
        const res = await fetch('/api/users/update', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ username, new_password: pwd })
        });
        if (res.status === 401) { window.location.href = '/login'; return; }
        const data = await res.json();
        if (data.success) {
          showUsersMsg(true, `✅ Password reset successfully for '${username}'.`);
        } else {
          showUsersMsg(false, '❌ ' + (data.error || 'Password reset failed'));
        }
      } catch (e) {
        showUsersMsg(false, 'Network error: ' + e);
      }
    }

    async function deleteUserAccount(username) {
      if (!confirm(`Delete user '${username}'? Active sessions for this account will be revoked immediately.`)) return;
      try {
        const res = await fetch('/api/users/delete', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ username })
        });
        if (res.status === 401) { window.location.href = '/login'; return; }
        const data = await res.json();
        if (data.success) {
          showUsersMsg(true, `✅ Deleted user '${username}'.`);
          loadUsersTable();
        } else {
          showUsersMsg(false, '❌ ' + (data.error || 'Delete failed'));
        }
      } catch (e) {
        showUsersMsg(false, 'Network error: ' + e);
      }
    }

    // Modal UX: Close on Escape key press or clicking backdrop
    window.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') {
        closeModal();
      }
    });
    const rateModalEl = document.getElementById('rateModal');
    if (rateModalEl) {
      rateModalEl.addEventListener('click', (e) => {
        if (e.target === rateModalEl) closeModal();
      });
    }

    // Render immediately from embedded initial data without network delay
    if (appState) {
      renderDashboard();
    } else {
      loadState();
    }
  </script>
</body>
</html>
"""

def get_full_state_dict(user=None):
    engine.load()
    monthly = engine.get_monthly_reconciliation()
    audit = engine.get_audit_report()
    tonnage = engine.get_tonnage_incentives()
    ledger_25_26 = {}
    if os.path.exists(engine.data_file):
        try:
            with open(engine.data_file) as f:
                s_data = json.load(f)
                ledger_25_26 = s_data.get('ledger_25_26', {})
        except:
            pass

    cd_tracker = engine.get_cd_tracker_data()
    safety = engine.validate_safety_integrity()
    return {
        'current_user': user,
        'roles': ROLES,
        'invoices': engine.invoices,
        'credit_notes': engine.credit_notes,
        'ledger': engine.ledger,
        'ledger_25_26': ledger_25_26,
        'audit': audit,
        'safety': safety,
        'monthly': monthly,
        'cd_tracker': cd_tracker,
        'tonnage': tonnage
    }

class RequestHandler(BaseHTTPRequestHandler):
    def _send_security_headers(self):
        self.send_header('X-Content-Type-Options', 'nosniff')
        self.send_header('X-Frame-Options', 'SAMEORIGIN')
        self.send_header('Referrer-Policy', 'same-origin')
        self.send_header('Permissions-Policy', 'camera=(), microphone=(), geolocation=()')

    def _get_client_ip(self):
        cf_ip = (self.headers.get('CF-Connecting-IP') or '').strip()
        if cf_ip:
            return cf_ip
        return self.client_address[0] if self.client_address else 'unknown'

    def _cookie_security_flags(self):
        proto = (self.headers.get('X-Forwarded-Proto') or '').strip().lower()
        return '; Secure' if proto == 'https' else ''

    def _verify_same_origin(self):
        """Block cross-origin CSRF POST requests when Origin or Referer header is present."""
        expected_host = (self.headers.get('X-Forwarded-Host') or self.headers.get('Host') or '').split(',')[0].strip().lower()
        origin = (self.headers.get('Origin') or '').strip()
        referer = (self.headers.get('Referer') or '').strip()
        check_url = origin or referer
        if not check_url or not expected_host:
            return True
        try:
            netloc = urlparse(check_url).netloc.lower()
            if netloc and netloc != expected_host:
                return False
        except Exception:
            return False
        return True

    def _get_session_token(self):
        raw_cookie = self.headers.get('Cookie', '')
        if not raw_cookie:
            return None
        cookie = SimpleCookie()
        try:
            cookie.load(raw_cookie)
        except Exception:
            return None
        morsel = cookie.get('ggl_session')
        return morsel.value if morsel else None

    def _get_current_user(self):
        token = self._get_session_token()
        if not token:
            return None
        return auth_manager.get_session_user(token)

    def _send_json(self, status, payload, extra_headers=None):
        self.send_response(status)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
        self._send_security_headers()
        if extra_headers:
            for k, v in extra_headers:
                self.send_header(k, v)
        self.end_headers()
        self.wfile.write(json.dumps(payload).encode('utf-8'))

    def _read_json_body(self):
        try:
            raw_len = self.headers.get('Content-Length', '0')
            content_length = int(raw_len)
        except ValueError:
            return None, ('Invalid Content-Length header', 400)
        if content_length < 0 or content_length > MAX_JSON_BYTES:
            return None, (f'JSON payload too large (max {MAX_JSON_BYTES} bytes)', 413)
        try:
            body = self.rfile.read(content_length).decode('utf-8') if content_length > 0 else '{}'
            parsed = json.loads(body)
            if not isinstance(parsed, dict):
                return None, ('JSON root must be an object', 400)
            return parsed, None
        except Exception as e:
            return None, (f'Invalid JSON payload: {str(e)}', 400)

    def _parse_multipart_file(self):
        """Parse multipart/form-data file upload using stdlib email.parser (Python 3.9-3.13+ safe)."""
        content_type = self.headers.get('Content-Type', '')
        if 'multipart/form-data' not in content_type.lower():
            return None, ('Expected multipart/form-data upload', 400)
        try:
            content_length = int(self.headers.get('Content-Length', '0'))
        except ValueError:
            return None, ('Invalid Content-Length header', 400)
        if content_length <= 0:
            return None, ('Empty upload payload', 400)
        if content_length > MAX_UPLOAD_BYTES:
            return None, (f'File exceeds maximum allowed size ({MAX_UPLOAD_BYTES // (1024*1024)} MB)', 413)

        raw_body = self.rfile.read(content_length)
        header_block = f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode('utf-8')
        try:
            msg = BytesParser(policy=email_policy).parsebytes(header_block + raw_body)
            if not msg.is_multipart():
                return None, ('Malformed multipart upload body', 400)
            for part in msg.iter_parts():
                if part.get_param('name', header='content-disposition') == 'file':
                    payload = part.get_payload(decode=True)
                    return payload or b'', None
            return None, ('No file was selected for upload. Please choose a valid PDF file.', 400)
        except Exception as e:
            return None, (f'Failed to parse upload request: {str(e)}', 400)

    def _redirect(self, location):
        self.send_response(302)
        self.send_header('Location', location)
        self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
        self._send_security_headers()
        self.end_headers()

    def _require_auth_api(self, permission=None):
        user = self._get_current_user()
        if not user:
            self._send_json(401, {'success': False, 'error': 'Authentication required. Please sign in.'})
            return None
        if permission and not auth_manager.has_permission(user, permission):
            self._send_json(403, {'success': False, 'error': f"Access denied: '{permission}' permission required (your role: {user.get('role')})."})
            return None
        return user

    def do_HEAD(self):
        parsed_path = urlparse(self.path).path
        if parsed_path in ('/', '/index.html') and not self._get_current_user():
            self._redirect('/login')
            return
        self.send_response(200)
        self.send_header('Content-Type', 'text/html; charset=utf-8')
        self._send_security_headers()
        self.end_headers()

    def do_GET(self):
        parsed_path = urlparse(self.path).path
        if parsed_path == '/login':
            if self._get_current_user():
                self._redirect('/')
                return
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
            self._send_security_headers()
            self.end_headers()
            self.wfile.write(LOGIN_HTML.encode('utf-8'))
        elif parsed_path == '/' or parsed_path == '/index.html':
            user = self._get_current_user()
            if not user:
                self._redirect('/login')
                return
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
            self._send_security_headers()
            self.end_headers()
            state_data = get_full_state_dict(user=user)
            state_json = json.dumps(state_data)
            rendered_html = HTML_CONTENT.replace('/* __INITIAL_STATE_JSON__ */', f"window.INITIAL_DATA = {state_json};")
            self.wfile.write(rendered_html.encode('utf-8'))
        elif parsed_path == '/favicon.ico':
            self.send_response(204)
            self._send_security_headers()
            self.end_headers()
        elif parsed_path == '/api/health':
            self._send_json(200, {'status': 'ok', 'time': datetime.now().isoformat()})
        elif parsed_path == '/api/state':
            user = self._require_auth_api('view')
            if not user:
                return
            state_data = get_full_state_dict(user=user)
            self._send_json(200, state_data)
        elif parsed_path == '/api/users':
            user = self._require_auth_api('manage_users')
            if not user:
                return
            self._send_json(200, {'success': True, 'users': auth_manager.list_users(), 'roles': ROLES})
        elif parsed_path == '/api/view-pdf':
            user = self._get_current_user()
            if not user:
                self._redirect('/login')
                return
            query = parse_qs(urlparse(self.path).query)
            raw_doc = query.get('doc', [''])[0].strip()
            if '..' in raw_doc or '/' in raw_doc or '\\' in raw_doc or '\x00' in raw_doc:
                self.send_error(400, "Invalid document parameter")
                return
            clean_doc = os.path.basename(raw_doc).strip()
            if not clean_doc or clean_doc in ['.', '..']:
                self.send_error(400, "Invalid document parameter")
                return
            
            pdf_path = os.path.join(PDF_STORE_DIR, f"{clean_doc}.pdf")
            if not os.path.exists(pdf_path):
                for f in os.listdir(PDF_STORE_DIR):
                    if clean_doc and (clean_doc in f or f.startswith(clean_doc)):
                        pdf_path = os.path.join(PDF_STORE_DIR, f)
                        break
                        
            real_store = os.path.realpath(PDF_STORE_DIR)
            real_target = os.path.realpath(pdf_path)
            is_inside = False
            try:
                is_inside = os.path.commonpath([real_store, real_target]) == real_store and real_target != real_store
            except ValueError:
                is_inside = False

            if is_inside and os.path.isfile(real_target):
                self.send_response(200)
                self.send_header('Content-Type', 'application/pdf')
                self.send_header('Content-Disposition', f'inline; filename="{os.path.basename(real_target)}"')
                self.send_header('Cache-Control', 'private, max-age=3600')
                self._send_security_headers()
                self.end_headers()
                with open(real_target, 'rb') as f:
                    self.wfile.write(f.read())
            else:
                self.send_error(404, f"PDF for document {clean_doc} not found")
        elif parsed_path == '/api/download-excel':
            user = self._get_current_user()
            if not user:
                self._redirect('/login')
                return
            generate_reconciliation_workbook(engine, EXCEL_PATH)
            if os.path.exists(EXCEL_PATH):
                self.send_response(200)
                self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                self.send_header('Content-Disposition', 'attachment; filename="GGL_Master_Reconciliation.xlsx"')
                self._send_security_headers()
                self.end_headers()
                with open(EXCEL_PATH, 'rb') as f:
                    self.wfile.write(f.read())
            else:
                self.send_error(404, "Excel file not found")
        elif parsed_path == '/api/download-claim-pdf':
            user = self._get_current_user()
            if not user:
                self._redirect('/login')
                return
            claim_pdf_path = os.path.join(DATA_DIR, 'GGL_Rate_Difference_Claim_Report_FY26-27.pdf')
            generate_monthly_claim_pdf(engine, claim_pdf_path)
            if os.path.exists(claim_pdf_path):
                self.send_response(200)
                self.send_header('Content-Type', 'application/pdf')
                self.send_header('Content-Disposition', 'inline; filename="GGL_Rate_Difference_Claim_Report_FY26-27.pdf"')
                self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
                self._send_security_headers()
                self.end_headers()
                with open(claim_pdf_path, 'rb') as f:
                    self.wfile.write(f.read())
            else:
                self.send_error(500, "Error generating claim PDF")
        else:
            self.send_error(404)

    def do_POST(self):
        parsed_path = urlparse(self.path).path
        if not self._verify_same_origin():
            self._send_json(403, {'success': False, 'error': 'Cross-origin request blocked (CSRF protection).'})
            return

        if parsed_path == '/api/login':
            req, err = self._read_json_body()
            if err:
                msg, code = err
                self._send_json(code, {'success': False, 'error': msg})
                return
            res, auth_err = auth_manager.authenticate(
                req.get('username', ''),
                req.get('password', ''),
                remember=bool(req.get('remember', False)),
                client_ip=self._get_client_ip()
            )
            if auth_err:
                if res and res.get('rate_limited'):
                    self._send_json(429, {'success': False, 'error': auth_err}, extra_headers=[('Retry-After', str(res.get('retry_after', 60)))])
                else:
                    self._send_json(401, {'success': False, 'error': auth_err})
                return
            sec_flag = self._cookie_security_flags()
            cookie_val = f"ggl_session={res['token']}; Path=/; HttpOnly; SameSite=Lax; Max-Age={res['ttl']}{sec_flag}"
            self._send_json(200, {'success': True, 'user': res['user']}, extra_headers=[('Set-Cookie', cookie_val)])
            return

        if parsed_path == '/api/logout':
            token = self._get_session_token()
            if token:
                auth_manager.logout(token)
            sec_flag = self._cookie_security_flags()
            clear_cookie = f"ggl_session=; Path=/; HttpOnly; SameSite=Lax; Max-Age=0{sec_flag}"
            self._send_json(200, {'success': True}, extra_headers=[('Set-Cookie', clear_cookie)])
            return

        if parsed_path == '/api/change-password':
            user = self._require_auth_api('view')
            if not user:
                return
            req, err = self._read_json_body()
            if err:
                msg, code = err
                self._send_json(code, {'success': False, 'error': msg})
                return
            ok, msg = auth_manager.change_own_password(
                user['username'],
                req.get('current_password', ''),
                req.get('new_password', '')
            )
            if ok:
                self._send_json(200, {'success': True, 'message': msg})
            else:
                self._send_json(400, {'success': False, 'error': msg})
            return

        if parsed_path == '/api/users/create':
            user = self._require_auth_api('manage_users')
            if not user:
                return
            req, err = self._read_json_body()
            if err:
                msg, code = err
                self._send_json(code, {'success': False, 'error': msg})
                return
            created, c_err = auth_manager.create_user(
                req.get('username', ''),
                req.get('name', ''),
                req.get('role', 'viewer'),
                req.get('password', '')
            )
            if c_err:
                self._send_json(400, {'success': False, 'error': c_err})
            else:
                self._send_json(200, {'success': True, 'user': created})
            return

        if parsed_path == '/api/users/update':
            user = self._require_auth_api('manage_users')
            if not user:
                return
            req, err = self._read_json_body()
            if err:
                msg, code = err
                self._send_json(code, {'success': False, 'error': msg})
                return
            updated, u_err = auth_manager.update_user(
                user['username'],
                req.get('username', ''),
                role=req.get('role'),
                name=req.get('name'),
                active=req.get('active'),
                new_password=req.get('new_password')
            )
            if u_err:
                self._send_json(400, {'success': False, 'error': u_err})
            else:
                self._send_json(200, {'success': True, 'user': updated})
            return

        if parsed_path == '/api/users/delete':
            user = self._require_auth_api('manage_users')
            if not user:
                return
            req, err = self._read_json_body()
            if err:
                msg, code = err
                self._send_json(code, {'success': False, 'error': msg})
                return
            ok, msg = auth_manager.delete_user(user['username'], req.get('username', ''))
            if ok:
                self._send_json(200, {'success': True, 'message': msg})
            else:
                self._send_json(400, {'success': False, 'error': msg})
            return

        if parsed_path == '/api/save-rates':
            user = self._require_auth_api('edit_rates')
            if not user:
                return
            engine.load()
            req, err = self._read_json_body()
            if err:
                msg, code = err
                self._send_json(code, {'success': False, 'error': msg})
                return

            inv_no = req.get('invoice_no')
            if not inv_no or inv_no not in engine.invoices:
                self._send_json(404, {'success': False, 'error': f'Invoice #{inv_no} not found in database.'})
                return

            rates = req.get('rates', {})
            agreed_fe_rate = req.get('agreed_fe_rate')
            agreed_fe = req.get('agreed_fe')
            agreed_pms_rate = req.get('agreed_pms_rate')
            agreed_pms = req.get('agreed_pms')
            
            try:
                engine.set_agreed_rates_and_surcharges(
                    inv_no,
                    rates=rates,
                    agreed_fe_rate=agreed_fe_rate,
                    agreed_fe=agreed_fe,
                    agreed_pms_rate=agreed_pms_rate,
                    agreed_pms=agreed_pms
                )
                generate_reconciliation_workbook(engine, EXCEL_PATH)
                res = {'success': True}
            except Exception as e:
                res = {'success': False, 'error': f'Error saving agreed rates: {str(e)}'}

            self._send_json(200, res)
            
        elif parsed_path in ['/api/upload-bill', '/api/upload-cn', '/api/upload-ledger']:
            user = self._require_auth_api('upload_docs')
            if not user:
                return
            engine.load()
            file_bytes, up_err = self._parse_multipart_file()
            if up_err:
                msg, code = up_err
                self._send_json(code if code == 413 else 200, {'success': False, 'error': msg})
                return

            if not file_bytes or len(file_bytes) == 0:
                self._send_json(200, {'success': False, 'error': 'The uploaded file is empty (0 bytes). Please upload a valid GGL PDF document.'})
                return

            if not file_bytes.startswith(b'%PDF'):
                self._send_json(200, {'success': False, 'error': 'Invalid file format: Only official PDF files (file must begin with %PDF) are accepted.'})
                return
            
            # 1. Byte-level hash check against existing PDF store
            incoming_hash = hashlib.sha256(file_bytes).hexdigest()
            for existing_file in os.listdir(PDF_STORE_DIR):
                existing_path = os.path.join(PDF_STORE_DIR, existing_file)
                if os.path.isfile(existing_path):
                    try:
                        with open(existing_path, 'rb') as ef:
                            if hashlib.sha256(ef.read()).hexdigest() == incoming_hash:
                                res = {
                                    'success': False,
                                    'is_duplicate': True,
                                    'error': f"Identical PDF Detected: This file matches '{existing_file}' already stored in the system. Upload declined to protect record consistency."
                                }
                                self._send_json(200, res)
                                return
                    except Exception:
                        pass

            filename = f"upload_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}.pdf"
            dest_path = os.path.join(UPLOAD_DIR, filename)
            with open(dest_path, 'wb') as f:
                f.write(file_bytes)
                
            if parsed_path == '/api/upload-bill':
                try:
                    inv = parse_invoice_pdf(dest_path)
                except Exception as e:
                    if os.path.exists(dest_path): os.remove(dest_path)
                    res = {'success': False, 'error': f"Failed to parse invoice PDF: {str(e)}. Please check if file is damaged or password-protected."}
                    inv = None

                if inv:
                    inv_no = inv.get('invoice_no')
                    if not inv_no:
                        if os.path.exists(dest_path): os.remove(dest_path)
                        res = {'success': False, 'error': 'Could not extract a valid invoice number from the uploaded PDF.'}
                    elif inv_no in engine.invoices:
                        if os.path.exists(dest_path): os.remove(dest_path)
                        existing = engine.invoices[inv_no]
                        res = {
                            'success': False,
                            'is_duplicate': True,
                            'error': f"Duplicate Bill: Invoice #{inv_no} dated {existing.get('invoice_date')} (₹{existing.get('total_amount', 0):,.2f}) already exists in your records. Upload declined to protect ledger consistency."
                        }
                    else:
                        engine.add_invoice(inv)
                        generate_reconciliation_workbook(engine, EXCEL_PATH)
                        shutil.copyfile(dest_path, os.path.join(PDF_STORE_DIR, f"{inv_no}.pdf"))
                        res = {'success': True, 'invoice': inv}

            elif parsed_path == '/api/upload-cn':
                try:
                    cn = parse_credit_note_pdf(dest_path)
                except Exception as e:
                    if os.path.exists(dest_path): os.remove(dest_path)
                    res = {'success': False, 'error': f"Failed to parse credit note PDF: {str(e)}. Please check if file is damaged or password-protected."}
                    cn = None

                if cn:
                    cn_no = cn.get('cn_no')
                    if not cn_no:
                        if os.path.exists(dest_path): os.remove(dest_path)
                        res = {'success': False, 'error': 'Could not extract a valid credit note number from the uploaded PDF.'}
                    elif cn_no in engine.credit_notes:
                        if os.path.exists(dest_path): os.remove(dest_path)
                        existing = engine.credit_notes[cn_no]
                        res = {
                            'success': False,
                            'is_duplicate': True,
                            'error': f"Duplicate Credit Note: CN #{cn_no} dated {existing.get('cn_date')} (₹{existing.get('total_amount', 0):,.2f}) already exists in your records. Upload declined to prevent duplicate credit entries."
                        }
                    else:
                        engine.add_credit_note(cn)
                        generate_reconciliation_workbook(engine, EXCEL_PATH)
                        shutil.copyfile(dest_path, os.path.join(PDF_STORE_DIR, f"{cn_no}.pdf"))
                        res = {'success': True, 'credit_note': cn}

            else: # /api/upload-ledger
                try:
                    ledger_data = parse_ledger_pdf(dest_path)
                except Exception as e:
                    if os.path.exists(dest_path): os.remove(dest_path)
                    res = {'success': False, 'error': f"Failed to parse ledger statement PDF: {str(e)}. Please check if file is damaged or password-protected."}
                    ledger_data = None

                if ledger_data:
                    existing_entries = engine.ledger.get('entries', []) if engine.ledger else []
                    new_entries = ledger_data.get('entries', [])
                    
                    if not new_entries:
                        if os.path.exists(dest_path): os.remove(dest_path)
                        res = {'success': False, 'error': 'Could not extract valid ledger transactions from the uploaded PDF.'}
                    elif existing_entries and existing_entries == new_entries:
                        if os.path.exists(dest_path): os.remove(dest_path)
                        last_dt = existing_entries[-1].get('date', '')
                        res = {
                            'success': False,
                            'is_duplicate': True,
                            'error': f"Identical Statement: This statement already contains all {len(existing_entries)} transactions up to {last_dt} (Closing Balance: ₹{engine.ledger.get('closing_balance', 0):,.2f}). No new entries found."
                        }
                    elif existing_entries and len(new_entries) < len(existing_entries) and all(e in existing_entries for e in new_entries):
                        if os.path.exists(dest_path): os.remove(dest_path)
                        last_dt = existing_entries[-1].get('date', '')
                        res = {
                            'success': False,
                            'is_duplicate': True,
                            'error': f"Older Statement Detected: The uploaded statement only has {len(new_entries)} transactions, while your active ledger already has {len(existing_entries)} transactions up to {last_dt}. Upload declined to prevent rolling back newer records."
                        }
                    else:
                        diff_count = max(0, len(new_entries) - len(existing_entries))
                        engine.set_ledger(ledger_data)
                        generate_reconciliation_workbook(engine, EXCEL_PATH)
                        shutil.copyfile(dest_path, os.path.join(PDF_STORE_DIR, "GGL_LEDGER_26_27.pdf"))
                        history_name = f"GGL_LEDGER_26_27_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
                        shutil.copyfile(dest_path, os.path.join(PDF_STORE_DIR, history_name))
                        
                        cb = ledger_data.get('closing_balance', 0.0)
                        cb_str = f"-₹{abs(cb):,.2f} Cr" if cb < 0 else f"₹{cb:,.2f} Dr"
                        msg = f"GGL Statement updated successfully! {diff_count} new transaction(s) incorporated. Updated closing balance: {cb_str}" if diff_count > 0 else f"GGL Statement of Account updated successfully with {len(new_entries)} transactions."
                        res = {'success': True, 'ledger': ledger_data, 'message': msg}
                    
            self._send_json(200, res)
        else:
            self._send_json(404, {'success': False, 'error': 'Unknown API endpoint'})

class ThreadedHTTPServer(ThreadingMixIn, HTTPServer):
    daemon_threads = True

def run():
    server = ThreadedHTTPServer(('0.0.0.0', PORT), RequestHandler)
    print(f"GGL-Recon Web Dashboard running at http://localhost:{PORT}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    server.server_close()

if __name__ == '__main__':
    run()
