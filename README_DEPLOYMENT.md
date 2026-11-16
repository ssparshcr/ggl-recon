# 🚀 24/7 Global Deployment Guide for GGL-Recon

To share a single permanent HTTPS link with anyone in the world that stays online **24 hours a day (even when your Mac is asleep or powered off)**, deploy this folder (`ggl_recon_app/`) to a cloud host or attach a **Named Cloudflare Tunnel** on an always-on machine.

---

## Option 1: Permanent 24/7 Cloud Hosting (Recommended — Zero Mac Dependency)

Everything in this folder (`requirements.txt`, `Dockerfile`, `Procfile`, `render.yaml`, and Python 3.9–3.13+ compatibility) is already configured for 1-click cloud deployment.

### Deploy on Render.com (Free Tier or Persistent Disk)
1. Push the `ggl_recon_app` folder to a **private** GitHub repository (important: keep it private because `data/` holds your invoices, ledger, and PDFs).
2. Sign in to [https://render.com](https://render.com) -> click **New +** -> **Web Service** -> connect your private repo.
3. Render automatically reads `render.yaml` (`pip install -r requirements.txt` and `python server.py`, health check `/api/health`).
4. Optional for persistent saved rate edits/uploads across redeploys: attach a Render Disk mounted at `/var/data` and set environment variable `DATA_DIR=/var/data` (copy `data/` into it on first deploy).
5. You get a permanent 24/7 global URL such as `https://ggl-recon-dashboard.onrender.com`.

*(Note on Render Free Plan: Free services spin down after 15 minutes of no traffic and wake up in ~30 seconds on the next visit. You can add a free external HTTP monitor on `/api/health` every 10 minutes or upgrade to Render Starter (\$7/mo) for zero cold-start latency.)*

### Deploy on Railway.app (Always-On Docker)
1. Push `ggl_recon_app` to a private GitHub repo.
2. Open [https://railway.app](https://railway.app) -> **New Project** -> **Deploy from GitHub repo** (auto-detects `Dockerfile`).
3. Attach a Persistent Volume mounted at `/var/data` and set `DATA_DIR=/var/data`.
4. Under **Settings -> Networking**, click **Generate Domain** for your permanent `https://...up.railway.app` URL.

---

## Option 2: Permanent URL from Your Own Mac / Office PC (Named Cloudflare Tunnel)

Quick `trycloudflare.com` links change URLs whenever restarted and stop if your Mac sleeps. To get a **fixed free subdomain** (`https://audit.yourdomain.com`) from an always-on computer:
1. Run `cloudflared tunnel login` (one-time browser auth with a free Cloudflare account + domain).
2. Create a named tunnel: `cloudflared tunnel create ggl-recon`.
3. Route DNS: `cloudflared tunnel route dns ggl-recon audit.yourdomain.com`.
4. Run `cloudflared tunnel run --url http://127.0.0.1:8085 ggl-recon` alongside `./start_24x7_local.sh`.
5. Use macOS **System Settings -> Battery / Energy -> Prevent automatic sleeping when display is off** (or run `caffeinate -s ./start_24x7_local.sh` while plugged in).
