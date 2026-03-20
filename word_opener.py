#!/usr/bin/env python3
"""
Word Opener - Lightweight DOCX previewer with OneDrive/Word Online integration

Usage:
    python word_opener.py [file.docx]
    python word_opener.py --setup      # Configure OneDrive client ID

Setup (first time for OneDrive):
    1. Go to https://portal.azure.com > App registrations > New registration
    2. Name it anything, select "Personal Microsoft accounts only"
    3. Under "Redirect URIs" add: http://localhost (type: Public client/native)
    4. Under "API permissions" add: Microsoft Graph > Delegated > Files.ReadWrite
    5. Copy the "Application (client) ID"
    6. Run: python word_opener.py --setup
"""

import os
import sys
import json
import webbrowser
import threading
import time
import argparse
import tempfile
from pathlib import Path

# ── dependency checks ────────────────────────────────────────────────────────

def _require(pkg, pip_name=None):
    try:
        return __import__(pkg)
    except ImportError:
        print(f"Missing dependency: pip install {pip_name or pkg}")
        sys.exit(1)

mammoth  = _require("mammoth")
flask    = _require("flask")
msal_mod = _require("msal")
requests = _require("requests")

from flask import Flask, render_template_string, request, jsonify
import msal

# ── paths & constants ────────────────────────────────────────────────────────

CONFIG_DIR        = Path.home() / ".word-opener"
CONFIG_FILE       = CONFIG_DIR / "config.json"
TOKEN_CACHE_FILE  = CONFIG_DIR / "token_cache.json"

SCOPES        = ["Files.ReadWrite", "offline_access"]
AUTHORITY     = "https://login.microsoftonline.com/consumers"
GRAPH_BASE    = "https://graph.microsoft.com/v1.0"

# ── app state ────────────────────────────────────────────────────────────────

_state = {
    "docx_path":    None,
    "html_content": None,
    "msal_app":     None,
    "token_cache":  msal.SerializableTokenCache(),
    "device_flow":  None,
}

app = Flask(__name__)
app.secret_key = os.urandom(24)

# ── config helpers ───────────────────────────────────────────────────────────

def load_config():
    if CONFIG_FILE.exists():
        with open(CONFIG_FILE) as f:
            return json.load(f)
    return {}

def save_config(cfg):
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    with open(CONFIG_FILE, "w") as f:
        json.dump(cfg, f, indent=2)

def _persist_cache():
    if _state["token_cache"].has_state_changed:
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        with open(TOKEN_CACHE_FILE, "w") as f:
            f.write(_state["token_cache"].serialize())

def _load_cache():
    if TOKEN_CACHE_FILE.exists():
        with open(TOKEN_CACHE_FILE) as f:
            _state["token_cache"].deserialize(f.read())

def get_msal_app():
    if _state["msal_app"] is None:
        cfg = load_config()
        client_id = cfg.get("client_id", "").strip()
        if not client_id:
            return None
        _load_cache()
        _state["msal_app"] = msal.PublicClientApplication(
            client_id,
            authority=AUTHORITY,
            token_cache=_state["token_cache"],
        )
    return _state["msal_app"]

def get_token_silent():
    app_obj = get_msal_app()
    if not app_obj:
        return None
    accounts = app_obj.get_accounts()
    if accounts:
        result = app_obj.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            _persist_cache()
            return result["access_token"]
    return None

# ── HTML template ────────────────────────────────────────────────────────────

TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Word Opener{% if filename %} – {{ filename }}{% endif %}</title>
  <style>
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

    body {
      font-family: 'Segoe UI', system-ui, sans-serif;
      background: #f3f4f6;
      color: #1f2937;
      min-height: 100vh;
    }

    /* ── toolbar ── */
    #toolbar {
      position: sticky; top: 0; z-index: 100;
      display: flex; align-items: center; gap: 12px;
      padding: 10px 20px;
      background: #1e3a5f;
      color: #fff;
      box-shadow: 0 2px 8px rgba(0,0,0,.25);
    }
    #toolbar .logo { font-size: 1.25rem; font-weight: 700; letter-spacing: -.5px; flex-shrink: 0; }
    #toolbar .filename {
      flex: 1; font-size: .9rem; opacity: .8;
      white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
    }
    #toolbar .actions { display: flex; gap: 8px; flex-shrink: 0; }

    .btn {
      display: inline-flex; align-items: center; gap: 6px;
      padding: 7px 14px; border-radius: 6px; border: none;
      font-size: .85rem; font-weight: 600; cursor: pointer;
      transition: background .15s, opacity .15s;
    }
    .btn:disabled { opacity: .5; cursor: not-allowed; }
    .btn-primary   { background: #2563eb; color: #fff; }
    .btn-primary:hover:not(:disabled) { background: #1d4ed8; }
    .btn-onedrive  { background: #0078d4; color: #fff; }
    .btn-onedrive:hover:not(:disabled) { background: #005a9e; }
    .btn-ghost     { background: rgba(255,255,255,.15); color: #fff; border: 1px solid rgba(255,255,255,.3); }
    .btn-ghost:hover:not(:disabled) { background: rgba(255,255,255,.25); }

    /* ── drop zone ── */
    #drop-zone {
      display: flex; flex-direction: column; align-items: center; justify-content: center;
      gap: 16px; min-height: calc(100vh - 60px);
      padding: 40px;
    }
    #drop-zone.has-file { display: none; }
    #drop-zone .icon { font-size: 4rem; }
    #drop-zone h2 { font-size: 1.4rem; color: #374151; }
    #drop-zone p  { color: #6b7280; font-size: .95rem; }
    #drop-zone.dragover { outline: 3px dashed #2563eb; outline-offset: -12px; background: #eff6ff; }

    #file-input { display: none; }

    /* ── document container ── */
    #doc-wrap {
      display: none;
      max-width: 860px; margin: 32px auto; padding: 0 20px 60px;
    }
    #doc-wrap.visible { display: block; }

    #doc-content {
      background: #fff;
      padding: 60px 72px;
      border-radius: 4px;
      box-shadow: 0 1px 4px rgba(0,0,0,.12), 0 4px 16px rgba(0,0,0,.08);
      min-height: 400px;
      line-height: 1.7;
    }

    /* DOCX content styles */
    #doc-content h1,#doc-content h2,#doc-content h3,
    #doc-content h4,#doc-content h5,#doc-content h6 {
      margin: 1.2em 0 .4em; line-height: 1.25;
    }
    #doc-content h1 { font-size: 2rem; }
    #doc-content h2 { font-size: 1.5rem; border-bottom: 1px solid #e5e7eb; padding-bottom: 6px; }
    #doc-content h3 { font-size: 1.25rem; }
    #doc-content p  { margin-bottom: .85em; }
    #doc-content ul,#doc-content ol { margin: .5em 0 .85em 1.5em; }
    #doc-content li { margin-bottom: .3em; }
    #doc-content table {
      border-collapse: collapse; width: 100%; margin: 1em 0; font-size: .9rem;
    }
    #doc-content th,#doc-content td {
      border: 1px solid #d1d5db; padding: 8px 12px; text-align: left;
    }
    #doc-content th { background: #f9fafb; font-weight: 600; }
    #doc-content img { max-width: 100%; height: auto; margin: 8px 0; }
    #doc-content strong { font-weight: 700; }
    #doc-content em    { font-style: italic; }
    #doc-content blockquote {
      border-left: 3px solid #d1d5db; margin: 1em 0;
      padding: .5em 1em; color: #6b7280;
    }

    /* ── modal ── */
    .modal-backdrop {
      display: none; position: fixed; inset: 0;
      background: rgba(0,0,0,.5); z-index: 200;
      align-items: center; justify-content: center;
    }
    .modal-backdrop.open { display: flex; }
    .modal {
      background: #fff; border-radius: 12px;
      padding: 32px; max-width: 440px; width: 90%;
      box-shadow: 0 20px 60px rgba(0,0,0,.3);
      text-align: center;
    }
    .modal h3 { font-size: 1.2rem; margin-bottom: 12px; }
    .modal p  { color: #6b7280; font-size: .9rem; margin-bottom: 20px; }

    .device-code-box {
      background: #f3f4f6; border-radius: 8px;
      padding: 18px; margin: 16px 0;
    }
    .device-code-box .code {
      font-family: monospace; font-size: 2rem; font-weight: 700;
      letter-spacing: .15em; color: #1e3a5f;
    }
    .device-code-box .url { font-size: .8rem; color: #6b7280; margin-top: 6px; }

    .spinner {
      width: 24px; height: 24px;
      border: 3px solid #e5e7eb;
      border-top-color: #2563eb;
      border-radius: 50%;
      animation: spin .7s linear infinite;
      margin: 0 auto 12px;
    }
    @keyframes spin { to { transform: rotate(360deg); } }

    .status-msg { font-size: .85rem; color: #6b7280; margin-top: 8px; }
    .error-msg  { font-size: .85rem; color: #dc2626; margin-top: 8px; }

    @media (max-width: 640px) {
      #doc-content { padding: 32px 24px; }
    }
  </style>
</head>
<body>

<!-- toolbar -->
<div id="toolbar">
  <div class="logo">📄 Word Opener</div>
  <div class="filename" id="toolbar-filename">{{ filename or 'No file loaded' }}</div>
  <div class="actions">
    <button class="btn btn-ghost" id="btn-open">Open file</button>
    <button class="btn btn-onedrive" id="btn-onedrive" {% if not filename %}disabled{% endif %}>
      ☁ Open in OneDrive
    </button>
  </div>
</div>
<input type="file" id="file-input" accept=".docx">

<!-- drop zone (shown when no file loaded) -->
<div id="drop-zone" {% if filename %}{% else %}{% endif %}>
  <div class="icon">📂</div>
  <h2>Open a DOCX file</h2>
  <p>Drag &amp; drop a file here, or click <strong>Open file</strong> above.</p>
  <button class="btn btn-primary" onclick="document.getElementById('file-input').click()">
    Choose file…
  </button>
</div>

<!-- document view -->
<div id="doc-wrap" {% if html_content %}class="visible"{% endif %}>
  <div id="doc-content">{{ html_content|safe if html_content else '' }}</div>
</div>

<!-- auth/upload modal -->
<div class="modal-backdrop" id="modal">
  <div class="modal">
    <div id="modal-body"><!-- filled dynamically --></div>
    <button class="btn btn-ghost" id="modal-cancel"
            style="color:#374151;border-color:#d1d5db;margin-top:8px"
            onclick="closeModal()">Cancel</button>
  </div>
</div>

<script>
const dropZone   = document.getElementById('drop-zone');
const docWrap    = document.getElementById('doc-wrap');
const docContent = document.getElementById('doc-content');
const fileInput  = document.getElementById('file-input');
const modal      = document.getElementById('modal');
const modalBody  = document.getElementById('modal-body');
const btnOneDrive = document.getElementById('btn-onedrive');
const toolbarFilename = document.getElementById('toolbar-filename');

{% if filename %}
dropZone.classList.add('has-file');
{% endif %}

// ── file loading ────────────────────────────────────────────────────────────

fileInput.addEventListener('change', e => {
  const file = e.target.files[0];
  if (file) uploadForPreview(file);
});

document.getElementById('btn-open').addEventListener('click', () => fileInput.click());

dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('dragover'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('dragover');
  const file = e.dataTransfer.files[0];
  if (file && file.name.endsWith('.docx')) uploadForPreview(file);
  else alert('Please drop a .docx file.');
});

async function uploadForPreview(file) {
  const fd = new FormData();
  fd.append('file', file);
  showModal(spinnerHTML('Converting document…'));
  try {
    const res  = await fetch('/preview', { method: 'POST', body: fd });
    const data = await res.json();
    if (data.error) { showModal(errorHTML(data.error)); return; }
    docContent.innerHTML = data.html;
    docWrap.classList.add('visible');
    dropZone.classList.add('has-file');
    toolbarFilename.textContent = data.filename;
    btnOneDrive.disabled = false;
    closeModal();
  } catch (err) {
    showModal(errorHTML('Failed to convert: ' + err.message));
  }
}

// ── OneDrive upload ─────────────────────────────────────────────────────────

btnOneDrive.addEventListener('click', startOneDriveFlow);

async function startOneDriveFlow() {
  // 1. Check if already authenticated
  showModal(spinnerHTML('Checking authentication…'));
  const checkRes  = await fetch('/onedrive/check-auth').then(r => r.json());
  if (checkRes.authenticated) {
    doUpload();
  } else if (checkRes.no_client_id) {
    showModal(setupRequiredHTML());
  } else {
    startDeviceCodeFlow();
  }
}

async function startDeviceCodeFlow() {
  showModal(spinnerHTML('Starting authentication…'));
  const res  = await fetch('/onedrive/start-auth', { method: 'POST' }).then(r => r.json());
  if (res.error) { showModal(errorHTML(res.error)); return; }

  // Show device code to user
  showModal(deviceCodeHTML(res.user_code, res.verification_uri, res.expires_in));
  // Copy-friendly: open the page automatically
  window.open(res.verification_uri, '_blank');

  // Poll for completion
  pollAuth();
}

let _pollTimer = null;
async function pollAuth() {
  const res = await fetch('/onedrive/poll-auth').then(r => r.json());
  if (res.authenticated) {
    closeModal();
    doUpload();
  } else if (res.error) {
    showModal(errorHTML(res.error));
  } else {
    _pollTimer = setTimeout(pollAuth, 3000);
  }
}

async function doUpload() {
  showModal(spinnerHTML('Uploading to OneDrive…'));
  const res = await fetch('/onedrive/upload', { method: 'POST' }).then(r => r.json());
  if (res.error) { showModal(errorHTML(res.error)); return; }
  showModal(successHTML(res.web_url));
  window.open(res.web_url, '_blank');
}

// ── modal helpers ───────────────────────────────────────────────────────────

function showModal(html) {
  modalBody.innerHTML = html;
  modal.classList.add('open');
}
function closeModal() {
  clearTimeout(_pollTimer);
  modal.classList.remove('open');
  fetch('/onedrive/cancel-flow', { method: 'POST' }).catch(() => {});
}

function spinnerHTML(msg) {
  return `<div class="spinner"></div><p>${msg}</p>`;
}
function errorHTML(msg) {
  return `<h3>Error</h3><p class="error-msg">${msg}</p>`;
}
function deviceCodeHTML(code, url, expires) {
  return `
    <h3>Sign in to Microsoft</h3>
    <p>A browser window has opened. Enter this code at <strong>${url}</strong>:</p>
    <div class="device-code-box">
      <div class="code">${code}</div>
      <div class="url">${url}</div>
    </div>
    <div class="spinner"></div>
    <p class="status-msg">Waiting for you to sign in… (expires in ${Math.floor(expires/60)} min)</p>`;
}
function successHTML(url) {
  return `
    <h3>✅ Uploaded!</h3>
    <p>Your file has been uploaded to OneDrive and Word Online is opening.</p>
    <a href="${url}" target="_blank" class="btn btn-onedrive" style="display:inline-flex;margin-top:8px">
      Open in Word Online
    </a>`;
}
function setupRequiredHTML() {
  return `
    <h3>OneDrive not configured</h3>
    <p>
      You need to register an Azure app and run<br>
      <code style="background:#f3f4f6;padding:2px 6px;border-radius:4px">
        python word_opener.py --setup
      </code><br>
      to add your client ID. See the README for instructions.
    </p>`;
}
</script>
</body>
</html>
"""

# ── Flask routes ─────────────────────────────────────────────────────────────

@app.route("/")
def index():
    html  = _state.get("html_content")
    fname = Path(_state["docx_path"]).name if _state.get("docx_path") else None
    return render_template_string(TEMPLATE, html_content=html, filename=fname)


@app.route("/preview", methods=["POST"])
def preview():
    f = request.files.get("file")
    if not f or not f.filename.endswith(".docx"):
        return jsonify(error="Please upload a .docx file.")
    try:
        # save to a temp file so we can re-use it for upload
        tmp = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
        f.save(tmp.name)
        _state["docx_path"] = tmp.name
        _state["_orig_filename"] = f.filename

        with open(tmp.name, "rb") as fh:
            result = mammoth.convert_to_html(fh)
        _state["html_content"] = result.value
        return jsonify(html=result.value, filename=f.filename)
    except Exception as e:
        return jsonify(error=str(e))


@app.route("/onedrive/check-auth")
def check_auth():
    cfg = load_config()
    if not cfg.get("client_id", "").strip():
        return jsonify(authenticated=False, no_client_id=True)
    token = get_token_silent()
    return jsonify(authenticated=bool(token))


@app.route("/onedrive/start-auth", methods=["POST"])
def start_auth():
    msal_app = get_msal_app()
    if not msal_app:
        return jsonify(error="OneDrive not configured. Run: python word_opener.py --setup")
    flow = msal_app.initiate_device_flow(scopes=SCOPES)
    if "error" in flow:
        return jsonify(error=flow.get("error_description", flow["error"]))
    _state["device_flow"] = flow
    return jsonify(
        user_code=flow["user_code"],
        verification_uri=flow["verification_uri"],
        expires_in=flow.get("expires_in", 900),
    )


@app.route("/onedrive/poll-auth")
def poll_auth():
    flow = _state.get("device_flow")
    if not flow:
        return jsonify(error="No auth flow in progress.")
    msal_app = get_msal_app()
    # Non-blocking check
    result = msal_app.acquire_token_by_device_flow(flow, exit_condition=lambda f: True)
    if result and "access_token" in result:
        _state["device_flow"] = None
        _persist_cache()
        return jsonify(authenticated=True)
    err = result.get("error", "") if result else ""
    if err == "authorization_pending":
        return jsonify(authenticated=False, pending=True)
    return jsonify(error=result.get("error_description", err) if result else "Unknown error")


@app.route("/onedrive/cancel-flow", methods=["POST"])
def cancel_flow():
    _state["device_flow"] = None
    return jsonify(ok=True)


@app.route("/onedrive/upload", methods=["POST"])
def upload_to_onedrive():
    token = get_token_silent()
    if not token:
        return jsonify(error="Not authenticated. Please sign in first.")

    docx_path = _state.get("docx_path")
    if not docx_path or not os.path.exists(docx_path):
        return jsonify(error="No document loaded.")

    filename = _state.get("_orig_filename") or Path(docx_path).name

    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/octet-stream"}
    upload_url = f"{GRAPH_BASE}/me/drive/root:/Word Opener/{filename}:/content"

    try:
        with open(docx_path, "rb") as fh:
            resp = requests.put(upload_url, headers=headers, data=fh)
        if resp.status_code not in (200, 201):
            return jsonify(error=f"Upload failed ({resp.status_code}): {resp.text[:200]}")
        data = resp.json()
        web_url = data.get("webUrl", "")
        # Append ?action=edit to open in Word Online editor directly
        if web_url and "?action=edit" not in web_url:
            web_url += "?action=edit"
        return jsonify(web_url=web_url, name=data.get("name"))
    except Exception as e:
        return jsonify(error=str(e))


# ── CLI setup ─────────────────────────────────────────────────────────────────

def run_setup():
    print("\n=== Word Opener – OneDrive Setup ===\n")
    print("You need a Microsoft Azure app to upload to OneDrive.")
    print("Steps:")
    print("  1. Go to https://portal.azure.com > App registrations > New registration")
    print("  2. Name: 'Word Opener'  |  Account type: Personal Microsoft accounts only")
    print("  3. Redirect URI: Public client/native  ->  http://localhost")
    print("  4. API Permissions: Add Microsoft Graph > Delegated > Files.ReadWrite")
    print("  5. Copy the Application (client) ID shown on the Overview page\n")
    client_id = input("Paste your Application (client) ID: ").strip()
    if not client_id:
        print("Aborted – no client ID entered.")
        sys.exit(1)
    cfg = load_config()
    cfg["client_id"] = client_id
    save_config(cfg)
    # Clear token cache so a fresh login is triggered
    if TOKEN_CACHE_FILE.exists():
        TOKEN_CACHE_FILE.unlink()
    print(f"\n✓ Saved to {CONFIG_FILE}")
    print("Run 'python word_opener.py' and click 'Open in OneDrive' to sign in.\n")


def open_browser_delayed(url, delay=1.0):
    def _open():
        time.sleep(delay)
        webbrowser.open(url)
    threading.Thread(target=_open, daemon=True).start()


def main():
    parser = argparse.ArgumentParser(description="Word Opener – DOCX previewer with OneDrive")
    parser.add_argument("file", nargs="?", help="DOCX file to open")
    parser.add_argument("--setup", action="store_true", help="Configure OneDrive client ID")
    parser.add_argument("--port", type=int, default=5000, help="Local server port (default 5000)")
    parser.add_argument("--no-browser", action="store_true", help="Don't auto-open browser")
    args = parser.parse_args()

    if args.setup:
        run_setup()
        return

    # Pre-load a file if provided
    if args.file:
        p = Path(args.file).resolve()
        if not p.exists():
            print(f"Error: file not found: {p}")
            sys.exit(1)
        if p.suffix.lower() != ".docx":
            print("Error: only .docx files are supported.")
            sys.exit(1)
        _state["docx_path"] = str(p)
        _state["_orig_filename"] = p.name
        with open(p, "rb") as fh:
            result = mammoth.convert_to_html(fh)
        _state["html_content"] = result.value

    url = f"http://localhost:{args.port}"
    print(f"\n Word Opener running at {url}")
    if _state.get("docx_path"):
        print(f"  File: {_state['_orig_filename']}")
    print("  Press Ctrl+C to stop.\n")

    if not args.no_browser:
        open_browser_delayed(url)

    app.run(host="127.0.0.1", port=args.port, debug=False)


if __name__ == "__main__":
    main()
