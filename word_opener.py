#!/usr/bin/env python3
"""
Word Opener - Lightweight DOCX previewer

Usage:
    python word_opener.py [file.docx]
    python word_opener.py --port 8080
"""

import os
import sys
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

mammoth = _require("mammoth")
_require("flask")

from flask import Flask, render_template_string, request, jsonify, send_file

# ── app state ────────────────────────────────────────────────────────────────

_state = {
    "docx_path":     None,
    "html_content":  None,
    "orig_filename": None,
}

app = Flask(__name__)

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
      text-decoration: none;
    }
    .btn:disabled { opacity: .5; cursor: not-allowed; }
    .btn-primary   { background: #2563eb; color: #fff; }
    .btn-primary:hover:not(:disabled) { background: #1d4ed8; }
    .btn-onedrive  { background: #0078d4; color: #fff; }
    .btn-onedrive:hover:not(:disabled) { background: #005a9e; }
    .btn-ghost     { background: rgba(255,255,255,.15); color: #fff; border: 1px solid rgba(255,255,255,.3); }
    .btn-ghost:hover:not(:disabled) { background: rgba(255,255,255,.25); }
    .btn-download  { background: #059669; color: #fff; }
    .btn-download:hover:not(:disabled) { background: #047857; }

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

    /* ── toast ── */
    #toast {
      position: fixed; bottom: 24px; left: 50%; transform: translateX(-50%);
      background: #1e3a5f; color: #fff;
      padding: 10px 20px; border-radius: 8px;
      font-size: .85rem; opacity: 0;
      transition: opacity .3s;
      pointer-events: none; z-index: 300;
    }
    #toast.show { opacity: 1; }

    @media (max-width: 640px) {
      #doc-content { padding: 32px 24px; }
      #toolbar .actions .btn span { display: none; }
    }
  </style>
</head>
<body>

<!-- toolbar -->
<div id="toolbar">
  <div class="logo">&#128196; Word Opener</div>
  <div class="filename" id="toolbar-filename">{{ filename or 'No file loaded' }}</div>
  <div class="actions">
    <button class="btn btn-ghost" id="btn-open">Open file</button>
    <button class="btn btn-download" id="btn-download" {% if not filename %}disabled{% endif %}>
      &#8595; Download
    </button>
    <a class="btn btn-onedrive" id="btn-onedrive"
       href="https://onedrive.live.com" target="_blank" rel="noopener"
       {% if not filename %}style="pointer-events:none;opacity:.5"{% endif %}>
      &#9729; OneDrive
    </a>
  </div>
</div>
<input type="file" id="file-input" accept=".docx">

<!-- drop zone (shown when no file loaded) -->
<div id="drop-zone">
  <div class="icon">&#128194;</div>
  <h2>Open a DOCX file</h2>
  <p>Drag &amp; drop a file here, or click <strong>Open file</strong> above.</p>
  <button class="btn btn-primary" onclick="document.getElementById('file-input').click()">
    Choose file&hellip;
  </button>
</div>

<!-- document view -->
<div id="doc-wrap" {% if html_content %}class="visible"{% endif %}>
  <div id="doc-content">{{ html_content|safe if html_content else '' }}</div>
</div>

<div id="toast"></div>

<script>
const dropZone        = document.getElementById('drop-zone');
const docWrap         = document.getElementById('doc-wrap');
const docContent      = document.getElementById('doc-content');
const fileInput       = document.getElementById('file-input');
const btnDownload     = document.getElementById('btn-download');
const btnOneDrive     = document.getElementById('btn-onedrive');
const toolbarFilename = document.getElementById('toolbar-filename');
const toast           = document.getElementById('toast');

{% if filename %}
dropZone.classList.add('has-file');
{% endif %}

// ── file loading ─────────────────────────────────────────────────────────────

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
  else showToast('Please drop a .docx file.');
});

async function uploadForPreview(file) {
  const fd = new FormData();
  fd.append('file', file);
  showToast('Converting\u2026');
  try {
    const res  = await fetch('/preview', { method: 'POST', body: fd });
    const data = await res.json();
    if (data.error) { showToast('Error: ' + data.error); return; }
    docContent.innerHTML = data.html;
    docWrap.classList.add('visible');
    dropZone.classList.add('has-file');
    toolbarFilename.textContent = data.filename;
    btnDownload.disabled = false;
    btnOneDrive.style.pointerEvents = '';
    btnOneDrive.style.opacity = '';
    showToast('Document loaded.');
  } catch (err) {
    showToast('Failed to convert: ' + err.message);
  }
}

// ── download current DOCX ────────────────────────────────────────────────────

btnDownload.addEventListener('click', () => {
  window.location.href = '/download';
});

// ── toast helper ─────────────────────────────────────────────────────────────

let _toastTimer;
function showToast(msg) {
  toast.textContent = msg;
  toast.classList.add('show');
  clearTimeout(_toastTimer);
  _toastTimer = setTimeout(() => toast.classList.remove('show'), 3000);
}
</script>
</body>
</html>
"""

# ── Flask routes ──────────────────────────────────────────────────────────────

@app.route("/")
def index():
    html  = _state.get("html_content")
    fname = _state.get("orig_filename")
    return render_template_string(TEMPLATE, html_content=html, filename=fname)


@app.route("/preview", methods=["POST"])
def preview():
    f = request.files.get("file")
    if not f or not f.filename.endswith(".docx"):
        return jsonify(error="Please upload a .docx file.")
    try:
        tmp = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
        f.save(tmp.name)
        _state["docx_path"]    = tmp.name
        _state["orig_filename"] = f.filename

        with open(tmp.name, "rb") as fh:
            result = mammoth.convert_to_html(fh)
        _state["html_content"] = result.value
        return jsonify(html=result.value, filename=f.filename)
    except Exception as e:
        return jsonify(error=str(e))


@app.route("/download")
def download():
    path = _state.get("docx_path")
    if not path or not os.path.exists(path):
        return "No document loaded.", 404
    name = _state.get("orig_filename") or Path(path).name
    return send_file(path, as_attachment=True, download_name=name,
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


# ── entrypoint ────────────────────────────────────────────────────────────────

def open_browser_delayed(url, delay=1.0):
    def _open():
        time.sleep(delay)
        webbrowser.open(url)
    threading.Thread(target=_open, daemon=True).start()


def main():
    parser = argparse.ArgumentParser(description="Word Opener – DOCX previewer")
    parser.add_argument("file",        nargs="?", help="DOCX file to open")
    parser.add_argument("--port",      type=int, default=5000, help="Local server port (default 5000)")
    parser.add_argument("--no-browser", action="store_true",   help="Don't auto-open browser")
    args = parser.parse_args()

    if args.file:
        p = Path(args.file).resolve()
        if not p.exists():
            print(f"Error: file not found: {p}")
            sys.exit(1)
        if p.suffix.lower() != ".docx":
            print("Error: only .docx files are supported.")
            sys.exit(1)
        _state["docx_path"]    = str(p)
        _state["orig_filename"] = p.name
        with open(p, "rb") as fh:
            result = mammoth.convert_to_html(fh)
        _state["html_content"] = result.value

    url = f"http://localhost:{args.port}"
    print(f"\n  Word Opener running at {url}")
    if _state.get("orig_filename"):
        print(f"  File: {_state['orig_filename']}")
    print("  Press Ctrl+C to stop.\n")

    if not args.no_browser:
        open_browser_delayed(url)

    app.run(host="127.0.0.1", port=args.port, debug=False)


if __name__ == "__main__":
    main()
