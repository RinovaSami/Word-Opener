# Word Opener

A lightweight DOCX previewer with one-click OneDrive upload and Word Online editing.

## Features

- **Preview** any `.docx` file in your browser — full formatting, tables, images
- **Drag & drop** or use the file picker to load documents
- **Open in OneDrive** — uploads the file to your personal OneDrive and opens it in Word Online for editing
- Token caching so you only sign in once

## Quick start

```bash
# Install dependencies
pip install -r requirements.txt

# Open a specific file
python word_opener.py document.docx

# Or launch and use the file picker / drag-and-drop
python word_opener.py
```

The app opens at `http://localhost:5000` automatically.

## OneDrive setup (one-time)

To use the **Open in OneDrive** button you need a free Azure app registration:

1. Go to [https://portal.azure.com](https://portal.azure.com) and sign in with your Microsoft account.
2. Search for **App registrations** → **New registration**.
3. Fill in:
   - **Name**: `Word Opener` (or anything)
   - **Supported account types**: *Personal Microsoft accounts only*
   - **Redirect URI**: Platform = `Public client/native (mobile & desktop)`, URI = `http://localhost`
4. Click **Register**.
5. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated** → select `Files.ReadWrite` → **Add**.
6. Copy the **Application (client) ID** from the Overview page.
7. Run the setup command:

```bash
python word_opener.py --setup
```

Paste your client ID when prompted. Done — your credentials are saved to `~/.word-opener/config.json`.

The first time you click **Open in OneDrive** a Microsoft sign-in page opens automatically (device code flow). After signing in, all future uploads happen silently.

## Uploaded file location

Files are stored in `Word Opener/` folder in the root of your OneDrive, preserving the original filename. The app opens them directly in Word Online (`?action=edit`).

## Options

```
python word_opener.py [file.docx] [options]

Positional:
  file          DOCX file to open at startup

Options:
  --setup       Configure OneDrive client ID
  --port N      Local port (default: 5000)
  --no-browser  Don't auto-open the browser
```

## Dependencies

| Package   | Purpose                          |
|-----------|----------------------------------|
| flask     | Local web server                 |
| mammoth   | DOCX → HTML conversion           |
| msal      | Microsoft identity / OAuth 2.0   |
| requests  | Microsoft Graph API calls        |
