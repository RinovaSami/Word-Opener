# Word Opener

A lightweight, self-contained DOCX previewer. No accounts, no API keys, no cloud setup required.

## Features

- **Preview** any `.docx` file in your browser — full formatting, tables, images
- **Drag & drop** or use the file picker to load documents
- **Download** the current file directly from the toolbar
- **Open OneDrive** button opens [onedrive.live.com](https://onedrive.live.com) in your browser so you can upload the file manually (assumes you're already signed in)

## Quick start

```bash
# Install dependencies (only two packages needed)
pip install -r requirements.txt

# Open a specific file
python word_opener.py document.docx

# Or launch and use the file picker / drag-and-drop
python word_opener.py
```

The app opens at `http://localhost:5000` automatically.

## Options

```
python word_opener.py [file.docx] [options]

Positional:
  file          DOCX file to open at startup

Options:
  --port N      Local port (default: 5000)
  --no-browser  Don't auto-open the browser
```

## Windows EXE (no Python required)

Download `WordOpener.exe` from the [Releases](../../releases) page (or the
latest [GitHub Actions artifact](../../actions/workflows/build-windows-exe.yml))
and double-click it — no installation needed.

```
WordOpener.exe                     # launch with drag-and-drop file picker
WordOpener.exe document.docx       # open a specific file
WordOpener.exe --port 8080         # use a custom port
WordOpener.exe --no-browser        # don't auto-open browser
```

### Build the EXE yourself (Windows)

```bat
git clone <this-repo>
cd Word-Opener
build.bat
:: output: dist\WordOpener.exe
```

`build.bat` installs `pyinstaller`, `flask`, and `mammoth` automatically, then
calls `pyinstaller word_opener.spec --clean` to produce a single, standalone
`dist\WordOpener.exe` that embeds the Python runtime and all dependencies.

## Dependencies

| Package  | Purpose               |
|----------|-----------------------|
| flask    | Local web server      |
| mammoth  | DOCX → HTML conversion |

Both are pure-Python and install via `pip` with no system dependencies.
