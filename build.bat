@echo off
setlocal

echo ============================================================
echo  Word Opener -- Windows EXE Build
echo ============================================================
echo.

:: ── install build deps ────────────────────────────────────────
echo [1/3] Installing dependencies...
pip install flask mammoth pyinstaller
if errorlevel 1 (
    echo ERROR: pip install failed.
    exit /b 1
)

echo.

:: ── run pyinstaller ───────────────────────────────────────────
echo [2/3] Building executable...
pyinstaller word_opener.spec --clean
if errorlevel 1 (
    echo ERROR: PyInstaller build failed.
    exit /b 1
)

echo.

:: ── done ──────────────────────────────────────────────────────
echo [3/3] Done!
echo.
echo   Output: dist\WordOpener.exe
echo.
echo Usage:
echo   dist\WordOpener.exe                     -- launch with file picker
echo   dist\WordOpener.exe document.docx       -- open a specific file
echo   dist\WordOpener.exe --port 8080         -- use a custom port
echo   dist\WordOpener.exe --no-browser        -- don't auto-open browser
echo.

endlocal
