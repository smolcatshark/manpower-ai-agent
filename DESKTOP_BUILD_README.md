# Manpower Report Agent — Desktop Build Kit

This folder turns the stable Streamlit application into a Windows
desktop application.

## Included files

- `streamlit_app.py` — stable application v0.8.4
- `desktop_launcher.py` — native-window launcher
- `Manpower_Report_Agent.spec` — PyInstaller build definition
- `requirements-desktop.txt` — desktop build dependencies
- `build_windows.bat` — local Windows build command
- `.github/workflows/build-windows.yml` — automatic GitHub Windows build

## Recommended build route: GitHub Actions

1. Copy all files into the root of the GitHub repository.
2. Commit and push them.
3. Open the repository on GitHub.
4. Open **Actions**.
5. Select **Build Windows Desktop App**.
6. Select **Run workflow**.
7. When the build finishes, download:
   `Manpower-Report-Agent-Windows`
8. Unzip the downloaded artifact.
9. Double-click:
   `Manpower Report Agent.exe`

The application starts Streamlit only on `127.0.0.1`, opens it inside
a pywebview desktop window, permits Excel downloads, and closes the
local server when the desktop window closes.

## Local Windows build

On a Windows computer with Python 3.12 installed, double-click:

`build_windows.bat`

The output is created under:

`dist\Manpower Report Agent\`

## Custom icon

A custom icon is optional. Create this file before building:

`assets\manpower_agent.ico`

The build works without an icon and uses the default executable icon.

## Important packaging choice

This first build uses PyInstaller **onedir** mode. It is more reliable
and starts faster for a Streamlit application with pandas, pdfplumber,
and openpyxl. The folder can later be wrapped in a normal Windows
installer with a desktop shortcut, so end users only interact with one
shortcut and the app window.
