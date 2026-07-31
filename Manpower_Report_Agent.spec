# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

from PyInstaller.utils.hooks import collect_all

project_root = Path(SPECPATH).resolve()

datas = [
    (
        str(project_root / "streamlit_app.py"),
        ".",
    )
]
binaries = []
hiddenimports = []

packages_to_collect = (
    "streamlit",
    "webview",
    "pdfplumber",
    "pdfminer",
    "openpyxl",
)

for package_name in packages_to_collect:
    package_data, package_binaries, package_hidden = (
        collect_all(package_name)
    )
    datas += package_data
    binaries += package_binaries
    hiddenimports += package_hidden

icon_path = (
    project_root
    / "assets"
    / "manpower_agent.ico"
)
icon_value = (
    str(icon_path)
    if icon_path.exists()
    else None
)

analysis = Analysis(
    ["desktop_launcher.py"],
    pathex=[str(project_root)],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "matplotlib.tests",
        "numpy.tests",
        "pandas.tests",
    ],
    noarchive=False,
)

python_archive = PYZ(
    analysis.pure,
)

executable = EXE(
    python_archive,
    analysis.scripts,
    [],
    exclude_binaries=True,
    name="Manpower Report Agent",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    icon=icon_value,
)

application = COLLECT(
    executable,
    analysis.binaries,
    analysis.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="Manpower Report Agent",
)
