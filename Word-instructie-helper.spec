# PyInstaller spec for Word-instructie-helper.
#
# Build met:
#   pip install -r requirements.txt pyinstaller
#   pyinstaller Word-instructie-helper.spec
#
# Resultaat staat in dist/Word-instructie-helper(.exe). Op Windows krijgt
# de eindgebruiker een enkele executable die de app als native venster
# (via pywebview + de WebView2/Edge-runtime) toont. Naast de .exe ontstaan
# automatisch ``word_files/`` (voor de eigen .docx-bestanden) en
# ``previews/`` (cache).

from PyInstaller.utils.hooks import collect_submodules, collect_data_files


hidden_imports: list[str] = []
# Uvicorn laadt protocols/loops via lazy imports die PyInstaller niet ziet.
hidden_imports += collect_submodules("uvicorn")
# FastAPI/Starlette/Pydantic gebruiken her en der dynamische imports.
hidden_imports += collect_submodules("fastapi")
hidden_imports += collect_submodules("starlette")
hidden_imports += collect_submodules("pydantic")
hidden_imports += collect_submodules("pydantic_core")
hidden_imports += collect_submodules("anyio")
# PyMuPDF is meestal vanzelf compleet, maar collect ze voor de zekerheid.
hidden_imports += collect_submodules("pymupdf")
hidden_imports += collect_submodules("fitz")
# Pywebview kiest zijn platform-backend op runtime via importlib.
hidden_imports += collect_submodules("webview")
# Edge/WebView2 op Windows draait via pythonnet/clr_loader.
hidden_imports += [
    "clr",
    "clr_loader",
    "pythonnet",
]

datas: list[tuple[str, str]] = []
# Bundle de statische HTML/JS-assets onder ``static/`` in het bundle.
datas.append(("static", "static"))
# Pydantic v2 heeft soms extra data files nodig.
datas += collect_data_files("pydantic")
# Pywebview heeft een paar JS-bridge bestanden die meegebundeld moeten worden.
datas += collect_data_files("webview")


a = Analysis(
    ["main.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="Word-instructie-helper",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    # console=False -> echte Windows-app: geen zwart consolevenster naast
    # het webview-venster. Logs gaan dan naar logbestand of stderr-buffer.
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
