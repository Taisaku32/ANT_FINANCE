# finanzas.spec  — PyInstaller build specification
# Run with:  pyinstaller finanzas.spec

import sys
from PyInstaller.utils.hooks import collect_all, collect_data_files

block_cipher = None

# ── Collect Streamlit (includes its web assets, templates, etc.) ──────────────
st_datas, st_binaries, st_hidden = collect_all("streamlit")

# ── Other packages that ship data files ──────────────────────────────────────
plotly_datas    = collect_data_files("plotly")
pandas_datas    = collect_data_files("pandas")
openpyxl_datas  = collect_data_files("openpyxl")

# ── Analysis ──────────────────────────────────────────────────────────────────
a = Analysis(
    ["main.py"],
    pathex=[],
    binaries=st_binaries,
    datas=(
        st_datas
        + plotly_datas
        + pandas_datas
        + openpyxl_datas
        + [
            # dashboard.py must live next to the .exe so --run-dashboard finds it
            ("dashboard.py", "."),
        ]
    ),
    hiddenimports=(
        st_hidden
        + [
            # Streamlit internals
            "streamlit.web.cli",
            "streamlit.web.bootstrap",
            "streamlit.runtime.scriptrunner.magic_funcs",
            # Plotly
            "plotly",
            "plotly.express",
            "plotly.graph_objects",
            # Data
            "pandas",
            "openpyxl",
            "openpyxl.styles",
            "openpyxl.utils",
            # Stdlib
            "sqlite3",
            "email.mime.text",
            "email.mime.multipart",
            "smtplib",
            # Tornado / asyncio (used by Streamlit)
            "tornado",
            "tornado.platform.asyncio",
        ]
    ),
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="FinanzasPersonales",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,       # No ventana de CMD
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon="icono.ico",  # Descomenta y agrega un .ico si quieres un ícono
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="FinanzasPersonales",
)
