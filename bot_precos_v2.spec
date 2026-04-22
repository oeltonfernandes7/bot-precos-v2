# bot_precos_v2.spec — Configuracao do PyInstaller para o BotPrecos
# Use: python -m PyInstaller bot_precos_v2.spec

from PyInstaller.utils.hooks import collect_all, collect_data_files

block_cipher = None

# Coleta recursiva de todos os arquivos de cada pacote
streamlit_datas,  streamlit_bins,  streamlit_hidden  = collect_all("streamlit")
plotly_datas,     plotly_bins,     plotly_hidden      = collect_all("plotly")
altair_datas,     altair_bins,     altair_hidden      = collect_all("altair")

# Arquivos do proprio projeto que precisam estar dentro do bundle
project_datas = [
    ("app.py",               "."),
    ("bot_precos.py",        "."),
    ("config.py",            "."),
    ("fontes.json",          "."),
    ("investigar_sites.py",  "."),
    (".streamlit/config.toml", ".streamlit"),
]

all_datas    = streamlit_datas   + plotly_datas   + altair_datas   + project_datas
all_binaries = streamlit_bins    + plotly_bins     + altair_bins
all_hidden   = streamlit_hidden  + plotly_hidden   + altair_hidden  + [
    "openpyxl",
    "openpyxl.cell._translator",
    "pandas",
    "pandas._libs.tslibs.timedeltas",
    "requests",
    "playwright",
    "psutil",
    "packaging",
    "pkg_resources",
    "streamlit.web.cli",
    "streamlit.runtime.scriptrunner.magic_funcs",
]

a = Analysis(
    ["launcher.py"],
    pathex=["."],
    binaries=all_binaries,
    datas=all_datas,
    hiddenimports=all_hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["tkinter", "matplotlib", "scipy", "PIL"],
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
    name="BotPrecos",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,        # Sem janela de console visivel
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="BotPrecos",
)
git init
git add .
git commit -m "Bot de precos inicial"
git branch -M main
git remote add origin https://github.com/oeltonfernandes7/bot-precos-v2.git
git push -u origin main