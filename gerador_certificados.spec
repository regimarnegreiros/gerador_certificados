# -*- mode: python ; coding: utf-8 -*-
"""
Configuração PyInstaller para o Gerador de Certificados ANM (GUI PySide6).
Modo --onedir --windowed --noupx para abertura imediata e compatibilidade com EDR
(sem descompactação de binários em %TEMP% e sem compressão UPX).
"""

from pathlib import Path

diretorio_raiz = Path(SPECPATH).resolve()

datas_extras = []
pasta_assets = diretorio_raiz / "assets"
if pasta_assets.exists():
    datas_extras.append((str(pasta_assets), "assets"))

icone_path = pasta_assets / "anm.ico"
icone_arg = str(icone_path) if icone_path.exists() else None
version_arg = str(diretorio_raiz / "version_info.txt") if (diretorio_raiz / "version_info.txt").exists() else None

modulos_excluidos = [
    "tkinter",
    "_tkinter",
    "unittest",
    "pytest",
    "matplotlib",
    "scipy",
    "IPython",
    "PySide6.QtWebEngineCore",
    "PySide6.QtWebEngineWidgets",
    "PySide6.QtWebEngineQuick",
    "PySide6.QtQml",
    "PySide6.QtQuick",
    "PySide6.QtQuickWidgets",
    "PySide6.QtQuickControls2",
    "PySide6.QtQuick3D",
    "PySide6.Qt3DCore",
    "PySide6.Qt3DRender",
    "PySide6.Qt3DInput",
    "PySide6.Qt3DExtras",
    "PySide6.Qt3DAnimation",
    "PySide6.Qt3DLogic",
    "PySide6.QtMultimedia",
    "PySide6.QtMultimediaWidgets",
    "PySide6.QtNetwork",
    "PySide6.QtSql",
    "PySide6.QtTest",
    "PySide6.QtDesigner",
    "PySide6.QtHelp",
    "PySide6.QtUiTools",
    "PySide6.QtPrintSupport",
    "PySide6.QtXml",
    "PySide6.QtDBus",
    "PySide6.QtConcurrent",
    "PySide6.QtPdf",
    "PySide6.QtPdfWidgets",
    "PySide6.QtCharts",
    "PySide6.QtGraphs",
    "PySide6.QtOpenGL",
    "PySide6.QtOpenGLWidgets",
]

a = Analysis(
    [str(diretorio_raiz / "app.py")],
    pathex=[str(diretorio_raiz)],
    binaries=[],
    datas=datas_extras,
    hiddenimports=[
        "win32com",
        "win32com.client",
        "pythoncom",
        "pywintypes",
        "openpyxl.cell._writer",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=modulos_excluidos,
    noarchive=False,
    optimize=0,
)

# Remove traduções e plugins Qt pesados/não utilizados pelo app Widgets
_PREFIXOS_DATA_EXCLUIR = (
    "PySide6\\translations\\",
    "PySide6/translations/",
    "PySide6\\plugins\\qmltooling\\",
    "PySide6/plugins/qmltooling/",
    "PySide6\\plugins\\designer\\",
    "PySide6/plugins/designer/",
    "PySide6\\plugins\\sqldrivers\\",
    "PySide6/plugins/sqldrivers/",
    "PySide6\\plugins\\networkinformation\\",
    "PySide6/plugins/networkinformation/",
    "PySide6\\plugins\\tls\\",
    "PySide6/plugins/tls/",
    "PySide6\\plugins\\generic\\",
    "PySide6/plugins/generic/",
    "PySide6\\plugins\\vectorimageformats\\",
    "PySide6/plugins/vectorimageformats/",
)

_DLLS_QT_EXCLUIR = (
    "qt6qml",
    "qt6quick",
    "qt6designer",
    "qt6lottie",
    "qt6labs",
    "qt6webengine",
    "qt6sql",
    "qt6network",
    "qt6test",
    "qt6multimedia",
    "qt63d",
    "qt6pdf",
    "qt6xml",
    "qt6printsupport",
    "qt6dbus",
    "qt6concurrent",
    "qt6help",
    "qt6opengl",
    "pyside6qml",
)


def _manter_binario(item) -> bool:
    nome_norm = item[0].replace("/", "\\").lower()
    if any(p.replace("/", "\\").lower() in nome_norm for p in _PREFIXOS_DATA_EXCLUIR):
        return False
    base = Path(nome_norm).name
    if any(base.startswith(pref) for pref in _DLLS_QT_EXCLUIR):
        return False
    return True


def _manter_data(item) -> bool:
    nome_norm = item[0].replace("/", "\\").lower()
    return not any(p.replace("/", "\\").lower() in nome_norm for p in _PREFIXOS_DATA_EXCLUIR)


a.binaries = [b for b in a.binaries if _manter_binario(b)]
a.datas = [d for d in a.datas if _manter_data(d)]

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="GeradorCertificadosANM",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    version=version_arg,
    icon=icone_arg,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="GeradorCertificadosANM",
)
