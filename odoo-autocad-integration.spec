# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['odoo.py'],
    pathex=[],
    binaries=[],
    datas=[('config', 'config'), ('fonts', 'fonts'), ('icon', 'icon'), ('db', 'db'), ('libs/autocad-mcp/src', 'autocad_mcp_src'), ('libs/autocad-mcp/lisp-code', 'lisp-code')],
    hiddenimports=['customtkinter', 'win32com.client', 'win32com.gen_py', 'pywintypes', 'win32api', 'tkinter', 'tkinter.ttk', 'sqlalchemy', 'sqlalchemy.ext.declarative', 'sqlalchemy.orm', 'mcp', 'mcp.server.fastmcp', 'mcp.types', 'uvicorn', 'starlette', 'structlog', 'httpx', 'anyio', 'sniffio', 'httpx_sse', 'pydantic', 'pydantic_settings', 'sse_starlette'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['pytest', 'unittest', 'doctest', 'pdb', 'matplotlib', 'numpy', 'pandas'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='odoo-autocad-integration',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['icon\\odoo_autocad.ico'],
)
