# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['odoo.py'],
    pathex=['libs/autocad-mcp/src'],
    binaries=[],
    # 不打包 config/ 與 db/：兩者都含 Odoo API token，
    # 且執行期實際讀的是 c:/odoo/config/*.yaml 與 %APPDATA%/OdooAutoCAD/database.db，
    # 打包進去的副本從未被讀取，只會讓憑證隨安裝包散佈。
    # 注意：build_and_package.* 走 PyInstaller CLI，會重新產生本檔並洗掉這段註解；
    #       修改打包內容時請同步 .spec 與兩支建置腳本的 --add-data。
    datas=[('fonts', 'fonts'), ('icon', 'icon'), ('libs/autocad-mcp/lisp-code', 'lisp-code')],
    hiddenimports=['customtkinter', 'win32com.client', 'win32com.gen_py', 'pywintypes', 'win32api', 'tkinter', 'tkinter.ttk', 'sqlalchemy', 'sqlalchemy.ext.declarative', 'sqlalchemy.orm', 'mcp', 'mcp.server.fastmcp', 'mcp.types', 'uvicorn', 'starlette', 'structlog', 'autocad_mcp', 'autocad_mcp.server', 'autocad_mcp.client', 'autocad_mcp.config', 'autocad_mcp.backends', 'autocad_mcp.backends.file_ipc', 'autocad_mcp.backends.ezdxf_backend', 'httpx', 'anyio', 'sniffio', 'httpx_sse', 'pydantic', 'pydantic_settings', 'sse_starlette'],
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
