# -*- mode: python ; coding: utf-8 -*-
#
# 本檔是打包設定的唯一來源。build_and_package.bat / .ps1 都以
#   pyinstaller --clean --noconfirm --distpath output odoo-autocad-integration.spec
# 呼叫，不再重複一長串 CLI 參數（過去 CLI 建置會重新產生本檔並洗掉註解，
# 且需人工同步三處設定）。
#
# 因此新增／移除 hidden import、資料檔、排除模組時，只要改這裡。


a = Analysis(
    ['odoo.py'],
    pathex=['libs/autocad-mcp/src'],
    binaries=[],
    # 不打包 config/ 與 db/：兩者都含 Odoo API token，
    # 且執行期實際讀的是 c:/odoo/config/*.yaml 與 %APPDATA%/OdooAutoCAD/database.db，
    # 打包進去的副本從未被讀取，只會讓憑證隨安裝包散佈。
    #
    # 不打包 fonts/：那三個 msjh*.ttc 是 C:\Windows\Fonts 系統字型的逐位元組複本
    # （合計 48.7 MB，佔舊版 EXE 的 41.8%）。ui/ui_fonts.py 是用 tkfont.families()
    # 依「名稱」解析系統已安裝字型，程式中沒有任何 AddFontResource 或讀取 fonts/
    # 的路徑，因此打包進去的字型從未被載入。
    datas=[('icon', 'icon'), ('libs/autocad-mcp/lisp-code', 'lisp-code')],
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
