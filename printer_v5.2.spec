# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all
import os

# 数据文件和二进制文件
datas = [
    # 包含SumatraPDF目录
    ('external/SumatraPDF', 'external/SumatraPDF'),
    # 包含资源文件
    ('resources', 'resources'),
    # 包含配置文件目录
    ('data', 'data'),
]
binaries = []

# 隐藏导入模块
hiddenimports = [
    # GUI相关
    'tkinterdnd2', 
    'tkinterdnd2.TkinterDnD',
    
    # Windows API
    'win32com.client',
    'win32print',
    'win32api',
    'win32con',
    'comtypes.client',
    'pythoncom',
    
    # 文档处理库
    'PyPDF2',
    'openpyxl',
    'xlwings',
    'docx',
    'pptx',
    
    # 图像处理
    'PIL',
    'PIL.Image',
    'PIL.ImageFile',
    
    # 标准库模块
    'pathlib',
    'subprocess',
    'threading',
    'queue',
    'json',
    'datetime',
    'time',
    'os',
    'sys',
    
    # 项目模块 - 处理器
    'src.handlers.pdf_handler',
    'src.handlers.word_handler',
    'src.handlers.powerpoint_handler',
    'src.handlers.excel_handler',
    'src.handlers.image_handler',
    'src.handlers.text_handler',
    'src.handlers.handler_registry',
    'src.handlers.base_handler',
    'src.handlers.print_utils',
    
    # 项目模块 - 核心
    'src.core.document_manager',
    'src.core.print_controller',
    'src.core.page_count_manager',
    'src.core.settings_manager',
    'src.core.printer_config_manager',
    'src.core.models',
    
    # 项目模块 - GUI
    'src.gui.main_window',
    'src.gui.print_settings_dialog',
    'src.gui.page_count_dialog',
    'src.gui.components.file_import_handler',
    'src.gui.components.list_operation_handler',
    'src.gui.components.window_manager',
    'src.gui.components.tooltip',
    
    # 项目模块 - 工具
    'src.utils.config_utils',
    'src.utils.path_utils',
]

# 收集tkinterdnd2相关文件
tmp_ret = collect_all('tkinterdnd2')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # 排除不需要的模块以减小文件大小
        'matplotlib',
        'numpy',
        'pandas',
        'scipy',
        'IPython',
        'jupyter',
        'pytest',
        'unittest',
    ],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='办公文档批量打印器v5.2',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[
        # 不压缩这些文件，避免运行时问题
        'SumatraPDF.exe',
        'vcruntime*.dll',
        'msvcp*.dll',
        'api-ms-win*.dll',
    ],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='resources\\app_icon.ico',
    version='version_info.txt',
)