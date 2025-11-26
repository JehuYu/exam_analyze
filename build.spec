# -*- mode: python ; coding: utf-8 -*-
"""
优化的 PyInstaller 打包配置
目标：减小文件大小，加快启动速度，确保跨平台兼容性
"""

import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

# 排除不需要的模块以减小文件大小
excludes = [
    # 测试模块
    'pytest', 'unittest', 'test', 'tests',
    # 开发工具
    'IPython', 'jupyter', 'notebook', 'sphinx',
    # 不使用的科学计算模块
    'scipy', 'sklearn', 'tensorflow', 'torch', 'keras',
    # 不使用的数据库
    'sqlalchemy', 'sqlite3', 'psycopg2', 'pymysql',
    # 网络相关
    'flask', 'django', 'tornado', 'aiohttp', 'requests',
    # 其他不需要的
    'PIL', 'cv2', 'tkinter.test', 'email', 'html', 'http',
    'xml', 'xmlrpc', 'curses', 'multiprocessing', 'concurrent',
    # matplotlib 后端 (只保留 Agg)
    'matplotlib.backends.backend_qt5agg',
    'matplotlib.backends.backend_qt4agg', 
    'matplotlib.backends.backend_gtk3agg',
    'matplotlib.backends.backend_gtk3cairo',
    'matplotlib.backends.backend_tkagg',
    'matplotlib.backends.backend_wxagg',
    # numpy 不需要的部分
    'numpy.random._examples',
]

# 隐式导入的模块
hiddenimports = [
    'pandas',
    'numpy',
    'openpyxl',
    'docx',
    'matplotlib',
    'matplotlib.backends.backend_agg',
    'customtkinter',
]

a = Analysis(
    ['成绩分析GUI.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# 移除不需要的数据文件以减小体积
a.datas = [d for d in a.datas if not any(x in d[0] for x in [
    'tcl', 'tk', 'matplotlib/mpl-data/sample_data',
    'matplotlib/mpl-data/images', 'numpy/core/tests',
    'pandas/tests', 'docx/templates',
])]

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='成绩分析系统',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,  # 使用 UPX 压缩
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # 不显示控制台窗口
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # 可以添加图标: icon='icon.ico'
)

