# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec file cho Portal BAU (app.py)
# Build: pyinstaller app.spec

import os
import sys

block_cipher = None

# Đường dẫn gốc project
BASE_PATH = os.path.dirname(os.path.abspath(SPEC))

a = Analysis(
    ['app.py'],
    pathex=[BASE_PATH],
    binaries=[],
    datas=[
        # Thư mục templates (HTML)
        ('templates', 'templates'),
        # Thư mục static (CSS, JS, uploads)
        ('static', 'static'),
        # File config và database module
        ('config.py', '.'),
        ('database.py', '.'),
        # Thư mục history (chứa DB) - tạo rỗng nếu chưa có
        ('history', 'history'),
    ],
    hiddenimports=[
        'flask',
        'pandas',
        'openpyxl',
        'pyzipper',
        'sqlite3',
        'jinja2',
        'markupsafe',
        'werkzeug',
        'click',
        'itsdangerous',
        'blinker',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'gunicorn',  # Không cần trên Windows
        'tkinter',
        'matplotlib',
        'numpy.testing',
        'pytest',
    ],
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
    name='PortalBAU',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,  # Hiện console để xem log
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
    name='PortalBAU',
)
