#!/usr/bin/env python3
"""Generate PyInstaller .spec file with Tree() injection for lxml, openpyxl, et_xmlfile."""
import importlib
import sys
from pathlib import Path


def _pkg_path(name: str) -> str:
    """Resolve a pip-installed package's directory via importlib (safe at runtime)."""
    return str(Path(importlib.import_module(name).__file__).resolve().parent)


def gen_spec(source_py: str, app_name: str, spec_path: str):
    lx = _pkg_path("lxml")
    op = _pkg_path("openpyxl")
    et = _pkg_path("et_xmlfile")
    parts = [
        "a = Analysis(",
        f"    [r'{source_py}'],",
        "    pathex=[],",
        "    binaries=[], datas=[],",
        "    hiddenimports=['lxml','lxml.etree','lxml._elementpath','openpyxl','et_xmlfile'],",
        "    hookspath=[], hooksconfig={}, runtime_hooks=[],",
        "    excludes=['IPython','matplotlib','numpy','pandas','PIL','Pillow','scipy','sklearn','tensorflow','torch'],",
        "    win_no_prefer_redirects=False, win_private_assemblies=False,",
        "    cipher=None, noarchive=False,",
        ")",
        "pyz = PYZ(a.pure, a.zipped_data, cipher=None)",
        "exe = EXE(",
        "    pyz, a.scripts,",
        f"    Tree(r'{lx}', prefix='lxml'),",
        f"    Tree(r'{op}', prefix='openpyxl'),",
        f"    Tree(r'{et}', prefix='et_xmlfile'),",
        "    a.binaries, a.zipfiles, a.datas,",
        f"    name='{app_name}',",
        "    debug=False, bootloader_ignore_signals=False,",
        "    strip=False, upx=True, console=True,",
        "    disable_windowed_traceback=False, argv_emulation=False,",
        "    target_arch=None, codesign_identity=None, entitlements_file=None,",
        ")",
    ]
    with open(spec_path, 'w', encoding='utf-8') as f:
        f.write('# -*- mode: python ; coding: utf-8 -*-\\n')
        f.write('from pathlib import Path\\n')
        f.write('\\n'.join(parts))
        f.write('\\n')
    print(f'Spec generated: {spec_path}')

if __name__ == '__main__':
    gen_spec(sys.argv[1], sys.argv[2], sys.argv[3])
