# -*- mode: python ; coding: utf-8 -*-

import os


proposal_summary_template = '/Users/vernabbott/Library/CloudStorage/OneDrive-Personal/1. Proposal Summary Template.emltpl'
datas = [
    ('templates', 'templates'),
    ('static', 'static'),
    ('roof_intelligence_single_address.py', '.'),
    ('roof_intelligence_area_batch.py', '.'),
    ('roof_report_naming.py', '.'),
]
try:
    with open(proposal_summary_template, 'rb'):
        pass
except OSError:
    pass
else:
    datas.append((proposal_summary_template, 'resources'))

a = Analysis(
    ['run_app.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=['xlwings', 'docx', 'docx2pdf'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['pandas', 'numpy'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='PCS_Proposal',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    # This Mac currently distributes PCS locally rather than through the App
    # Store. Let PyInstaller apply one consistent ad-hoc signature to the
    # executable and bundled Python framework.
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='PCS_Proposal',
)
app = BUNDLE(
    coll,
    name='PCS_Proposal.app',
    icon='build_assets/PCS_Proposal.icns',
    bundle_identifier='com.procoatingsystems.pcsproposal',
)
