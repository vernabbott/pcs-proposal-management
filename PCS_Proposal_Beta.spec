# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path


project_dir = Path(SPEC).resolve().parent
icon_path = project_dir / "build_assets" / "PCS_Proposal.icns"
proposal_summary_template = Path(
    "/Users/vernabbott/Library/CloudStorage/OneDrive-Personal/"
    "1. Proposal Summary Template.emltpl"
)

datas = [
    (str(project_dir / "templates"), "templates"),
    (str(project_dir / "static"), "static"),
    (str(project_dir / "roof_intelligence_single_address.py"), "."),
    (str(project_dir / "roof_intelligence_area_batch.py"), "."),
    (str(project_dir / "roof_report_naming.py"), "."),
]
if proposal_summary_template.is_file():
    datas.append((str(proposal_summary_template), "resources"))

a = Analysis(
    [str(project_dir / "run_beta_app.py")],
    pathex=[str(project_dir)],
    binaries=[],
    datas=datas,
    hiddenimports=["xlwings", "docx", "docx2pdf"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["pandas", "numpy"],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="PCS_Proposal_Beta",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=str(icon_path),
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="PCS_Proposal_Beta",
)
app = BUNDLE(
    coll,
    name="PCS_Proposal_Beta.app",
    icon=str(icon_path),
    bundle_identifier="com.procoatingsystems.pcsproposal.beta",
    info_plist={
        "CFBundleDisplayName": "PCS Proposal Beta",
        "CFBundleName": "PCS Proposal Beta",
        "NSHighResolutionCapable": True,
    },
)
