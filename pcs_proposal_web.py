import os

from dotenv import load_dotenv
load_dotenv()

# Directory constants (driven by environment variables; safe defaults for local dev)
PROPOSALS_DIR  = os.environ.get("PROPOSALS_DIR", "./proposals")
CONTRACTS_DIR  = os.environ.get("CONTRACTS_DIR", "./contracts")
COMPLETED_DIR  = os.environ.get("COMPLETED_DIR", "./completed")
DEADFILE_DIR   = os.environ.get("DEADFILE_DIR", "./dead")
TEMPLATE_DIR   = os.environ.get("TEMPLATE_DIR", "./templates")
LIBREOFFICE_PATH = os.environ.get("LIBREOFFICE_PATH", "/usr/bin/soffice")

def create_proposal_from_fields(customer_name,
                                street_address,
                                city,
                                state,
                                zip_code,
                                roof_type,
                                total_squares,
                                warranty_incl,
                                product,
                                proposal_language,
                                submitted_by,
                                target_folder: str | None = None,
                                mapped_data: dict | None = None,
                                pdf_async: bool = True,
                                use_libreoffice: bool = True):
    today = datetime.date.today()
    formatted_date = today.strftime("%B %d, %Y")

    # Determine placeholder values based on warranty and product
    if product == "Gaco":
        if warranty_incl == "Yes":
            price_includes_text = "* Price Includes material, labor, trash pickup, haul away and Gaco Warranty Fee"
            warranty_text = "IS INCLUDED"
        else:
            price_includes_text = "* Price Includes material, labor, trash pickup and haul away"
            warranty_text = "IS NOT INCLUDED"
    else:
        price_includes_text = proposal_language if proposal_language else ' '
        warranty_text = warranty_incl

    # Use computed totals passed from the UI/calculation_routine (authoritative)
    # Fall back to 0 if not provided.
    tp10 = 0
    tp15 = 0
    tp20 = 0
    if mapped_data:
        try:
            tp10 = float(mapped_data.get("total_price_10", 0) or 0)
            tp15 = float(mapped_data.get("total_price_15", 0) or 0)
            tp20 = float(mapped_data.get("total_price_20", 0) or 0)
        except Exception:
            tp10, tp15, tp20 = 0, 0, 0

    # Prepare replacements for placeholders (using double-bracket format as in template)
    replacements = {
        '[[CustomerName]]': customer_name,
        '[[ProjectStreetAddr]]': street_address,
        '[[ProjectCity]]': city,
        '[[ProjectState]]': state,
        '[[ProjectZip]]': zip_code,
        '[[Date]]': formatted_date,
        '[[Squares]]': total_squares,
        '[[PriceIncludesLanguage]]': price_includes_text,
        '[[WarrantyIncluded]]': warranty_text,
        '[[SubmittedBy]]': submitted_by,
        '[[10YrTotalPrice]]': f"{tp10:,.0f}",
        '[[15YrTotalPrice]]': f"{tp15:,.0f}",
        '[[20YrTotalPrice]]': f"{tp20:,.0f}",
        '[[AdditionalLanguage]]': proposal_language if proposal_language else ' '
    }

    # Folder and file names
    # Folder and file names
    if target_folder:
        proposal_folder = target_folder
        os.makedirs(proposal_folder, exist_ok=True)
        folder_name = os.path.basename(proposal_folder)
    else:
        folder_name = f"{customer_name} - {street_address}"
        proposal_folder = os.path.join(PROPOSALS_DIR, folder_name)
        os.makedirs(proposal_folder, exist_ok=True)

    # Map roof type to suffix
    roof_suffix_map = {
        "TPO/EPDM": "TPO EPDM Metal.docx",
        "Metal": "TPO EPDM Metal.docx",
        "Mod Bit": "Mod Bit.docx",
        "Rock/Foam/Coat": "RFC.docx",
        "Ballasted 45 mil": "Ballasted 45mil.docx",
        "Ballasted 60 mil": "Ballasted 60mil.docx"
    }
    roof_suffix = roof_suffix_map.get(roof_type, "Unknown.docx")

    # Select template based on product and roof type
    prefix = "Gaco S42 Proposal - " if product == "Gaco" else "Uniflex Proposal - "
    doc_template_name = f"{prefix}{roof_suffix}"
    doc_output_name = f"{prefix}{street_address}.docx"
    doc_template_path = os.path.join(TEMPLATE_DIR, doc_template_name)
    doc_output_path = os.path.join(proposal_folder, doc_output_name)

    # Copy template and replace placeholders directly in the output file
    shutil.copy(doc_template_path, doc_output_path)
    doc = Document(doc_output_path)
    replace_placeholder_blocks(doc, replacements)
    doc.save(doc_output_path)

    # Convert Word doc to PDF and save in same folder (headless if possible)
    _convert_to_pdf(
        doc_output_path,
        proposal_folder,
        use_libreoffice=use_libreoffice,
        async_mode=pdf_async,
    )

    # Copy Excel files
    profit_template = os.path.join(TEMPLATE_DIR, "Profit Summary.xlsm")
    profit_output = os.path.join(proposal_folder, f"Profit Summary - {street_address}.xlsm")
    shutil.copy(profit_template, profit_output)

    # Update Profit Summary.xlsm using central map
    app_excel = xw.App(visible=False)
    app_excel.display_alerts = False
    app_excel.screen_updating = False
    try:
        wb_profit = app_excel.books.open(profit_output)
        # Prepare default header-only map, then merge any provided mapped_data
        default_header_map = {
            "customer_name": customer_name,
            "street_address": street_address,
            "city": city,
            "state": state,
            "zip_code": zip_code,
            "squares": total_squares,
            "current_roof": roof_type,
            "product": product,
            "warranty_incl": warranty_incl,
            "submitted_by": submitted_by,
            # Optional seed values; leave commented unless you want to pre-populate
            # "price_per_sq_10": None,
            # "labor_days": None,
            "proposal_note": "",
        }
        merged_map = dict(default_header_map)
        if mapped_data:
            merged_map.update({k: v for k, v in mapped_data.items() if k in EXCEL_CELL_MAP and EXCEL_CELL_MAP[k]})
        write_fields_to_profit_summary(wb_profit, merged_map)
        wb_profit.save()
        wb_profit.close()
    finally:
        app_excel.quit()

    return folder_name
from flask import Flask, render_template, request, send_file, redirect, url_for, flash
from docx2pdf import convert
from docx import Document
import pandas as pd
import os
import math
import shutil
import datetime
import glob
import xlwings as xw
from decimal import Decimal, ROUND_HALF_UP
import subprocess
import threading
import shlex
import sys

# Flask app, List and Detail forms were saved and are working correctly at 8/28 2:24PM

# Resolve template/static paths for both dev and frozen app (PyInstaller)
if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS  # type: ignore[attr-defined]
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, 'templates')
STATIC_PATH = os.path.join(BASE_DIR, 'static')
app = Flask(__name__, template_folder=TEMPLATE_PATH, static_folder=STATIC_PATH)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-key")

# ---- Jinja filters for number/currency blank if 0 ----
@app.template_filter("num_blank0")
def jinja_num_blank0(val, decimals=0):
    try:
        if val is None:
            return ""
        # NaN check
        if isinstance(val, float) and math.isnan(val):
            return ""
        if float(val) == 0:
            return ""
        fmt = "{:,.%df}" % int(decimals)
        return fmt.format(float(val))
    except Exception:
        return ""

@app.template_filter("currency_blank0")
def jinja_currency_blank0(val, decimals=0):
    s = jinja_num_blank0(val, decimals)
    return f"${s}" if s else ""

# Excel-style rounding (ROUND_HALF_UP) to match Excel's ROUND behavior
def excel_round(value, digits=0):
    try:
        q = Decimal('1') if digits == 0 else Decimal(f'1e-{digits}')
        return float(Decimal(str(value)).quantize(q, rounding=ROUND_HALF_UP))
    except Exception:
        # Fallback: return original value if rounding fails
        return value


def _delete_old_artifacts(proposal_folder: str):
    """Remove generated files before regenerating."""
    patterns = [
        os.path.join(proposal_folder, "Gaco S42 Proposal - *.docx"),
        os.path.join(proposal_folder, "Uniflex Proposal - *.docx"),
        os.path.join(proposal_folder, "*.pdf"),
        os.path.join(proposal_folder, "Profit Summary - *.xlsm"),
    ]
    for patt in patterns:
        for path in glob.glob(patt):
            try:
                os.remove(path)
            except Exception as e:
                print(f"Warning: could not remove {path}: {e}")

def _libreoffice_convert_sync(doc_path: str, outdir: str, timeout: int = 180):
    """
    Convert a DOCX to PDF using LibreOffice headless.
    Blocks until done (or raises on failure).
    """
    os.makedirs(outdir, exist_ok=True)
    cmd = [
        LIBREOFFICE_PATH,
        "--headless",
        "--convert-to", "pdf",
        "--outdir", outdir,
        doc_path,
    ]
    try:
        completed = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=timeout,
            check=True,
            text=True,
        )
        if completed.returncode != 0:
            raise RuntimeError(f"LibreOffice returned {completed.returncode}: {completed.stderr.strip()}")
    except FileNotFoundError:
        raise FileNotFoundError(
            f"LibreOffice not found at {LIBREOFFICE_PATH}. "
            "Install it from libreoffice.org and update LIBREOFFICE_PATH if needed."
        )
    except subprocess.CalledProcessError as e:
        raise RuntimeError(f"LibreOffice conversion failed: {e.stderr or e.stdout}")


def _convert_to_pdf(doc_path: str, outdir: str, use_libreoffice: bool = True, async_mode: bool = True):
    """
    Dispatch PDF conversion. If LibreOffice is available and requested, use it;
    otherwise fall back to docx2pdf (Word). Optionally run async so the UI returns immediately.
    """
    def _worker():
        try:
            if use_libreoffice and os.path.exists(LIBREOFFICE_PATH):
                _libreoffice_convert_sync(doc_path, outdir)
            else:
                # Fallback to Word/docx2pdf (may pop Word)
                convert(doc_path, outdir)
        except Exception as e:
            print(f"PDF conversion failed: {e}")

    if async_mode:
        threading.Thread(target=_worker, daemon=True).start()
    else:
        _worker()

# Base prices
PCS_BASE_LABOR_RATE = 3250
GACO_S42_BASE_PRICE = 210
GACO_PATCH_BASE_PRICE = 125
BLEED_TRAP_BASE_PRICE = 168
DRAINAGE_MAT_BASE_PRICE = 150
UNIFLEX_BASE_PRICE = 240
SW_1FLASH_BASE_PRICE = 162
SW_BLEED_BLOCK_BASE_PRICE = 100
GACO_FOAM_BASE_PRICE = 2430
UNIFLEX_FOAM_BASE_PRICE = 2490
RFC_LABOR_RATE = 250
BASE_OFFICE_FEE_PCT = 0.03
DAVIDS_OFFICE_FEE_PCT = 0.05
PROFIT_SHARE_PCT = 0.10

COMMISSION_PCT = 0.10

# ---- Central Excel mapping & writer ----
EXCEL_CELL_MAP = {
    # Header (known)
    "customer_name": "C1",
    "street_address": "H1",
    "city": "N1",
    "state": "S1",
    "zip_code": "U1",
    "squares": "E3",
    "current_roof": "E5",
    "product": "H3",
    "warranty_incl": "H5",
    "submitted_by": "H7",
    "price_per_sq_10": "M3",
    "price_per_sq_15": None,  # do not update
    "price_per_sq_20": None,  # do not update
    "labor_days": "E7",
    "total_price_10": None,   # do not update
    "total_price_15": None,   # do not update
    "total_price_20": None,   # do not update
    "silicone_units_10": "C11",
    "gaco_patch_units": "C12",
    "bleed_trap_units": "C13",
    "sw_1flash_units": "C14",
    "sw_bleed_block_units": "C15",
    "drainage_mat_units": "C16",
    "foam_units": "C17",
    "silicone_price": "D11",
    "gaco_patch_price": "D12",
    "bleed_trap_price": "D13",
    "sw_1flash_price": "D14",
    "sw_bleed_block_price": "D15",
    "drainage_mat_price": "D16",
    "foam_price": "D17",
    "rfc_labor_price": "D18",
    "pcs_labor_price": "D20",
    "scarifying_total": "E19",
    "travel_total": "E21",
    "misc_costs_total": "E22",
    "adjusted_coverage": None, 
    "office_fee_pct": None,     
    "proposal_note": "C40",     
    "proposal_language": "C41",  
}

def write_fields_to_profit_summary(wb_profit, data: dict):
    """
    Writes values from `data` to the first sheet of wb_profit based on EXCEL_CELL_MAP.
    Fields with mapping None are skipped.
    """
    sht = wb_profit.sheets[0]
    for field, cell in EXCEL_CELL_MAP.items():
        if cell:
            sht.range(cell).value = data.get(field)

# ---- Blank defaults for starting without Excel ----
def make_blank_data():
    return {
        "squares": 0,
        "product": "Uniflex",            # default per request
        "current_roof": "",               # force user to choose
        "warranty_incl": "Yes",          # default Yes because product is Uniflex
        "labor_days": 0,
        "commission_pct": 0,
        "submitted_by": "",               # force user to choose
        "price_per_sq_10": 0,
        "price_per_sq_15": 0,
        "price_per_sq_20": 0,
        "total_price_10": 0,
        "total_price_15": 0,
        "total_price_20": 0,
        "silicone_units_10": 0,
        "silicone_price": 0,
        "gaco_patch_units": 0,
        "gaco_patch_price": 0,
        "bleed_trap_units": 0,
        "bleed_trap_price": 0,
        "sw_1flash_units": 0,
        "sw_1flash_price": 0,
        "sw_bleed_block_units": 0,
        "sw_bleed_block_price": 0,
        "drainage_mat_units": 0,
        "drainage_mat_price": 0,
        "foam_units": 0,
        "foam_price": 0,
        "rfc_labor_price": 0,
        "pcs_labor_price": 0,
        "scarifying_total": 0,
        "travel_total": 0,
        "misc_costs_total": 0,
        "warranty_10_total": 0,
        "office_fee_total": 0,
        "total_cost": 0,
        "pcs_labor_total": 0,
        "rfc_labor_total": 0,
        "pcs_profit": 0,
        "profit_pct": 0,
        "daily_profit": 0,
        "profit_share": 0,
        "commission_amt": 0,
        "coverage_10": 0,
        "coverage_15": 0,
        "coverage_20": 0,
        "adjusted_coverage": 0,
        "office_fee_pct": None,  # None so calc uses Submitted By default
        "previous_squares": 0,
        "previous_roof_type": "",         
        "previous_product": "Uniflex",
        "previous_warranty_incl": "Yes",
        "previous_adjusted_coverage": 0,
        "previous_submitted_by": "",
        "proposal_note": "",
        "street_address": "",
        "city": "",
        "state": "",
        "zip_code": "",
    }


# Coverage amounts for Gaco and Uniflex by roof type and warranty duration
roof_types = ["TPO/EPDM", "Metal", "Mod Bit", "Ballasted 60 mil", "Ballasted 45 mil", "Rock/Foam/Coat"]

# Pricing arrays moved here for global access
pricing10 = [330, 335, 340, 480, 575, 690]
pricing15 = [370, 375, 380, 520, 615, 730]
pricing20 = [410, 415, 420, 560, 655, 770]

coverage_amounts = {
    "Gaco": {
        "TPO/EPDM":    {10: 1.25, 15: 1.75, 20: 2.25},
        "Metal":       {10: 1.25, 15: 1.75, 20: 2.25},
        "Mod Bit":     {10: 1.25, 15: 1.75, 20: 2.25},
        "Ballasted 60 mil": {10: 2.5, 15: 3.25, 20: 3.75},
        "Ballasted 45 mil": {10: 3.0, 15: 4.5,  20: 5.5},
        "Rock/Foam/Coat":   {10: 1.25, 15: 1.75, 20: 2.25},
    },
    "Uniflex": {
        "TPO/EPDM":    {10: 1.5,  15: 2.0,  20: 2.5},
        "Metal":       {10: 1.5,  15: 2.0,  20: 2.5},
        "Mod Bit":     {10: 1.5,  15: 2.0,  20: 2.5},
        "Ballasted 60 mil": {10: 3.0, 15: 3.5,  20: 4.0},
        "Ballasted 45 mil": {10: 3.5, 15: 5.0,  20: 6.0},
        "Rock/Foam/Coat":   {10: 1.5,  15: 2.0,  20: 2.5},
    }
}

def calculation_routine(
    squares,
    product,
    roof_type,
    labor_days,
    warranty_incl,
    price_per_sq_10,
    commission_pct,
    submitted_by,
    previous_submitted_by,
    office_fee_pct,
    adjusted_coverage,
    silicone_units_10,
    silicone_price,
    gaco_patch_units,
    gaco_patch_price,
    sw_1flash_units,
    sw_1flash_price,
    bleed_trap_units,
    bleed_trap_price,
    sw_bleed_block_units,
    sw_bleed_block_price,
    drainage_mat_units,
    drainage_mat_price,
    foam_units,
    foam_price,
    rfc_labor_price,
    pcs_labor_price,
    scarifying_total,
    travel_total,
    misc_costs_total,
    previous_squares,
    previous_roof_type,
    previous_product,
    previous_adjusted_coverage,
    previous_silicone_units_10,
    proposal_note
):
    # Labor days logic
    if roof_type in ["Ballasted 60 mil", "Ballasted 45 mil"]:
        base_labor_days = math.ceil(squares / 30)
    else:
        base_labor_days = math.ceil(squares / 45)

    labor_days_recalc = (previous_roof_type != roof_type) or (previous_squares != squares)

    if labor_days_recalc:
        labor_days = base_labor_days
    else:
        if labor_days is None or (isinstance(labor_days, float) and math.isnan(labor_days)):
            labor_days = base_labor_days

    # Set price_per_sq_* with safe defaults. Allow user override only for 10-yr price.
    # 15/20 are always derived from the pricing tables based on roof_type.
    def _is_blank_zero_or_nan(v):
        if v is None:
            return True
        if isinstance(v, float) and math.isnan(v):
            return True
        try:
            return float(v) == 0.0
        except (TypeError, ValueError):
            return True

    try:
        roof_type_index = roof_types.index(roof_type)
        base_pps10 = pricing10[roof_type_index]
        base_pps15 = pricing15[roof_type_index]
        base_pps20 = pricing20[roof_type_index]
    except ValueError:
        # Unknown roof type: fall back to zeros to avoid UnboundLocalError
        base_pps10 = 0
        base_pps15 = 0
        base_pps20 = 0

    # If the roof type changed, reset 10/15/20 to base.
    if previous_roof_type != roof_type:
        price_per_sq_10 = base_pps10
        price_per_sq_15 = base_pps15
        price_per_sq_20 = base_pps20
    else:
        # Respect user-entered 10-yr price when provided; otherwise use base.
        user_pps10 = price_per_sq_10 if not _is_blank_zero_or_nan(price_per_sq_10) else base_pps10
        # Apply the same delta from base_10 to 15 and 20 so they "recalculate" in line with the override
        delta10 = 0.0
        try:
            delta10 = float(user_pps10) - float(base_pps10)
        except Exception:
            delta10 = 0.0
        price_per_sq_10 = user_pps10
        price_per_sq_15 = float(base_pps15) + delta10
        price_per_sq_20 = float(base_pps20) + delta10

    # Look up coverage factors
    coverage_factors = coverage_amounts.get(product, {}).get(roof_type, {})
    coverage_10 = coverage_factors.get(10, 0)
    coverage_15 = coverage_factors.get(15, 0)
    coverage_20 = coverage_factors.get(20, 0)

    # Apply adjusted coverage
    try:
        adj = 0 if adjusted_coverage is None else float(adjusted_coverage)
    except (TypeError, ValueError):
        adj = 0
    if not (isinstance(adj, float) and math.isnan(adj)) and adj != 0:
        coverage_10 += adj
        coverage_15 += adj
        coverage_20 += adj

    # Calculate current coverage-based units
    calc_units_10 = (squares / 5) * coverage_10
    calc_units_15 = (squares / 5) * coverage_15
    calc_units_20 = (squares / 5) * coverage_20

    # Silicone units logic
    def _norm_adj(v):
        if v is None or (isinstance(v, float) and math.isnan(v)):
            return 0.0
        try:
            return float(v)
        except (TypeError, ValueError):
            return 0.0

    def _almost_equal(a, b, tol=1e-6):
        try:
            return abs(float(a) - float(b)) <= tol
        except Exception:
            return False

    # Detect manual change of silicone units in THIS submit
    user_changed_units = (
        silicone_units_10 is not None
        and not (isinstance(silicone_units_10, float) and math.isnan(silicone_units_10))
        and not _almost_equal(silicone_units_10, previous_silicone_units_10)
    )

    # If the user manually entered silicone units, adjusted_coverage is ignored/reset
    if user_changed_units:
        adjusted_coverage = 0.0
        base_cov_factors = coverage_amounts.get(product, {}).get(roof_type, {})
        coverage_10 = base_cov_factors.get(10, 0)
        coverage_15 = base_cov_factors.get(15, 0)
        coverage_20 = base_cov_factors.get(20, 0)
        calc_units_10 = (squares / 5) * coverage_10
        calc_units_15 = (squares / 5) * coverage_15
        calc_units_20 = (squares / 5) * coverage_20

    # Recalc when product, roof_type, or squares change (always). For adjusted_coverage changes,
    # only recalc if the user did NOT manually override silicone units.
    recalc_trigger = (
        (previous_product != product)
        or (previous_roof_type != roof_type)
        or (squares != previous_squares)
        or ((not user_changed_units) and (_norm_adj(adjusted_coverage) != _norm_adj(previous_adjusted_coverage)))
    )

    if recalc_trigger:
        silicone_units_10 = calc_units_10
    else:
        def _is_blank_zero_or_nan(v):
            if v is None:
                return True
            if isinstance(v, float) and math.isnan(v):
                return True
            try:
                return float(v) == 0.0
            except (TypeError, ValueError):
                return True
        if _is_blank_zero_or_nan(silicone_units_10):
            silicone_units_10 = calc_units_10

    silicone_units_15 = calc_units_15
    silicone_units_20 = calc_units_20

    # If user overrode 10-yr units, derive 15/20 from 10 using coverage ratios
    if user_changed_units:
        if coverage_10:
            silicone_units_15 = silicone_units_10 * (coverage_15 / coverage_10)
            silicone_units_20 = silicone_units_10 * (coverage_20 / coverage_10)
        else:
            silicone_units_15 = silicone_units_10
            silicone_units_20 = silicone_units_10

    # Normalize silicone units to whole numbers by **rounding up** (ceiling)
    try:
        silicone_units_10 = math.ceil(float(silicone_units_10 or 0))
    except Exception:
        silicone_units_10 = 0
    try:
        silicone_units_15 = math.ceil(float(silicone_units_15 or 0))
    except Exception:
        silicone_units_15 = 0
    try:
        silicone_units_20 = math.ceil(float(silicone_units_20 or 0))
    except Exception:
        silicone_units_20 = 0
    
    # Silicone price logic:
    base_silicone_price = (
        GACO_S42_BASE_PRICE if product == "Gaco" else (
            UNIFLEX_BASE_PRICE if product == "Uniflex" else silicone_price
        )
    )
    if previous_product != product:
        silicone_price = base_silicone_price
    else:
        if silicone_price is None or (isinstance(silicone_price, float) and math.isnan(silicone_price)):
            silicone_price = base_silicone_price

    # Gaco patch units logic (units depend on product & squares)
    if product == "Gaco":
        base_gaco_patch_units = math.ceil(squares / 10)
    else:
        base_gaco_patch_units = 0

    gaco_patch_recalc = (previous_product != product) or (previous_squares != squares)

    if gaco_patch_recalc:
        gaco_patch_units = base_gaco_patch_units
    else:
        if gaco_patch_units is None or (isinstance(gaco_patch_units, float) and math.isnan(gaco_patch_units)):
            gaco_patch_units = base_gaco_patch_units

    # Gaco patch price logic
    base_gaco_patch_price = GACO_PATCH_BASE_PRICE if product == "Gaco" else 0

    if previous_product != product:
        gaco_patch_price = base_gaco_patch_price
    else:
        if gaco_patch_price is None or (isinstance(gaco_patch_price, float) and math.isnan(gaco_patch_price)):
            gaco_patch_price = base_gaco_patch_price
    
    # Bleed Trap logic (units & price)
    if product == "Gaco" and roof_type == "Mod Bit":
        base_bleed_units = math.ceil(squares / 5)
        base_bleed_price = BLEED_TRAP_BASE_PRICE
    else:
        base_bleed_units = 0
        base_bleed_price = 0

    bleed_recalc_trigger = (
        (previous_product != product)
        or (previous_roof_type != roof_type)
        or (previous_squares != squares)
    )

    if bleed_recalc_trigger:
        bleed_trap_units = base_bleed_units
        bleed_trap_price = base_bleed_price
    else:
        if bleed_trap_units is None or (isinstance(bleed_trap_units, float) and math.isnan(bleed_trap_units)):
            bleed_trap_units = base_bleed_units
        if bleed_trap_price is None or (isinstance(bleed_trap_price, float) and math.isnan(bleed_trap_price)):
            bleed_trap_price = base_bleed_price

    # SW 1-Flash logic (units & price)
    if product == "Uniflex":
        base_sw_1flash_price = SW_1FLASH_BASE_PRICE
        if roof_type in ["TPO/EPDM", "Mod Bit", "Rock/Foam/Coat"]:
            base_sw_1flash_units = math.ceil(squares / 20)  # rounded up, no decimals
        else:
            base_sw_1flash_units = math.ceil(squares / 10)  # rounded up, no decimals
    else:
        base_sw_1flash_price = 0
        base_sw_1flash_units = 0

    sw1_recalc_trigger = (
        (previous_product != product)
        or (previous_roof_type != roof_type)
        or (previous_squares != squares)
    )

    if sw1_recalc_trigger:
        sw_1flash_units = base_sw_1flash_units
        sw_1flash_price = base_sw_1flash_price
    else:
        if sw_1flash_units is None or (isinstance(sw_1flash_units, float) and math.isnan(sw_1flash_units)):
            sw_1flash_units = base_sw_1flash_units
        if sw_1flash_price is None or (isinstance(sw_1flash_price, float) and math.isnan(sw_1flash_price)):
            sw_1flash_price = base_sw_1flash_price

    # SW Bleed Block logic (units & price)
    if product == "Uniflex" and roof_type == "Mod Bit":
        base_sw_bleed_block_units = math.ceil(squares / 5)
        base_sw_bleed_block_price = SW_BLEED_BLOCK_BASE_PRICE
    else:
        base_sw_bleed_block_units = 0
        base_sw_bleed_block_price = 0

    sw_bleed_block_recalc = (
        (previous_product != product)
        or (previous_roof_type != roof_type)
        or (previous_squares != squares)
        or (previous_adjusted_coverage != adjusted_coverage)
    )

    if sw_bleed_block_recalc:
        sw_bleed_block_units = base_sw_bleed_block_units
        sw_bleed_block_price = base_sw_bleed_block_price
    else:
        if sw_bleed_block_units is None or (isinstance(sw_bleed_block_units, float) and math.isnan(sw_bleed_block_units)):
            sw_bleed_block_units = base_sw_bleed_block_units
        if sw_bleed_block_price is None or (isinstance(sw_bleed_block_price, float) and math.isnan(sw_bleed_block_price)):
            sw_bleed_block_price = base_sw_bleed_block_price

    # Drainage Mat logic (units & price)
    if roof_type in ["Ballasted 60 mil", "Ballasted 45 mil"]:
        base_drainage_units = math.ceil(squares / 18)
        base_drainage_price = DRAINAGE_MAT_BASE_PRICE
    else:
        base_drainage_units = 0
        base_drainage_price = 0

    drainage_recalc = (previous_roof_type != roof_type) or (previous_squares != squares)

    if drainage_recalc:
        drainage_mat_units = base_drainage_units
        drainage_mat_price = base_drainage_price
    else:
        if drainage_mat_units is None or (isinstance(drainage_mat_units, float) and math.isnan(drainage_mat_units)):
            drainage_mat_units = base_drainage_units
        if drainage_mat_price is None or (isinstance(drainage_mat_price, float) and math.isnan(drainage_mat_price)):
            drainage_mat_price = base_drainage_price

    # Foam logic (units & price)
    if roof_type == "Rock/Foam/Coat":
        base_foam_units = math.ceil(squares / 25)  # rounded up, no decimals
        if product == "Gaco":
            base_foam_price = GACO_FOAM_BASE_PRICE
        elif product == "Uniflex":
            base_foam_price = UNIFLEX_FOAM_BASE_PRICE
        else:
            base_foam_price = 0
    else:
        base_foam_units = 0
        base_foam_price = 0

    # Recalc foam when roof type OR product OR squares changes so base price updates correctly
    foam_recalc = (
        (previous_roof_type != roof_type)
        or (previous_product != product)
        or (previous_squares != squares)
    )

    if foam_recalc:
        foam_units = base_foam_units
        foam_price = base_foam_price
    else:
        if foam_units is None or (isinstance(foam_units, float) and math.isnan(foam_units)):
            foam_units = base_foam_units
        if foam_price is None or (isinstance(foam_price, float) and math.isnan(foam_price)):
            foam_price = base_foam_price

    # RFC labor price logic (aka rfc_price)
    base_rfc_price = RFC_LABOR_RATE if roof_type == "Rock/Foam/Coat" else 0

    rfc_recalc = (previous_roof_type != roof_type)

    if rfc_recalc:
        rfc_labor_price = base_rfc_price
    else:
        if rfc_labor_price is None or (isinstance(rfc_labor_price, float) and math.isnan(rfc_labor_price)):
            rfc_labor_price = base_rfc_price

    # PCS labor price logic
    base_pcs_labor_price = PCS_BASE_LABOR_RATE
    if (
        pcs_labor_price is None
        or (isinstance(pcs_labor_price, float) and math.isnan(pcs_labor_price))
        or pcs_labor_price == 0
    ):
        pcs_labor_price = base_pcs_labor_price

    # Ensure all units and per-unit prices are whole numbers before multiplying
    silicone_total       = excel_round(silicone_units_10, 0)        * excel_round(silicone_price, 0)
    gaco_patch_total     = excel_round(gaco_patch_units, 0)         * excel_round(gaco_patch_price, 0)
    bleed_trap_total     = excel_round(bleed_trap_units, 0)         * excel_round(bleed_trap_price, 0)
    sw_bleed_block_total = excel_round(sw_bleed_block_units, 0)     * excel_round(sw_bleed_block_price, 0)
    sw_1flash_total      = excel_round(sw_1flash_units, 0)          * excel_round(sw_1flash_price, 0)
    drainage_mat_total   = excel_round(drainage_mat_units, 0)       * excel_round(drainage_mat_price, 0)
    foam_total           = excel_round(foam_units, 0)               * excel_round(foam_price, 0)

    # Labor totals remain as-is (they are rate * quantity, not unit-count * unit-price pairs)
    rfc_labor_total = rfc_labor_price * squares
    pcs_labor_total = pcs_labor_price * labor_days

    # --- Enforce warranty_incl based on product rules ---
    if product == "Uniflex":
        warranty_incl = "Yes"

    # Normalize once for comparisons below
    warranty_flag = (warranty_incl or "No").strip().lower()

    if warranty_flag != "yes":
        warranty_10_total = 0
        warranty_15_total = 0
        warranty_20_total = 0
    else:
        if product == "Uniflex":
            warranty_10_total = 500
            warranty_15_total = 500
            warranty_20_total = 500
        elif product == "Gaco":
            warranty_10_total = 750 if squares < 75 else 10 * squares
            warranty_15_total = 1125 if squares < 75 else 15 * squares
            warranty_20_total = 1500 if squares < 75 else 20 * squares
        else:
            warranty_10_total = 0
            warranty_15_total = 0
            warranty_20_total = 0

    # --- Office Fee % effective value ---
    def _blank_or_nan(v):
        try:
            if v is None:
                return True
            if isinstance(v, float) and math.isnan(v):
                return True
            return float(v) == 0.0
        except Exception:
            return True

    # Re-evaluate Office Fee %
    if submitted_by != previous_submitted_by:
        office_fee_pct = DAVIDS_OFFICE_FEE_PCT if (submitted_by == "David Estes") else BASE_OFFICE_FEE_PCT
    else:
        if _blank_or_nan(office_fee_pct):
            office_fee_pct = DAVIDS_OFFICE_FEE_PCT if (submitted_by == "David Estes") else BASE_OFFICE_FEE_PCT
        else:
            office_fee_pct = float(office_fee_pct)

    effective_office_fee_pct = office_fee_pct

    # Total Price logic (moved after warranty totals are set)
    total_price_10 = (
        (squares * price_per_sq_10)
        + warranty_10_total
        + (travel_total or 0)
        + (misc_costs_total or 0)
    )
    total_price_15 = (
        (squares * price_per_sq_15)
        + warranty_15_total
        + (travel_total or 0)
        + (misc_costs_total or 0)
    )
    total_price_20 = (
        (squares * price_per_sq_20)
        + warranty_20_total
        + (travel_total or 0)
        + (misc_costs_total or 0)
    )

    office_fee_total = excel_round(total_price_10 * effective_office_fee_pct, 0)

    # --- Commission percent & amount ---
    if submitted_by in ("David Estes", "Vern Abbott"):
        commission_pct = COMMISSION_PCT
    else:
        commission_pct = 0.0
    commission_amt = excel_round(commission_pct * total_price_10, 0)

    total_cost = sum([
        silicone_total,
        gaco_patch_total,
        bleed_trap_total,
        sw_1flash_total,
        sw_bleed_block_total,
        drainage_mat_total,
        foam_total,
        rfc_labor_total,
        pcs_labor_total,
        scarifying_total,
        travel_total,
        misc_costs_total,
        warranty_10_total,
        office_fee_total,
        commission_amt
    ])

    profit_share_amt = excel_round(PROFIT_SHARE_PCT * (total_price_10 - total_cost), 0)
    pcs_profit = total_price_10 - total_cost - profit_share_amt
    profit_pct = excel_round(pcs_profit / total_price_10, 2) if total_price_10 else 0
    daily_profit = excel_round(pcs_profit / labor_days, 0) if labor_days else 0

    result = {
        "labor_days": labor_days,
        "submitted_by": submitted_by,
        "price_per_sq_10": price_per_sq_10,
        "price_per_sq_15": price_per_sq_15,
        "price_per_sq_20": price_per_sq_20,
        "total_price_10": total_price_10,
        "total_price_15": total_price_15,
        "total_price_20": total_price_20,
        "silicone_units_10": silicone_units_10,
        "silicone_price": silicone_price,
        "silicone_total": silicone_total,
        "gaco_patch_units": gaco_patch_units,
        "gaco_patch_price": gaco_patch_price,
        "gaco_patch_total": gaco_patch_total,
        "bleed_trap_units": bleed_trap_units,
        "bleed_trap_price": bleed_trap_price,
        "bleed_trap_total": bleed_trap_total,
        "sw_1flash_units": sw_1flash_units,
        "sw_1flash_price": sw_1flash_price,
        "sw_1flash_total": sw_1flash_total,
        "sw_bleed_block_units": sw_bleed_block_units,
        "sw_bleed_block_price": sw_bleed_block_price,
        "sw_bleed_block_total": sw_bleed_block_total,
        "drainage_mat_units": drainage_mat_units,
        "drainage_mat_price": drainage_mat_price,
        "drainage_mat_total": drainage_mat_total,
        "foam_units": foam_units,
        "foam_price": foam_price,
        "foam_total": foam_total,
        "rfc_labor_price": rfc_labor_price,
        "pcs_labor_price": pcs_labor_price,
        "rfc_labor_total": rfc_labor_total,
        "pcs_labor_total": pcs_labor_total,
        "scarifying_total": scarifying_total,
        "travel_total": travel_total,
        "misc_costs_total": misc_costs_total,
        "office_fee_total": office_fee_total,
        "pcs_profit": pcs_profit,
        "profit_pct": profit_pct,
        "daily_profit": daily_profit,
        "profit_share": profit_share_amt,
        "warranty_10_total": warranty_10_total,
        "warranty_15_total": warranty_15_total,
        "warranty_20_total": warranty_20_total,
        "coverage_10": coverage_10,
        "coverage_15": coverage_15,
        "coverage_20": coverage_20,
        "silicone_units_15": silicone_units_15,
        "silicone_units_20": silicone_units_20,
        "commission_amt": commission_amt,
        "commission_pct": commission_pct,
        "total_cost": total_cost,
        "warranty_incl": warranty_incl,
        "office_fee_pct": effective_office_fee_pct,
        "adjusted_coverage": adjusted_coverage,
        "previous_submitted_by": submitted_by,
        "previous_roof_type": roof_type,
        "previous_squares": squares,
        "previous_product": product,
        "previous_adjusted_coverage": adjusted_coverage,
        "previous_silicone_units_10": silicone_units_10,
        "proposal_note": proposal_note,
    }
    return result


@app.route('/')
def proposal_list():
    # Which tab is selected: 'open' (default) or 'under'
    status = (request.args.get('status') or 'open').strip().lower()

    # Collect Open Proposals from PROPOSALS_DIR
    try:
        open_folders = [
            f for f in os.listdir(PROPOSALS_DIR)
            if os.path.isdir(os.path.join(PROPOSALS_DIR, f))
        ]
    except Exception:
        open_folders = []

    # Collect Under Contract from CONTRACTS_DIR
    try:
        contract_folders = [
            f for f in os.listdir(CONTRACTS_DIR)
            if os.path.isdir(os.path.join(CONTRACTS_DIR, f))
        ]
    except Exception:
        contract_folders = []

    open_folders.sort(key=str.lower)
    contract_folders.sort(key=str.lower)

    return render_template(
        'proposal_list.html',
        open_folders=open_folders,
        contract_folders=contract_folders,
        status=status,
    )

def find_profit_summary_file(folder_path):
    # Safely handle missing/non-existent folder
    if not folder_path or not os.path.isdir(folder_path):
        return None
    for f in os.listdir(folder_path):
        if f.startswith("Profit Summary") and f.endswith((".xlsm", ".xlsx")):
            return os.path.join(folder_path, f)
    return None



@app.route('/update-proposal/<folder_name>', methods=['POST'])
def update_proposal(folder_name):
    folder_path = os.path.join(PROPOSALS_DIR, folder_name)
    allow_blank = (folder_name in ("NEW", "__blank__"))

    action = (request.form.get('action') or '').strip().lower()

    excel_file = None
    if not allow_blank:
        excel_file = find_profit_summary_file(folder_path)
        if not excel_file:
            return f"No 'Profit Summary' Excel file found in {folder_name}", 404

    # Collect updated data from the form and convert to appropriate types
    def parse_float(val, default=0.0):
        try:
            if val is None:
                return default
            if isinstance(val, str):
                cleaned = val.replace('$', '').replace(',', '').strip()
                if cleaned == '':
                    return default
                return float(cleaned)
            return float(val)
        except (TypeError, ValueError):
            return default

    def parse_int(val, default=0):
        try:
            return int(val)
        except (TypeError, ValueError):
            return default

    # If the Blank Proposal flow hits the Create button, build artifacts and redirect
    if allow_blank and action == 'create':
        # Pull the minimal required fields from the posted form
        customer_name = (request.form.get('customer_name') or '').strip()
        street_address = (request.form.get('street_address') or '').strip()
        city = (request.form.get('city') or '').strip()
        state = (request.form.get('state') or '').strip()
        zip_code = (request.form.get('zip_code') or '').strip()
        roof_type = (request.form.get('current_roof') or request.form.get('roof_type') or '').strip()
        try:
            total_squares = int(parse_float(request.form.get('squares'), 0))
        except Exception:
            total_squares = 0
        warranty_incl = (request.form.get('warranty_incl') or 'No').strip()
        product = (request.form.get('product') or '').strip()
        submitted_by = (request.form.get('submitted_by') or '').strip()
        includes_text = (request.form.get('includes_text') or '').strip()
        proposal_language = (request.form.get('proposal_language') or includes_text or '').strip()

        # Collect any additional mapped fields present on the form for initial write
        def _pf(name, default=None):
            val = request.form.get(name)
            if val is None or str(val).strip() == '':
                return default
            try:
                return float(val.replace('$','').replace(',',''))
            except Exception:
                return val

        mapped_data_full = {
            "price_per_sq_10": _pf("price_per_sq_10"),
            "labor_days": _pf("labor_days"),
            "silicone_units_10": _pf("silicone_units_10"),
            "gaco_patch_units": _pf("gaco_patch_units"),
            "bleed_trap_units": _pf("bleed_trap_units"),
            "sw_1flash_units": _pf("sw_1flash_units"),
            "sw_bleed_block_units": _pf("sw_bleed_block_units"),
            "drainage_mat_units": _pf("drainage_mat_units"),
            "foam_units": _pf("foam_units"),
            "silicone_price": _pf("silicone_price"),
            "gaco_patch_price": _pf("gaco_patch_price"),
            "bleed_trap_price": _pf("bleed_trap_price"),
            "sw_1flash_price": _pf("sw_1flash_price"),
            "sw_bleed_block_price": _pf("sw_bleed_block_price"),
            "drainage_mat_price": _pf("drainage_mat_price"),
            "foam_price": _pf("foam_price"),
            "rfc_labor_price": _pf("rfc_labor_price"),
            "pcs_labor_price": _pf("pcs_labor_price"),
            "scarifying_total": _pf("scarifying_total"),
            "travel_total": _pf("travel_total"),
            "misc_costs_total": _pf("misc_costs_total"),
            "proposal_note": (request.form.get("proposal_note") or "").strip(),
            "proposal_language": (request.form.get("proposal_language") or "").strip(),
            "total_price_10": _pf("total_price_10"),
            "total_price_15": _pf("total_price_15"),
            "total_price_20": _pf("total_price_20"),
        }
        # Remove Nones to avoid overwriting with blanks
        mapped_data_full = {k: v for k, v in mapped_data_full.items() if v is not None}

        # Create artifacts using the helper (same behavior as /new)
        new_folder = create_proposal_from_fields(
            customer_name=customer_name,
            street_address=street_address,
            city=city,
            state=state,
            zip_code=zip_code,
            roof_type=roof_type,
            total_squares=total_squares,
            warranty_incl=warranty_incl,
            product=product,
            proposal_language=proposal_language,
            submitted_by=submitted_by,
            mapped_data=mapped_data_full,
            pdf_async=True,
            use_libreoffice=True,
        )
        return redirect(url_for('proposal_list'))

    squares = parse_float(request.form.get('squares'))
    product = request.form.get('product')
    roof_type = request.form.get('current_roof')
    labor_days = parse_int(request.form.get('labor_days'))
    warranty_incl = request.form.get('warranty_incl', 'No').strip()
    previous_warranty_incl = request.form.get('previous_warranty_incl', warranty_incl)
    price_per_sq_10 = parse_float(request.form.get('price_per_sq_10'))
    commission_pct = parse_float(request.form.get('commission_pct'))

    submitted_by = request.form.get('submitted_by')
    previous_submitted_by = request.form.get('previous_submitted_by', '')

    raw_office_fee_pct = request.form.get('office_fee_pct')
    if raw_office_fee_pct is None or str(raw_office_fee_pct).strip() == '':
        office_fee_pct = None  # allow defaulting based on Submitted By in calc
    else:
        cleaned_office = str(raw_office_fee_pct).replace('%', '').strip()
        office_fee_value = parse_float(cleaned_office)
        if office_fee_value is None:
            office_fee_pct = None
        elif office_fee_value > 1:
            office_fee_pct = office_fee_value / 100.0  # e.g., "5" -> 0.05
        else:
            office_fee_pct = office_fee_value         # already decimal like 0.05

    raw_adjusted_coverage = request.form.get('adjusted_coverage') or request.form.get('adjust_coverage')
    adjusted_coverage = None if raw_adjusted_coverage is None or str(raw_adjusted_coverage).strip() == '' else parse_float(raw_adjusted_coverage)
    raw_silicone_units_10 = request.form.get('silicone_units_10')
    silicone_units_10 = None if raw_silicone_units_10 is None or str(raw_silicone_units_10).strip() == '' else parse_float(raw_silicone_units_10)
    raw_silicone_price = request.form.get('silicone_price')
    silicone_price = None if raw_silicone_price is None or str(raw_silicone_price).strip() == '' else parse_float(raw_silicone_price)
    raw_gaco_patch_units = request.form.get('gaco_patch_units')
    gaco_patch_units = None if raw_gaco_patch_units is None or str(raw_gaco_patch_units).strip() == '' else parse_float(raw_gaco_patch_units)
    raw_gaco_patch_price = request.form.get('gaco_patch_price')
    gaco_patch_price = None if raw_gaco_patch_price is None or str(raw_gaco_patch_price).strip() == '' else parse_float(raw_gaco_patch_price)
    raw_bleed_trap_units = request.form.get('bleed_trap_units') or request.form.get('sw_bleed_trap_units')
    bleed_trap_units = None if raw_bleed_trap_units is None or str(raw_bleed_trap_units).strip() == '' else parse_float(raw_bleed_trap_units)
    raw_bleed_trap_price = request.form.get('bleed_trap_price') or request.form.get('sw_bleed_trap_price')
    bleed_trap_price = None if raw_bleed_trap_price is None or str(raw_bleed_trap_price).strip() == '' else parse_float(raw_bleed_trap_price)
    raw_sw_1flash_units = request.form.get('sw_1flash_units')
    sw_1flash_units = None if raw_sw_1flash_units is None or str(raw_sw_1flash_units).strip() == '' else parse_float(raw_sw_1flash_units)
    raw_sw_1flash_price = request.form.get('sw_1flash_price')
    sw_1flash_price = None if raw_sw_1flash_price is None or str(raw_sw_1flash_price).strip() == '' else parse_float(raw_sw_1flash_price)
    raw_sw_bleed_block_units = request.form.get('sw_bleed_block_units')
    sw_bleed_block_units = None if raw_sw_bleed_block_units is None or str(raw_sw_bleed_block_units).strip() == '' else parse_float(raw_sw_bleed_block_units)
    raw_sw_bleed_block_price = request.form.get('sw_bleed_block_price')
    sw_bleed_block_price = None if raw_sw_bleed_block_price is None or str(raw_sw_bleed_block_price).strip() == '' else parse_float(raw_sw_bleed_block_price)
    raw_drainage_mat_units = request.form.get('drainage_mat_units')
    drainage_mat_units = None if raw_drainage_mat_units is None or str(raw_drainage_mat_units).strip() == '' else parse_float(raw_drainage_mat_units)
    raw_drainage_mat_price = request.form.get('drainage_mat_price')
    drainage_mat_price = None if raw_drainage_mat_price is None or str(raw_drainage_mat_price).strip() == '' else parse_float(raw_drainage_mat_price)
    raw_foam_units = request.form.get('foam_units')
    foam_units = None if raw_foam_units is None or str(raw_foam_units).strip() == '' else parse_float(raw_foam_units)
    raw_foam_price = request.form.get('foam_price')
    foam_price = None if raw_foam_price is None or str(raw_foam_price).strip() == '' else parse_float(raw_foam_price)
    raw_rfc_labor_price = request.form.get('rfc_labor_price')
    rfc_labor_price = None if raw_rfc_labor_price is None or str(raw_rfc_labor_price).strip() == '' else parse_float(raw_rfc_labor_price)
    pcs_labor_price = parse_float(request.form.get('pcs_labor_price'))
    raw_scarifying_total = request.form.get('scarifying_total')
    scarifying_total = parse_float(raw_scarifying_total)
    travel_total = parse_float(request.form.get('travel_total'))
    misc_costs_total = parse_float(request.form.get('misc_costs_total'))
    # Use explicit fallbacks that reflect a prior/blank state so changes are detectable
    _prev_sq_raw = request.form.get('previous_squares')
    previous_squares = parse_float(_prev_sq_raw, 0.0)  # default to 0, not current squares

    previous_roof_type = request.form.get('previous_roof_type', '')  # default to '' so a change is caught

    previous_product = request.form.get('previous_product', '')  # safe default; not used for labor_days but consistent

    _prev_adj_raw = request.form.get('previous_adjusted_coverage')
    previous_adjusted_coverage = parse_float(_prev_adj_raw, 0.0)

    _prev_units_raw = request.form.get('previous_silicone_units_10')
    # Default to current silicone_units_10 if the hidden field is missing on first render
    previous_silicone_units_10 = parse_float(_prev_units_raw, (silicone_units_10 or 0.0))

    # Simple text field; persist across recalcs
    proposal_note = (request.form.get('proposal_note') or '').strip()
    proposal_language = (request.form.get('proposal_language') or '').strip()
    customer_name = (request.form.get('customer_name') or '').strip()
    street_address = (request.form.get('street_address') or '').strip()
    city = (request.form.get('city') or '').strip()
    state = (request.form.get('state') or '').strip()
    zip_code = (request.form.get('zip_code') or '').strip()

    # Use proposal_language as the single source of truth for downstream Word/Excel writes
    includes_text = proposal_language

    # Carry read-only flag through POST round-trips, supporting both new and legacy formats
    # New: read_only = "Yes"/"No"; Legacy: readonly = "1"/"0"
    read_only_param = request.form.get('read_only')
    if read_only_param is not None:
        readonly = (read_only_param.strip().lower() == 'yes')
    else:
        readonly = (request.form.get('readonly') == '1')

    # Prepare data dictionary for template (may include more fields as needed)
    data = {
        'squares': squares,
        'product': product,
        'current_roof': roof_type,
        'labor_days': labor_days,
        'warranty_incl': warranty_incl,
        'price_per_sq_10': price_per_sq_10,
        'commission_pct': commission_pct,
        'adjusted_coverage': adjusted_coverage,
        'silicone_units_10': silicone_units_10,
        'silicone_price': silicone_price,
        'gaco_patch_units': gaco_patch_units,
        'gaco_patch_price': gaco_patch_price,
        'sw_1flash_units': sw_1flash_units,
        'sw_1flash_price': sw_1flash_price,
        'bleed_trap_units': bleed_trap_units,
        'bleed_trap_price': bleed_trap_price,
        'sw_bleed_block_units': sw_bleed_block_units,
        'sw_bleed_block_price': sw_bleed_block_price,
        'drainage_mat_units': drainage_mat_units,
        'drainage_mat_price': drainage_mat_price,
        'foam_units': foam_units,
        'foam_price': foam_price,
        'rfc_labor_price': rfc_labor_price,
        'pcs_labor_price': pcs_labor_price,
        'scarifying_total': scarifying_total,
        'travel_total': travel_total,
        'misc_costs_total': misc_costs_total,
        'previous_squares': previous_squares,
        'previous_roof_type': previous_roof_type,
        'previous_product': previous_product,
        'previous_warranty_incl': previous_warranty_incl,
        'previous_adjusted_coverage': previous_adjusted_coverage,
        'previous_silicone_units_10': previous_silicone_units_10,
        'coverage_10': 0,
        'coverage_15': 0,
        'coverage_20': 0,
        'submitted_by': submitted_by,
        'office_fee_pct': office_fee_pct,
        'previous_submitted_by': previous_submitted_by,
        'proposal_note': proposal_note,
        'proposal_language': proposal_language,
        'customer_name': customer_name,
        'street_address': street_address,
        'city': city,
        'state': state,
        'zip_code': zip_code,
        'includes_text': includes_text,
    }

    # If saving an existing proposal, delete old artifacts and regenerate in the same folder
    if action == 'save' and not allow_blank and folder_name:
        proposal_folder = folder_path
        _delete_old_artifacts(proposal_folder)
        # Collect any additional mapped fields present on the form for initial write
        def _pf(name, default=None):
            val = request.form.get(name)
            if val is None or str(val).strip() == '':
                return default
            try:
                return float(val.replace('$','').replace(',',''))
            except Exception:
                return val

        mapped_data_full = {
            "price_per_sq_10": _pf("price_per_sq_10"),
            "labor_days": _pf("labor_days"),
            "silicone_units_10": _pf("silicone_units_10"),
            "gaco_patch_units": _pf("gaco_patch_units"),
            "bleed_trap_units": _pf("bleed_trap_units"),
            "sw_1flash_units": _pf("sw_1flash_units"),
            "sw_bleed_block_units": _pf("sw_bleed_block_units"),
            "drainage_mat_units": _pf("drainage_mat_units"),
            "foam_units": _pf("foam_units"),
            "silicone_price": _pf("silicone_price"),
            "gaco_patch_price": _pf("gaco_patch_price"),
            "bleed_trap_price": _pf("bleed_trap_price"),
            "sw_1flash_price": _pf("sw_1flash_price"),
            "sw_bleed_block_price": _pf("sw_bleed_block_price"),
            "drainage_mat_price": _pf("drainage_mat_price"),
            "foam_price": _pf("foam_price"),
            "rfc_labor_price": _pf("rfc_labor_price"),
            "pcs_labor_price": _pf("pcs_labor_price"),
            "scarifying_total": _pf("scarifying_total"),
            "travel_total": _pf("travel_total"),
            "misc_costs_total": _pf("misc_costs_total"),
            "proposal_note": (request.form.get("proposal_note") or "").strip(),
            "proposal_language": (request.form.get("proposal_language") or "").strip(),
            "total_price_10": _pf("total_price_10"),
            "total_price_15": _pf("total_price_15"),
            "total_price_20": _pf("total_price_20"),
        }
        # Remove Nones to avoid overwriting with blanks
        mapped_data_full = {k: v for k, v in mapped_data_full.items() if v is not None}

        create_proposal_from_fields(
            customer_name=customer_name,
            street_address=street_address,
            city=city,
            state=state,
            zip_code=zip_code,
            roof_type=roof_type,
            total_squares=int(squares) if squares else 0,
            warranty_incl=warranty_incl,
            product=product,
            proposal_language=proposal_language,
            submitted_by=submitted_by,
            target_folder=proposal_folder,
            mapped_data=mapped_data_full,
            pdf_async=True,
            use_libreoffice=True,
        )
        return redirect(url_for('proposal_list'))

    # Call calculation_routine and merge results
    calc_result = calculation_routine(
        squares,
        product,
        roof_type,
        labor_days,
        warranty_incl,
        price_per_sq_10,
        commission_pct,
        submitted_by=submitted_by,
        previous_submitted_by=previous_submitted_by,
        office_fee_pct=office_fee_pct,
        adjusted_coverage=adjusted_coverage,
        silicone_units_10=silicone_units_10,
        silicone_price=silicone_price,
        gaco_patch_units=gaco_patch_units,
        gaco_patch_price=gaco_patch_price,
        sw_1flash_units=sw_1flash_units,
        sw_1flash_price=sw_1flash_price,
        bleed_trap_units=bleed_trap_units,
        bleed_trap_price=bleed_trap_price,
        sw_bleed_block_units=sw_bleed_block_units,
        sw_bleed_block_price=sw_bleed_block_price,
        drainage_mat_units=drainage_mat_units,
        drainage_mat_price=drainage_mat_price,
        foam_units=foam_units,
        foam_price=foam_price,
        rfc_labor_price=rfc_labor_price,
        pcs_labor_price=pcs_labor_price,
        scarifying_total=scarifying_total,
        travel_total=travel_total,
        misc_costs_total=misc_costs_total,
        previous_squares=previous_squares,
        previous_roof_type=previous_roof_type,
        previous_product=previous_product,
        previous_adjusted_coverage=previous_adjusted_coverage,
        previous_silicone_units_10=previous_silicone_units_10,
        proposal_note=proposal_note
    )
    # Persist key header fields and note across round trip so they are not lost
    calc_result.update({
        "customer_name": customer_name,
        "street_address": street_address,
        "city": city,
        "state": state,
        "zip_code": zip_code,
        "proposal_note": proposal_note,
        "proposal_language": includes_text,
        "includes_text": includes_text,
    })

    data.update(calc_result)
    return render_template(
        "proposal_details.html",
        data=data,
        folder_name=folder_name,
        readonly=readonly,
        is_blank=(folder_name in ("NEW", "__blank__"))
    )

@app.route('/proposal_details/new', methods=['GET'])
def proposal_details_new():
    data = make_blank_data()
    # Support both new and legacy readonly query parameters
    read_only_param = request.args.get('read_only')
    if read_only_param is not None:
        readonly = (read_only_param.strip().lower() == 'yes')
    else:
        readonly = (request.args.get('readonly') == '1')
    return render_template(
        "proposal_details.html",
        data=data,
        folder_name="NEW",
        readonly=readonly,
        is_blank=True,
    )

@app.route('/proposal_details/<folder_name>')
def proposal_details(folder_name):
    # Serve blank form if requested
    if folder_name == "__blank__":
        data = make_blank_data()
        # Reuse the blank POST flow by setting folder_name to NEW for the template's form action
        # Support both new and legacy readonly query parameters
        read_only_param = request.args.get('read_only')
        if read_only_param is not None:
            readonly = (read_only_param.strip().lower() == 'yes')
        else:
            readonly = (request.args.get('readonly') == '1')
        return render_template(
            "proposal_details.html",
            data=data,
            folder_name="NEW",
            readonly=readonly,
            is_blank=True,
        )

    # --- Deadfile indicator logic ---
    dead_ind = request.args.get('dead_ind')
    if dead_ind is not None and str(dead_ind).strip().lower() in ('yes', 'true', '1'):
        # Attempt to move the folder from proposals to deadfile
        src_path = os.path.join(PROPOSALS_DIR, folder_name)
        dest_path = os.path.join(DEADFILE_DIR, folder_name)
        try:
            # Check if source exists
            if not os.path.exists(src_path):
                flash(f"Source folder '{src_path}' does not exist.", "error")
            elif os.path.exists(dest_path):
                flash(f"Target folder '{dest_path}' already exists.", "error")
            else:
                shutil.move(src_path, dest_path)
                flash(f"Proposal '{folder_name}' moved to dead file.", "success")
        except Exception as e:
            flash(f"Error moving proposal: {e}", "error")
        return redirect(url_for('proposal_list'))

    # --- Contract indicator logic ---
    contract_ind = request.args.get('contract_ind')
    if contract_ind is not None and str(contract_ind).strip().lower() in ('yes', 'true', '1'):
        # Attempt to move the folder from proposals to contracts
        src_path = os.path.join(PROPOSALS_DIR, folder_name)
        dest_path = os.path.join(CONTRACTS_DIR, folder_name)
        try:
            # Check if source exists
            if not os.path.exists(src_path):
                flash(f"Source folder '{src_path}' does not exist.", "error")
            elif os.path.exists(dest_path):
                flash(f"Target folder '{dest_path}' already exists.", "error")
            else:
                shutil.move(src_path, dest_path)
                flash(f"Proposal '{folder_name}' moved to contracts.", "success")
        except Exception as e:
            flash(f"Error moving proposal: {e}", "error")
        return redirect(url_for('proposal_list'))

    # --- Close contract indicator logic ---
    close_ind = request.args.get('close_ind')
    if close_ind is not None and str(close_ind).strip().lower() in ('yes', 'true', '1'):
        # Attempt to move the folder from contracts to completed
        src_path = os.path.join(CONTRACTS_DIR, folder_name)
        dest_path = os.path.join(COMPLETED_DIR, folder_name)
        try:
            if not os.path.exists(src_path):
                flash(f"Source folder '{src_path}' does not exist.", "error")
            elif os.path.exists(dest_path):
                flash(f"Target folder '{dest_path}' already exists.", "error")
            else:
                shutil.move(src_path, dest_path)
                flash(f"Contract '{folder_name}' closed and moved to Completed.", "success")
        except Exception as e:
            flash(f"Error closing contract: {e}", "error")
        return redirect(url_for('proposal_list', status='under'))

    # Determine source root (Open Proposals vs Contracts) by checking where the folder exists
    safe_folder = os.path.basename(folder_name)
    proposals_path = os.path.join(PROPOSALS_DIR, safe_folder)
    contracts_path = os.path.join(CONTRACTS_DIR, safe_folder)

    if os.path.isdir(proposals_path):
        folder_path = proposals_path
    elif os.path.isdir(contracts_path):
        folder_path = contracts_path
    else:
        return f"Folder not found in either PROPOSALS_DIR or CONTRACTS_DIR: {safe_folder}", 404

    # Find the first file in the folder that starts with 'Profit Summary'
    profit_files = [f for f in os.listdir(folder_path) if f.startswith("Profit Summary") and f.endswith((".xlsm", ".xlsx"))]
    if not profit_files:
        return f"No Profit Summary file found in folder: {folder_name}"

    file_path = os.path.join(folder_path, profit_files[0])

    # Read the Excel file into a summary_data 2D list
    summary_data = pd.read_excel(file_path, header=None).values.tolist()

    # Safely read Proposal Note from C40 (row 40, col C)
    try:
        _proposal_note_import = summary_data[39][2]
    except Exception:
        _proposal_note_import = ""

    # Safely read proposal language from C41 (row 41, col C)
    try:
        _proposal_language_import = summary_data[40][2]  # C41
    except Exception:
        _proposal_language_import = ""

    # Extract specific values
    data = {
        "squares": summary_data[2][4],               # E3
        "product": summary_data[2][7],               # H3
        "price_per_sq_10": summary_data[2][12],      # M3
        "total_price_10": summary_data[2][15],       # P3
        "current_roof": summary_data[4][4],          # E5
        "warranty_incl": summary_data[4][7],         # H5
        "price_per_sq_15": summary_data[4][12],      # M5
        "total_price_15": summary_data[4][15],       # P5
        "labor_days": summary_data[6][4],            # E7
        "price_per_sq_20": summary_data[6][12],      # M7
        "total_price_20": summary_data[6][15],       # P7
        "includes_text": _proposal_language_import,
        "proposal_language": _proposal_language_import,
        "submitted_by": summary_data[6][7],          # H7
        "previous_submitted_by": summary_data[6][7], # H7
        "silicone_units_10": summary_data[10][2],    # C11
        "silicone_price": summary_data[10][3],       # D11
        "silicone_total": summary_data[10][4],       # E11
        "gaco_patch_units": summary_data[11][2],     # C12
        "gaco_patch_price": summary_data[11][3],     # D12
        "gaco_patch_total": summary_data[11][4],     # E12
        "bleed_trap_units": summary_data[12][2],     # C13
        "bleed_trap_price": summary_data[12][3],     # D13
        "bleed_trap_total": summary_data[12][4],     # E13
        "sw_1flash_units": summary_data[13][2],      # C14
        "sw_1flash_price": summary_data[13][3],      # D14
        "sw_1flash_total": summary_data[13][4],      # E14
        "sw_bleed_block_units": summary_data[14][2], # C15
        "sw_bleed_block_price": summary_data[14][3], # D15
        "sw_bleed_block_total": summary_data[14][4], # E15
        "drainage_mat_units": summary_data[15][2],   # C16
        "drainage_mat_price": summary_data[15][3],   # D16
        "drainage_mat_total": summary_data[15][4],   # E16
        "foam_units": summary_data[16][2],           # C17
        "foam_price": summary_data[16][3],           # D17
        "foam_total": summary_data[16][4],           # E17
        "rfc_labor_price": summary_data[17][3],      # D18
        "rfc_labor_total": summary_data[17][4],      # E18
        "scarifying_total": summary_data[18][4],     # E19
        "pcs_labor_price": summary_data[19][3],      # D20
        "pcs_labor_total": summary_data[19][4],      # E20
        "travel_total": summary_data[20][4],         # E21
        "misc_costs_total": summary_data[21][4],     # E22
        "warranty_10_total": summary_data[22][4],    # E23
        "office_fee_total": summary_data[23][4],     # E24
        "total_cost": summary_data[25][4],           # E26
        "pcs_profit": summary_data[27][4],           # E28
        "profit_pct": summary_data[28][4],           # E29
        "daily_profit": summary_data[29][4],         # E30
        "profit_share": summary_data[30][4],         # E31
        "commission_amt": summary_data[31][4],       # E32
        "customer_name": summary_data[0][2],         # C1
        "street_address": summary_data[0][7],        # H1
        "city": summary_data[0][13],                 # N1
        "state": summary_data[0][18],                # S1
        "zip_code": summary_data[0][20],             # U1
        "proposal_note": _proposal_note_import,      # C40
    }

    # Ensure required keys exist for the template & triggers (Excel import init only)
    data.setdefault("coverage_10", 0)
    data.setdefault("coverage_15", 0)
    data.setdefault("coverage_20", 0)
    data.setdefault("adjusted_coverage", 0)

    # Initialize previous_* to current for first round-trip after Excel import
    data["previous_roof_type"] = str(data.get("current_roof") or "")
    data["previous_product"] = str(data.get("product") or "")
    try:
        data["previous_squares"] = float(data.get("squares") or 0)
    except Exception:
        data["previous_squares"] = 0.0
    try:
        data["previous_adjusted_coverage"] = float(data.get("adjusted_coverage") or 0)
    except Exception:
        data["previous_adjusted_coverage"] = 0.0

    # Derive Office Fee % and Commission % from Submitted By (Excel import should not overwrite with wrong cell)
    submitted_by_import = str(data.get("submitted_by") or "").strip()
    if submitted_by_import == "David Estes":
        office_fee_pct_import = DAVIDS_OFFICE_FEE_PCT
    else:
        office_fee_pct_import = BASE_OFFICE_FEE_PCT
    data["office_fee_pct"] = office_fee_pct_import

    if submitted_by_import in ("David Estes", "Vern Abbott"):
        data["commission_pct"] = COMMISSION_PCT
    else:
        data["commission_pct"] = 0.0

    # Recompute Office Fee total using Excel-style rounding so totals align on first render
    try:
        tp10_import = float(data.get("total_price_10") or 0)
    except Exception:
        tp10_import = 0.0
    data["office_fee_total"] = excel_round(tp10_import * data["office_fee_pct"], 0)

    # Carry read-only flag through GET round-trips, supporting both new and legacy formats
    read_only_param = request.args.get('read_only')
    if read_only_param is not None:
        readonly = (read_only_param.strip().lower() == 'yes')
    else:
        readonly = (request.args.get('readonly') == '1')
    return render_template(
        "proposal_details.html",
        data=data,
        folder_name=folder_name,
        readonly=readonly,
        is_blank=False,
    )

        
def replace_placeholder_blocks(doc, replacements):
    def replace_text_in_block(paragraph_or_cell):
        full_text = ''.join(run.text for run in paragraph_or_cell.runs)
        for key, val in replacements.items():
            full_text = full_text.replace(key, str(val))
        for run in paragraph_or_cell.runs:
            run.text = ''
        if paragraph_or_cell.runs:
            paragraph_or_cell.runs[0].text = full_text

    for para in doc.paragraphs:
        if any(key in para.text for key in replacements):
            replace_text_in_block(para)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if any(key in para.text for key in replacements):
                        replace_text_in_block(para)


if __name__ == "__main__":
    app.run(debug=True)
