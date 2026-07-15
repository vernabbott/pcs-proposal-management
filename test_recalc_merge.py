import ast
import datetime
import math
import os
import pathlib
import shutil
import tempfile
import threading
import time
import unittest
from urllib.parse import quote, urlsplit, urlunsplit


class CommissionCalculationTests(unittest.TestCase):
    def test_commission_excludes_office_fee_for_all_terms(self):
        import pcs_proposal_web as app

        result = app.calculation_routine(
            squares=100,
            product="Gaco",
            roof_type="TPO/EPDM",
            labor_days=0,
            warranty_incl="No",
            include_travel="No",
            price_per_sq_10=0,
            commission_pct=0,
            submitted_by="Vern",
            previous_submitted_by="Vern",
            office_fee_pct=0.05,
            adjusted_coverage=0,
            silicone_units_10=0,
            silicone_price=0,
            gaco_patch_units=0,
            gaco_patch_price=0,
            sw_1flash_units=0,
            sw_1flash_price=0,
            bleed_trap_units=0,
            bleed_trap_price=0,
            sw_bleed_block_units=0,
            sw_bleed_block_price=0,
            drainage_mat_units=0,
            drainage_mat_price=0,
            foam_units=0,
            foam_price=0,
            rfc_labor_price=0,
            pcs_labor_price=0,
            scarifying_total=0,
            travel_total=0,
            repair_costs_total=0,
            previous_squares=100,
            previous_roof_type="TPO/EPDM",
            previous_product="Gaco",
            previous_adjusted_coverage=0,
            previous_silicone_units_10=0,
            proposal_note="",
            pcs_or_roofer_ind="PCS Direct",
            previous_pcs_or_roofer_ind="PCS Direct",
            previous_include_travel="No",
            previous_calc_travel_total=0,
        )

        self.assertEqual(result["office_fee_total"], 1750)
        self.assertEqual(result["total_price_10"], 36750)
        self.assertEqual(result["commission_amt"], 3500)
        self.assertEqual(result["commission_amt_15"], 3900)
        self.assertEqual(result["commission_amt_20"], 4300)


class RockFoamCalculationTests(unittest.TestCase):
    def _calculate(self, **overrides):
        import pcs_proposal_web as app

        params = dict(
            squares=100,
            product="Gaco",
            roof_type="Rock/Foam/Coat",
            labor_days=0,
            warranty_incl="No",
            include_travel="No",
            price_per_sq_10=0,
            commission_pct=0,
            submitted_by="Vern",
            previous_submitted_by="Vern",
            office_fee_pct=0,
            adjusted_coverage=0,
            silicone_units_10=0,
            silicone_price=0,
            gaco_patch_units=0,
            gaco_patch_price=0,
            sw_1flash_units=0,
            sw_1flash_price=0,
            bleed_trap_units=0,
            bleed_trap_price=0,
            sw_bleed_block_units=0,
            sw_bleed_block_price=0,
            drainage_mat_units=0,
            drainage_mat_price=0,
            foam_units=0,
            foam_price=0,
            rfc_labor_price=0,
            pcs_labor_price=0,
            scarifying_total=0,
            travel_total=0,
            repair_costs_total=0,
            previous_squares=100,
            previous_roof_type="Rock/Foam/Coat",
            previous_product="Gaco",
            previous_adjusted_coverage=0,
            previous_silicone_units_10=0,
            proposal_note="",
            pcs_or_roofer_ind="PCS Direct",
            previous_pcs_or_roofer_ind="PCS Direct",
            previous_include_travel="No",
            previous_calc_travel_total=0,
        )
        params.update(overrides)
        return app.calculation_routine(**params)

    def test_rock_foam_zero_formula_cache_recalculates_foam_price(self):
        result = self._calculate()

        self.assertEqual(result["foam_units"], 4)
        self.assertEqual(result["foam_price"], 2600)
        self.assertEqual(result["foam_total"], 10400)

    def test_rock_foam_zero_formula_cache_recalculates_removal_price(self):
        result = self._calculate()

        self.assertEqual(result["rfc_labor_price"], 250)
        self.assertEqual(result["rfc_labor_total"], 25000)


class FullDetailPreviewTests(unittest.TestCase):
    def test_blank_proposal_can_preview_full_detail_from_posted_values(self):
        import pcs_proposal_web as app

        payload = {
            "action": "full_detail_preview",
            "customer_name": "Preview Customer",
            "street_address": "123 Preview St",
            "city": "Denver",
            "state": "CO",
            "zip_code": "80202",
            "flat_roof_squares": "100",
            "wall_squares": "0",
            "squares": "100",
            "current_roof": "TPO/EPDM",
            "product": "Gaco",
            "submitted_by": "Vern",
            "pcs_or_roofer_ind": "PCS Direct",
            "warranty_incl": "No",
            "include_travel": "No",
        }

        with app.app.test_client() as client:
            response = client.post("/update-proposal/NEW", data=payload)

        html = response.get_data(as_text=True)
        self.assertEqual(response.status_code, 200)
        self.assertIn("Full Detail - Preview Customer", html)
        self.assertIn("$36,750", html)
        self.assertIn("$3,500", html)


def load_merge_display_fallbacks(read_profit_summary_for_display):
    source_path = pathlib.Path(__file__).with_name("pcs_proposal_web.py")
    source = source_path.read_text(encoding="utf-8")
    module = ast.parse(source, filename=str(source_path))

    fn_node = next(
        node for node in module.body
        if isinstance(node, ast.FunctionDef) and node.name == "merge_display_fallbacks"
    )

    isolated_module = ast.Module(body=[fn_node], type_ignores=[])
    ast.fix_missing_locations(isolated_module)

    namespace = {
        "math": math,
        "os": os,
        "read_profit_summary_for_display": read_profit_summary_for_display,
    }
    exec(compile(isolated_module, str(source_path), "exec"), namespace)
    return namespace["merge_display_fallbacks"]


def load_copy_helpers():
    source_path = pathlib.Path(__file__).with_name("pcs_proposal_web.py")
    source = source_path.read_text(encoding="utf-8")
    module = ast.parse(source, filename=str(source_path))

    target_names = {
        "_is_pdf_ready_for_copy",
        "_is_generated_proposal_artifact",
        "_dated_archive_folder",
        "_archive_existing_artifacts",
        "_sync_directory_contents",
        "_wait_for_generated_pdfs",
        "copy_proposal_to_submitter_destination",
    }
    selected_nodes = [
        node for node in module.body
        if isinstance(node, ast.FunctionDef) and node.name in target_names
    ]

    isolated_module = ast.Module(body=selected_nodes, type_ignores=[])
    ast.fix_missing_locations(isolated_module)

    namespace = {
        "os": os,
        "shutil": shutil,
        "time": time,
        "datetime": datetime,
        "_safe_debug": lambda message: None,
        "get_copy_destination_for_submitter": lambda submitted_by: None,
    }
    exec(compile(isolated_module, str(source_path), "exec"), namespace)
    return namespace


def load_email_helpers(plain_template, html_template):
    source_path = pathlib.Path(__file__).with_name("pcs_proposal_web.py")
    source = source_path.read_text(encoding="utf-8")
    module = ast.parse(source, filename=str(source_path))

    target_names = {
        "_format_currency",
        "_format_square_count",
        "_insert_proposal_summary_extras_plain",
        "_insert_proposal_summary_extras_html",
        "_format_folder_link_html",
        "_build_proposal_summary_email_bodies",
        "build_proposal_email_subject",
        "build_proposal_summary_email_html",
        "build_proposal_summary_email_text",
    }
    selected_nodes = [
        node for node in module.body
        if isinstance(node, ast.FunctionDef) and node.name in target_names
    ]

    isolated_module = ast.Module(body=selected_nodes, type_ignores=[])
    ast.fix_missing_locations(isolated_module)

    namespace = {
        "html": __import__("html"),
        "_load_proposal_summary_email_template": lambda: (plain_template, html_template),
    }
    exec(compile(isolated_module, str(source_path), "exec"), namespace)
    return namespace


def load_folder_link_helper():
    source_path = pathlib.Path(__file__).with_name("pcs_proposal_web.py")
    source = source_path.read_text(encoding="utf-8")
    module = ast.parse(source, filename=str(source_path))

    target_names = {
        "_join_url_path",
        "build_proposal_folder_link",
    }
    selected_nodes = [
        node for node in module.body
        if isinstance(node, ast.FunctionDef) and node.name in target_names
    ]

    isolated_module = ast.Module(body=selected_nodes, type_ignores=[])
    ast.fix_missing_locations(isolated_module)

    destination_root = "/submitter/root"
    namespace = {
        "os": os,
        "pathlib": pathlib,
        "quote": quote,
        "urlsplit": urlsplit,
        "urlunsplit": urlunsplit,
        "get_copy_destination_for_submitter": lambda submitted_by: destination_root,
        "get_copy_destination_web_url_for_submitter": lambda submitted_by: "https://example.test/open",
    }
    exec(compile(isolated_module, str(source_path), "exec"), namespace)
    namespace["destination_root"] = destination_root
    return namespace


class MergeDisplayFallbacksTests(unittest.TestCase):
    def test_post_refresh_keeps_fresh_recalculated_values(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            incoming = {
                "office_fee_total": 700,
                "commission_amt": 900,
                "pcs_profit": 5000,
                "daily_profit": 625,
                "profit_share": 550,
                "price_per_sq_15": 395,
            }
            persisted = {
                "office_fee_total": 400,
                "commission_amt": 0,
                "pcs_profit": 4400,
                "daily_profit": 550,
                "profit_share": 490,
                "price_per_sq_15": 365,
            }

            merge_display_fallbacks = load_merge_display_fallbacks(
                lambda folder_path: persisted
            )
            merged = merge_display_fallbacks(
                dict(incoming),
                temp_dir,
                "Existing Proposal",
                prefer_saved_derived=False,
            )

        self.assertEqual(merged["office_fee_total"], 700)
        self.assertEqual(merged["commission_amt"], 900)
        self.assertEqual(merged["pcs_profit"], 5000)
        self.assertEqual(merged["daily_profit"], 625)
        self.assertEqual(merged["profit_share"], 550)
        self.assertEqual(merged["price_per_sq_15"], 395)

    def test_get_refresh_can_still_use_saved_derived_values(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            incoming = {
                "office_fee_total": 700,
                "commission_amt": 900,
                "pcs_profit": 5000,
            }
            persisted = {
                "office_fee_total": 400,
                "commission_amt": 0,
                "pcs_profit": 4400,
            }

            merge_display_fallbacks = load_merge_display_fallbacks(
                lambda folder_path: persisted
            )
            merged = merge_display_fallbacks(
                dict(incoming),
                temp_dir,
                "Existing Proposal",
            )

        self.assertEqual(merged["office_fee_total"], 400)
        self.assertEqual(merged["commission_amt"], 0)
        self.assertEqual(merged["pcs_profit"], 4400)


class SubmitterDestinationCopyTests(unittest.TestCase):
    def test_archive_leaves_unsupported_files_in_place(self):
        helpers = load_copy_helpers()
        archive_existing_artifacts = helpers["_archive_existing_artifacts"]

        with tempfile.TemporaryDirectory() as proposal_dir:
            managed_docx = pathlib.Path(proposal_dir, "Gaco S42 Proposal - Main St.docx")
            managed_xlsx = pathlib.Path(proposal_dir, "Profit Summary - Main St.xlsx")
            managed_pdf = pathlib.Path(proposal_dir, "Gaco S42 Proposal - Main St.pdf")
            unrelated_docx = pathlib.Path(proposal_dir, "Customer Notes.docx")
            unrelated_xlsx = pathlib.Path(proposal_dir, "Roof Measurements.xlsx")
            unrelated_pdf = pathlib.Path(proposal_dir, "Inspection Photos.pdf")
            unsupported_doc = pathlib.Path(proposal_dir, "Legacy Proposal.doc")
            unsupported_txt = pathlib.Path(proposal_dir, "notes.txt")

            for path in (
                managed_docx,
                managed_xlsx,
                managed_pdf,
                unrelated_docx,
                unrelated_xlsx,
                unrelated_pdf,
                unsupported_doc,
                unsupported_txt,
            ):
                path.write_text(path.name, encoding="utf-8")

            archive_existing_artifacts(proposal_dir)

            archive_root = pathlib.Path(proposal_dir, "Archive")
            archive_dirs = [path for path in archive_root.iterdir() if path.is_dir()]
            self.assertEqual(len(archive_dirs), 1)
            archive_dir = archive_dirs[0]
            for path in (
                managed_docx,
                managed_xlsx,
                managed_pdf,
                unrelated_docx,
                unrelated_xlsx,
                unrelated_pdf,
                unsupported_doc,
                unsupported_txt,
            ):
                self.assertTrue((archive_dir / path.name).exists())
            self.assertFalse(managed_docx.exists())
            self.assertFalse(managed_xlsx.exists())
            self.assertFalse(managed_pdf.exists())
            self.assertTrue(unrelated_docx.exists())
            self.assertTrue(unrelated_xlsx.exists())
            self.assertTrue(unrelated_pdf.exists())
            self.assertTrue(unsupported_doc.exists())
            self.assertTrue(unsupported_txt.exists())

    def test_destination_sync_preserves_unsupported_files(self):
        helpers = load_copy_helpers()
        sync_directory_contents = helpers["_sync_directory_contents"]

        with tempfile.TemporaryDirectory() as source_dir, tempfile.TemporaryDirectory() as dest_dir:
            pathlib.Path(source_dir, "Gaco S42 Proposal - Main St.docx").write_text("docx", encoding="utf-8")
            pathlib.Path(source_dir, "Inspection Photos.pdf").write_text("source-pdf", encoding="utf-8")
            pathlib.Path(source_dir, "source-notes.txt").write_text("source", encoding="utf-8")
            pathlib.Path(dest_dir, "Uniflex Proposal - Old St.pdf").write_text("old", encoding="utf-8")
            pathlib.Path(dest_dir, "Inspection Photos.pdf").write_text("dest-pdf", encoding="utf-8")
            pathlib.Path(dest_dir, "keep-notes.txt").write_text("keep", encoding="utf-8")

            sync_directory_contents(source_dir, dest_dir)

            self.assertTrue(pathlib.Path(dest_dir, "Gaco S42 Proposal - Main St.docx").exists())
            self.assertFalse(pathlib.Path(dest_dir, "Uniflex Proposal - Old St.pdf").exists())
            self.assertEqual(
                pathlib.Path(dest_dir, "Inspection Photos.pdf").read_text(encoding="utf-8"),
                "dest-pdf",
            )
            self.assertFalse(pathlib.Path(dest_dir, "source-notes.txt").exists())
            self.assertTrue(pathlib.Path(dest_dir, "keep-notes.txt").exists())

    def test_copy_waits_for_pdf_before_syncing_destination(self):
        helpers = load_copy_helpers()
        copy_proposal_to_submitter_destination = helpers["copy_proposal_to_submitter_destination"]

        with tempfile.TemporaryDirectory() as source_dir, tempfile.TemporaryDirectory() as dest_root:
            folder_name = "Example Proposal"
            docx_path = os.path.join(source_dir, "Gaco S42 Proposal - Main St.docx")
            xlsx_path = os.path.join(source_dir, "Profit Summary - Main St.xlsx")
            pdf_path = os.path.join(source_dir, "Gaco S42 Proposal - Main St.pdf")

            pathlib.Path(docx_path).write_text("docx", encoding="utf-8")
            pathlib.Path(xlsx_path).write_text("xlsx", encoding="utf-8")

            def create_pdf_later():
                time.sleep(0.2)
                pathlib.Path(pdf_path).write_bytes(b"%PDF-1.4\nbody\n%%EOF\n")

            writer = threading.Thread(target=create_pdf_later, daemon=True)
            writer.start()

            helpers["get_copy_destination_for_submitter"] = lambda submitted_by: dest_root
            copy_proposal_to_submitter_destination.__globals__["get_copy_destination_for_submitter"] = (
                helpers["get_copy_destination_for_submitter"]
            )

            copy_proposal_to_submitter_destination(
                source_dir,
                folder_name,
                "Vern",
                wait_for_pdfs=True,
            )

            dest_folder = os.path.join(dest_root, folder_name)
            self.assertTrue(os.path.exists(os.path.join(dest_folder, "Gaco S42 Proposal - Main St.docx")))
            self.assertTrue(os.path.exists(os.path.join(dest_folder, "Profit Summary - Main St.xlsx")))
            self.assertTrue(os.path.exists(os.path.join(dest_folder, "Gaco S42 Proposal - Main St.pdf")))

    def test_copy_waits_for_pdf_to_finish_writing(self):
        helpers = load_copy_helpers()
        copy_proposal_to_submitter_destination = helpers["copy_proposal_to_submitter_destination"]

        with tempfile.TemporaryDirectory() as source_dir, tempfile.TemporaryDirectory() as dest_root:
            folder_name = "Example Proposal"
            docx_path = os.path.join(source_dir, "Gaco S42 Proposal - Main St.docx")
            pdf_path = os.path.join(source_dir, "Gaco S42 Proposal - Main St.pdf")

            pathlib.Path(docx_path).write_text("docx", encoding="utf-8")

            def write_pdf_in_stages():
                with open(pdf_path, "wb") as handle:
                    handle.write(b"%PDF-")
                    handle.flush()
                    os.fsync(handle.fileno())
                    time.sleep(0.35)
                    handle.write(b"1.4\nbody\n%%EOF\n")
                    handle.flush()
                    os.fsync(handle.fileno())

            writer = threading.Thread(target=write_pdf_in_stages, daemon=True)
            writer.start()

            helpers["get_copy_destination_for_submitter"] = lambda submitted_by: dest_root
            copy_proposal_to_submitter_destination.__globals__["get_copy_destination_for_submitter"] = (
                helpers["get_copy_destination_for_submitter"]
            )

            copy_proposal_to_submitter_destination(
                source_dir,
                folder_name,
                "Vern",
                wait_for_pdfs=True,
            )

            dest_pdf_path = os.path.join(dest_root, folder_name, "Gaco S42 Proposal - Main St.pdf")
            self.assertTrue(os.path.exists(dest_pdf_path))
            self.assertEqual(
                pathlib.Path(dest_pdf_path).read_bytes(),
                pathlib.Path(pdf_path).read_bytes(),
            )
            self.assertGreater(len(pathlib.Path(dest_pdf_path).read_bytes()), 5)


class ProposalSummaryEmailLinkTests(unittest.TestCase):
    def test_submitter_destination_path_builds_email_link(self):
        helpers = load_folder_link_helper()
        build_proposal_folder_link = helpers["build_proposal_folder_link"]
        destination_root = helpers["destination_root"]
        folder_name = "Customer - 123 Main St"
        folder_path = os.path.join(destination_root, folder_name)

        link = build_proposal_folder_link(
            folder_path,
            submitted_by="David",
            folder_name=folder_name,
        )

        self.assertEqual(
            link,
            "https://example.test/open/Customer%20-%20123%20Main%20St",
        )

    def test_test_site_path_is_replaced_with_submitter_destination_link(self):
        helpers = load_folder_link_helper()
        build_proposal_folder_link = helpers["build_proposal_folder_link"]
        folder_name = "Customer - 123 Main St"

        link = build_proposal_folder_link(
            f"/test/site/{folder_name}",
            submitted_by="David",
            folder_name=folder_name,
        )

        self.assertEqual(
            link,
            "https://example.test/open/Customer%20-%20123%20Main%20St",
        )

    def test_html_email_links_folder_name(self):
        helpers = load_email_helpers(
            plain_template="Proposal for\r\nFolderName\r\nDaily 10-year profit - 10YrProfit",
            html_template="<html><body><div>Proposal for</div><div>FolderName</div><ul><li>Daily 10-year profit - 10YrProfit</li></ul></body></html>",
        )

        html_body = helpers["build_proposal_summary_email_html"](
            customer_name="Example Customer",
            street_address="123 Main St",
            folder_link="file:///tmp/Example%20Customer%20-%20123%20Main%20St",
            total_squares=10,
            flat_roof_squares=8,
            wall_squares=2,
            roof_type="Metal",
            daily_profit=1234,
            proposal_note="",
            proposal_language="",
        )

        self.assertIn(
            '<a href="file:///tmp/Example%20Customer%20-%20123%20Main%20St">Example Customer 123 Main St</a>',
            html_body,
        )
        self.assertNotIn(">FolderName<", html_body)

    def test_plain_text_email_keeps_folder_name_without_html(self):
        helpers = load_email_helpers(
            plain_template="Proposal for\r\nFolderName\r\nDaily 10-year profit - 10YrProfit",
            html_template="<html><body>FolderName<ul></ul></body></html>",
        )

        plain_body = helpers["build_proposal_summary_email_text"](
            customer_name="Example Customer",
            street_address="123 Main St",
            folder_link="file:///tmp/Example",
            total_squares=10,
            flat_roof_squares=8,
            wall_squares=2,
            roof_type="Metal",
            daily_profit=1234,
            proposal_note="",
            proposal_language="",
        )

        self.assertIn("Example Customer 123 Main St", plain_body)
        self.assertNotIn("<a href=", plain_body)


if __name__ == "__main__":
    unittest.main()
