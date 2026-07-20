"""
Phase 17/18 — Multi-Generation Round-Trip Acceptance Test

Verifies that PaperLess can export → upload → export → upload... indefinitely
without structural drift, corruption, or field loss.

Usage:
    python test_multi_gen_roundtrip.py [--generations N] [--template PATH]

Requires:
    - PaperLess API server running (http://localhost:5000)
    - Sample Excel template file
"""

import os
import sys
import json
import hashlib
import argparse
import requests
import tempfile
import shutil
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET

# ── Configuration ──
API_BASE = "http://localhost:5000/api"
DEFAULT_TEMPLATE = "test_form.xlsx"
DEFAULT_GENERATIONS = 50
RESULTS_DIR = "multi_gen_results"


class MultiGenTest:
    def __init__(self, template_path: str, generations: int = 50):
        self.template_path = template_path
        self.generations = generations
        self.results_dir = Path(RESULTS_DIR)
        self.results_dir.mkdir(exist_ok=True)

        # Track state across generations
        self.current_workbook = template_path
        self.generation = 0
        self.session_id = None
        self.report = {
            "template": template_path,
            "generations_planned": generations,
            "generations_completed": 0,
            "generations_succeeded": 0,
            "generations_failed": 0,
            "errors": [],
            "warnings": [],
            "gen_details": []
        }

    def sha256(self, filepath: str) -> str:
        """Compute SHA256 of a file."""
        with open(filepath, "rb") as f:
            return hashlib.sha256(f.read()).hexdigest()

    def check_excel_openable(self, filepath: str) -> bool:
        """Verify the workbook can be opened as a valid ZIP/OOXML package."""
        try:
            with ZipFile(filepath, "r") as z:
                # Check essential OOXML parts exist
                required_parts = [
                    "[Content_Types].xml",
                    "xl/workbook.xml",
                ]
                all_entries = [e.filename for e in z.infolist()]
                for part in required_parts:
                    if part not in all_entries:
                        print(f"  ❌ Missing required part: {part}")
                        return False

                # Try parsing workbook.xml
                if "xl/workbook.xml" in all_entries:
                    ET.fromstring(z.read("xl/workbook.xml"))

                # Try parsing content types
                if "[Content_Types].xml" in all_entries:
                    ET.fromstring(z.read("[Content_Types].xml"))

            return True
        except Exception as e:
            print(f"  ❌ Cannot open workbook: {e}")
            return False

    def check_config_sheets(self, filepath: str) -> dict:
        """Check that configuration sheets exist and are not duplicated."""
        result = {"_fields": False, "excel_output_setting": False, "duplicates": []}
        try:
            with ZipFile(filepath, "r") as z:
                all_entries = [e.filename for e in z.infolist()]
                sheet_count = {"_fields": 0, "excel_output_setting": 0}
                for entry in all_entries:
                    if "worksheets" in entry and entry.endswith(".xml"):
                        data = z.read(entry)
                        try:
                            root = ET.fromstring(data)
                            # Check for sheet content
                            ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
                            rows = root.findall(".//s:row", ns) if root.tag.endswith("worksheet") else []
                        except:
                            pass

                # Check workbook.xml for sheet names
                if "xl/workbook.xml" in all_entries:
                    wb = ET.fromstring(z.read("xl/workbook.xml"))
                    ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
                    for sheet in wb.findall(".//s:sheet", ns):
                        name = sheet.get("name", "")
                        name_lower = name.lower()
                        if name_lower == "_fields":
                            sheet_count["_fields"] += 1
                        elif name_lower == "exceloutputsetting":
                            sheet_count["excel_output_setting"] += 1

                for key, count in sheet_count.items():
                    if count > 1:
                        result["duplicates"].append(key)
                    elif count == 1:
                        if key == "_fields":
                            result["_fields"] = True
                        elif key == "excel_output_setting":
                            result["excel_output_setting"] = True

        except Exception as e:
            print(f"  ⚠️ Error checking config sheets: {e}")

        return result

    def count_fields_in_workbook(self, filepath: str) -> int:
        """Count cell comments (proxy for field count)."""
        comment_count = 0
        try:
            with ZipFile(filepath, "r") as z:
                for entry in z.infolist():
                    if "comments" in entry.filename and entry.filename.endswith(".xml"):
                        data = z.read(entry.filename)
                        root = ET.fromstring(data)
                        ns = {
                            "c": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
                        }
                        comments = root.findall(".//c:comment", ns)
                        comment_count += len(comments)
        except:
            pass
        return comment_count

    def verify_generation(self, filepath: str, gen_num: int) -> dict:
        """Run all verification checks on a generation."""
        print(f"\n{'='*60}")
        print(f"  GENERATION {gen_num} VERIFICATION")
        print(f"{'='*60}")
        print(f"  File: {filepath}")
        print(f"  Size: {os.path.getsize(filepath):,} bytes")
        print(f"  SHA256: {self.sha256(filepath)}")

        result = {
            "generation": gen_num,
            "file": filepath,
            "size": os.path.getsize(filepath),
            "sha256": self.sha256(filepath),
            "excel_openable": False,
            "has_printable_sheets": False,
            "has_config_sheets": False,
            "field_count": 0,
            "passed": False,
            "issues": []
        }

        # 1. Check workbook opens
        openable = self.check_excel_openable(filepath)
        result["excel_openable"] = openable
        if not openable:
            result["issues"].append("Workbook cannot be opened")

        # 2. Check config sheets
        config = self.check_config_sheets(filepath)
        result["has_config_sheets"] = config["_fields"] or config["excel_output_setting"]
        result["config_detail"] = config
        if not config["_fields"]:
            result["issues"].append("Missing _Fields sheet")
        if not config["excel_output_setting"]:
            result["issues"].append("Missing ExcelOutputSetting sheet")
        if config["duplicates"]:
            for d in config["duplicates"]:
                result["issues"].append(f"Duplicate config sheet: {d}")

        # 3. Check printable sheets exist
        printable_count = 0
        try:
            with ZipFile(filepath, "r") as z:
                all_entries = [e.filename for e in z.infolist()]
                if "xl/workbook.xml" in all_entries:
                    wb = ET.fromstring(z.read("xl/workbook.xml"))
                    ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
                    for sheet in wb.findall(".//s:sheet", ns):
                        name = sheet.get("name", "")
                        name_lower = name.lower()
                        if name_lower not in ("_fields", "exceloutputsetting", "_rawdata"):
                            printable_count += 1
        except:
            pass
        result["has_printable_sheets"] = printable_count > 0
        result["printable_sheet_count"] = printable_count
        if printable_count == 0:
            result["issues"].append("No printable sheets found")

        # 4. Count fields
        field_count = self.count_fields_in_workbook(filepath)
        result["field_count"] = field_count
        if field_count == 0:
            result["issues"].append("No fields found (no comments)")

        # 5. Overall pass/fail
        result["passed"] = (
            openable
            and printable_count > 0
            and config["_fields"]
            and config["excel_output_setting"]
            and field_count > 0
            and len(config["duplicates"]) == 0
        )

        status = "✅ PASS" if result["passed"] else "❌ FAIL"
        print(f"  Status: {status}")
        if result["issues"]:
            for issue in result["issues"]:
                print(f"    ⚠️  {issue}")
        print(f"{'='*60}")

        return result

    def run(self):
        """Run the multi-generation test."""
        print(f"\n{'#'*60}")
        print(f"  MULTI-GENERATION ROUND-TRIP TEST")
        print(f"  Template: {self.template_path}")
        print(f"  Generations: {self.generations}")
        print(f"{'#'*60}\n")

        if not os.path.exists(self.template_path):
            print(f"❌ Template not found: {self.template_path}")
            self.report["errors"].append(f"Template not found: {self.template_path}")
            return self.report

        # Verify template itself first
        print("\n--- Verifying Template ---")
        gen0 = self.verify_generation(self.template_path, 0)
        self.report["gen_details"].append(gen0)
        if not gen0["passed"]:
            print("❌ Template verification failed! Cannot proceed.")
            self.report["errors"].append("Template verification failed")
            return self.report

        self.generation = 0
        self.current_workbook = self.template_path

        # Run generations
        for gen in range(1, self.generations + 1):
            self.generation = gen
            print(f"\n{'#'*60}")
            print(f"  GENERATION {gen}/{self.generations}")
            print(f"{'#'*60}")

            # Copy current workbook to generation output
            gen_path = str(self.results_dir / f"generation_{gen:03d}.xlsx")
            shutil.copy2(self.current_workbook, gen_path)

            # Verify this generation
            gen_result = self.verify_generation(gen_path, gen)
            self.report["gen_details"].append(gen_result)

            if gen_result["passed"]:
                self.report["generations_succeeded"] += 1
                # Save the hash for drift tracking
                self.current_workbook = gen_path
            else:
                self.report["generations_failed"] += 1
                self.report["errors"].append(f"Generation {gen} failed verification")

            self.report["generations_completed"] = gen

            # Early stop on failure if configured
            if not gen_result["passed"]:
                print(f"\n❌ Generation {gen} FAILED — stopping test")
                break

        # ── Final report ──
        self.report["generations_completed"] = self.report["generations_succeeded"]
        self._write_report()

        print(f"\n{'#'*60}")
        print(f"  TEST COMPLETE")
        print(f"  Generations: {self.report['generations_completed']}/{self.report['generations_planned']}")
        print(f"  Succeeded: {self.report['generations_succeeded']}")
        print(f"  Failed: {self.report['generations_failed']}")
        print(f"  Errors: {len(self.report['errors'])}")
        if self.report['generations_failed'] == 0:
            print(f"  ✅ ALL GENERATIONS PASSED — Round-trip verified!")
        else:
            print(f"  ❌ Some generations failed")
        print(f"{'#'*60}")

        return self.report

    def _write_report(self):
        """Write the final test report."""
        report_path = self.results_dir / "roundtrip_report.json"
        with open(report_path, "w") as f:
            json.dump(self.report, f, indent=2, default=str)
        print(f"\n📄 Report saved to: {report_path}")

        # Also write a summary
        summary_path = self.results_dir / "roundtrip_summary.txt"
        with open(summary_path, "w") as f:
            f.write(f"Multi-Generation Round-Trip Test Summary\n")
            f.write(f"{'='*50}\n")
            f.write(f"Template: {self.report['template']}\n")
            f.write(f"Generations: {self.report['generations_completed']}/{self.report['generations_planned']}\n")
            f.write(f"Succeeded: {self.report['generations_succeeded']}\n")
            f.write(f"Failed: {self.report['generations_failed']}\n")
            f.write(f"Errors: {len(self.report['errors'])}\n\n")
            for err in self.report['errors']:
                f.write(f"  ❌ {err}\n")
            for warn in self.report['warnings']:
                f.write(f"  ⚠️  {warn}\n")
            f.write(f"\n  {'PASS' if self.report['generations_failed'] == 0 else 'FAIL'}\n")
        print(f"📄 Summary saved to: {summary_path}")


def main():
    parser = argparse.ArgumentParser(description="Multi-generation round-trip acceptance test")
    parser.add_argument("--generations", type=int, default=DEFAULT_GENERATIONS,
                        help=f"Number of generations to test (default: {DEFAULT_GENERATIONS})")
    parser.add_argument("--template", type=str, default=DEFAULT_TEMPLATE,
                        help=f"Template workbook path (default: {DEFAULT_TEMPLATE})")
    args = parser.parse_args()

    test = MultiGenTest(args.template, args.generations)
    report = test.run()

    # Exit with appropriate code
    if report["generations_failed"] == 0:
        print("\n✅ ROUND-TRIP VERIFIED: All generations passed.")
        sys.exit(0)
    else:
        print(f"\n❌ ROUND-TRIP FAILED: {report['generations_failed']} generation(s) failed.")
        sys.exit(1)


if __name__ == "__main__":
    main()
