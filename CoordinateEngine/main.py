import argparse, json, sys, os, traceback, zipfile, time, hashlib
from xml.etree import ElementTree as ET

# Ensure CoordinateEngine is importable as a package when running as a script.
# When python CoordinateEngine/main.py is executed, Python adds the engine
# directory to sys.path but NOT its parent, making "import CoordinateEngine.*"
# fail. Inserting the parent dir resolves this without requiring pip install.
_ce_parent = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ce_parent not in sys.path:
    sys.path.insert(0, _ce_parent)

import CoordinateEngine.workbook_reader as workbook_reader
import CoordinateEngine.pdf_reader as pdf_reader
import CoordinateEngine.image_reader as image_reader
import CoordinateEngine.field_detector as field_detector
import CoordinateEngine.geometry_engine as geometry_engine
import CoordinateEngine.runtime_generator as runtime_generator
import CoordinateEngine.debug_overlay as debug_overlay
import CoordinateEngine.debug_visualizer as debug_visualizer
import CoordinateEngine.validator as validator
import CoordinateEngine.cross_validator as cross_validator
import CoordinateEngine.geometry_tracer as geometry_tracer
import CoordinateEngine.confidence_scorer as confidence_scorer
import CoordinateEngine.error_diagnostics as error_diagnostics
import CoordinateEngine.workbook_analyzer as workbook_analyzer
from CoordinateEngine.utils.file_utils import ensure_dir


def parse_args(argv=None):
    p = argparse.ArgumentParser(description="PaperLess Coordinate Engine")
    p.add_argument("--excel", required=True, help="Workbook.xlsx path")
    p.add_argument("--pdf", required=True, help="Background.pdf path")
    p.add_argument("--png", required=True, help="Background.png path")
    p.add_argument("--output", required=True, help="Output runtime.json path")
    p.add_argument("--dpi", type=int, default=300, help="DPI (default: 300)")
    p.add_argument("--debug-dir", default=None, help="Debug output directory")
    p.add_argument(
        "--validate",
        action="store_true",
        help="Run full cross-validation (slower, requires PDF+PNG)",
    )
    p.add_argument(
        "--trace",
        action="store_true",
        help="Generate geometry trace report",
    )
    p.add_argument(
        "--diagnose",
        action="store_true",
        help="Run error diagnostics",
    )
    p.add_argument(
        "--analyze",
        action="store_true",
        help="Generate workbook_analysis.md",
    )
    return p.parse_args(argv)


def main():
    args = parse_args()
    ensure_dir(os.path.dirname(args.output) or ".")
    timings = {}
    t0 = time.time()

    try:
        # Stage 1: Workbook Parsing (with verbose trace)
        t1 = time.time()
        verbose_mode = args.trace or args.diagnose
        info = workbook_reader.read(args.excel, verbose=verbose_mode)
        timings["workbook_parsing"] = time.time() - t1
        print(f"Workbook: {info.name}")
        print(f"Page: {info.page_width_pt:.0f}x{info.page_height_pt:.0f}pt")
        print(f"Margins: L={info.margin_left_pt:.2f} R={info.margin_right_pt:.2f} T={info.margin_top_pt:.2f} B={info.margin_bottom_pt:.2f}")
        if info.print_area_addr:
            print(f"PrintArea: {info.print_area_addr}")
        print(f"Cells: {len(getattr(info, 'cell_values', {}))}")

        # Stage 2: PDF Content Analysis
        t1 = time.time()
        content = pdf_reader.measure_content_bounds(args.pdf)
        timings["pdf_analysis"] = time.time() - t1
        print(f"PDF content: x0={content.x0:.2f} y0={content.y0:.2f} W={content.width_pt:.2f} H={content.height_pt:.2f}pt")

        # Stage 3: Layout Computation
        t1 = time.time()
        layout = geometry_engine.build_layout(info, content, args.dpi)
        timings["geometry_computation"] = time.time() - t1
        print(f"Centered: H={layout.centered_h} V={layout.centered_v}")
        print(f"Origin: ({layout.origin_x_pt:.2f}, {layout.origin_y_pt:.2f})pt")
        print(f"Scale: {layout.scale_w:.6f} x {layout.scale_h:.6f}")
        print(f"Effective dims: {layout.eff_w_pt:.2f}x{layout.eff_h_pt:.2f}pt (PA: {layout.paw_pt:.2f}x{layout.pah_pt:.2f}pt)")

        # Stage 4: Field Detection (now using WorkbookInfo)
        t1 = time.time()
        fields = field_detector.detect(info, verbose=verbose_mode)
        timings["field_detection"] = time.time() - t1
        print(f"Fields detected: {len(fields)}")

        # Stage 5: Geometry Computation
        t1 = time.time()
        geometries = []
        for f in fields:
            g = geometry_engine.compute_field_geometry(f, layout, info, args.dpi)
            geometries.append(g)
            print(f"  {f.addr:8s}  page=({g.page_left_pt:8.2f},{g.page_top_pt:8.2f})  "
                  f"w{g.page_width_pt:7.2f}h{g.page_height_pt:7.2f}pt  "
                  f"px=({g.left_px:7.1f},{g.top_px:7.1f})  "
                  f"ratio=({g.left_ratio:.6f},{g.top_ratio:.6f})")
        timings["per_field_geometry"] = time.time() - t1

        # Stage 6: Validation
        ok, msg = validator.validate_fields(geometries)
        if not ok:
            print(f"Validation FAILED: {msg}", file=sys.stderr)
            return 1
        print(f"Validation: {msg}")

        # Stage 7: Cross-Validation (if requested)
        comparisons = []
        stats = None
        if args.validate:
            t1 = time.time()
            print("Running cross-validation...")
            comparisons = cross_validator.compare_sources(
                args.pdf, args.png, fields, geometries, info, layout, args.dpi
            )
            stats = cross_validator.compute_accuracy_stats(comparisons)
            timings["cross_validation"] = time.time() - t1
            if stats:
                print(f"Accuracy: mean={stats.mean_error:.4f}pt "
                      f"median={stats.median_error:.4f}pt "
                      f"max={stats.max_error:.4f}pt "
                      f"RMS={stats.rms_error:.4f}pt "
                      f"std={stats.std_dev:.4f}pt "
                      f"(n={stats.field_count})")

        # Stage 8: Runtime Generation
        t1 = time.time()

        conf = None
        if comparisons:
            overall_err = stats.mean_error if stats else 0
            conf = cross_validator.compute_confidence_scores(comparisons, overall_err)

        runtime = runtime_generator.generate_runtime(
            info, fields, geometries, args.dpi, workbook_info=info
        )
        if conf:
            runtime["_validation"] = {
                "confidence": conf,
                "accuracy": {
                    "mean_error_pt": round(stats.mean_error, 4) if stats else None,
                    "max_error_pt": round(stats.max_error, 4) if stats else None,
                    "rms_error_pt": round(stats.rms_error, 4) if stats else None,
                }
                if stats
                else None,
            }

        with open(args.output, "w", encoding="utf-8") as f:
            json.dump(runtime, f, indent=2)

        with open(args.output, "rb") as f:
            output_hash = hashlib.sha256(f.read()).hexdigest()

        timings["runtime_generation"] = time.time() - t1
        print(f"Wrote {args.output}")
        print(f"SHA-256: {output_hash[:16]}...")

        # Stage 9: Debug Output
        if args.debug_dir:
            ensure_dir(args.debug_dir)
            t1 = time.time()

            vis = debug_visualizer.generate_all(
                args.pdf,
                args.png,
                geometries,
                comparisons if args.validate else None,
                layout,
                info,
                args.debug_dir,
                args.dpi,
            )
            overlay = debug_visualizer.generate_overlay(
                args.pdf, args.png, geometries, args.debug_dir, args.dpi
            )
            dump = debug_overlay.generate_dump(
                geometries, layout, info, args.debug_dir
            )
            report = debug_overlay.generate_report(
                geometries, layout, info, args.debug_dir
            )

            timings["debug_output"] = time.time() - t1
            print(f"Debug images: {', '.join(vis.values())}")
            print(f"Debug report: {dump}")

            # Stage 10: Geometry Trace (if requested)
            if args.trace:
                t1 = time.time()
                trace_path = geometry_tracer.generate_geometry_trace(
                    fields, geometries, info, layout, args.dpi, args.debug_dir
                )
                timings["geometry_trace"] = time.time() - t1
                print(f"Trace: {trace_path}")

            # Stage 11: Error Diagnostics (if requested)
            if args.diagnose:
                t1 = time.time()
                diagnostic = error_diagnostics.diagnose(
                    geometries, layout, info,
                    comparisons if args.validate else None,
                    args.dpi,
                )
                diag_path = error_diagnostics.generate_diagnostic_report(
                    diagnostic, args.debug_dir
                )
                timings["diagnostics"] = time.time() - t1
                print(f"Diagnostics: {diagnostic.summary}")
                print(f"  Severity: {diagnostic.severity}")
                for cat in diagnostic.categories:
                    print(f"  [{cat['severity']}] {cat['category']}")

            # Stage 12: Workbook Analysis (if requested or always with debug-dir)
            if args.analyze:
                t1 = time.time()
                wb_analysis_path = workbook_analyzer.analyze_workbook(
                    args.excel, args.debug_dir
                )
                timings["workbook_analysis"] = time.time() - t1
                print(f"Workbook analysis: {wb_analysis_path}")

            # Stage 13: XML Dump (always when debug-dir is set)
            t1 = time.time()
            xml_dir = workbook_reader.xml_dump_to_debug(args.excel, args.debug_dir)
            timings["xml_dump"] = time.time() - t1
            if xml_dir:
                print(f"XML dump: {xml_dir}")

            # Stage 14: Workbook Reader Report (always when debug-dir is set)
            t1 = time.time()
            wbr_path = workbook_reader.generate_workbook_reader_report(info, args.debug_dir)
            timings["workbook_reader_report"] = time.time() - t1
            print(f"Reader report: {wbr_path}")

        total = time.time() - t0
        timings["total"] = total
        print(f"\n{'='*50}")
        print("Performance:")
        for stage, dur in timings.items():
            pct = dur / total * 100 if total > 0 else 0
            print(f"  {stage:25s}: {dur:.3f}s ({pct:5.1f}%)")
        print(f"{'='*50}")

        return 0

    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        traceback.print_exc(file=sys.stderr)

        if args.debug_dir:
            ensure_dir(args.debug_dir)
            lines = ["# Error Diagnostic Report", "",
                     "## Error", f"`{traceback.format_exc()}`"]
            path = os.path.join(args.debug_dir, "diagnostic_report.md")
            with open(path, "w") as f:
                f.write("\n".join(lines))
            print(f"Error diagnostic: {path}")

        return 1


if __name__ == "__main__":
    sys.exit(main())