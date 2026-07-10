# -*- coding: utf-8 -*-
"""
Coordinate Validation Framework v2
- Infers print area from cluster cell addresses (most forms)
- Extracts actual XLSX for key forms (228, 283, 299, 300, 546)
- Compares engine predictions vs stored ratios
- Classifies errors automatically
- No engine code modifications
"""
import psycopg2, re, os, json, zipfile
from collections import defaultdict
from io import BytesIO

conn = psycopg2.connect(host="127.0.0.1", port=5432, dbname="irepodb", user="postgres", password="cimtops")
cur = conn.cursor()

print("=== Gathering form + cluster data ===")

cur.execute("SELECT def_top_id, designer_version, sys_regist_time, xml_data FROM def_top ORDER BY def_top_id")
all_forms = {r[0]: r for r in cur.fetchall()}

cur.execute("""
    SELECT ds.def_top_id, dc.cluster_id, dc.left_position, dc.top_position,
           dc.right_position, dc.bottom_position, dc.cell_addr
    FROM def_cluster dc
    JOIN def_sheet ds ON dc.def_sheet_id = ds.def_sheet_id
    ORDER BY ds.def_top_id, dc.cluster_id
""")
all_clusters = defaultdict(list)
for r in cur.fetchall():
    all_clusters[r[0]].append(r)

print(f"Forms: {len(all_forms)}, With clusters: {len(all_clusters)}")

COL_WIDTHS = {"Aptos48": 48.00, "Calibri50.04": 50.04, 
              "PDF50.165": 50.165, "BackSolv50.12": 50.12}

def col_letter_to_num(s):
    n = 0
    for c in s.upper():
        n = n * 26 + (ord(c) - 64)
    return n

def extract_xlsx_print_area(fid):
    """Extract true print area from stored XLSX."""
    cur.execute("SELECT def_file FROM def_top WHERE def_top_id = %s", (fid,))
    row = cur.fetchone()
    if not row or not row[0]:
        return None, None
    try:
        zf = zipfile.ZipFile(BytesIO(bytes(row[0])))
        wb = zf.read("xl/workbook.xml").decode("utf-8")
        pa = re.search(r'Print_Area[^>]*>([^<]+)</definedName>', wb)
        if pa:
            ref = pa.group(1)
            m = re.search(r'\$(\w+)\$(\d+):\$(\w+)\$(\d+)', ref)
            if m:
                c1 = col_letter_to_num(m.group(1))
                c2 = col_letter_to_num(m.group(3))
                return ref, c2 - c1 + 1
        zf.close()
    except:
        pass
    return None, None

def parse_cell_addr_range(cell_addr):
    """Extract column range from cell address like $A$1:$B$2."""
    m = re.match(r'\$(\w+)\$(\d+):\$(\w+)\$(\d+)', cell_addr)
    if m:
        return col_letter_to_num(m.group(1)), col_letter_to_num(m.group(3))
    return None, None

results = []
errors_x = []
class_counts = defaultdict(int)
class_forms = defaultdict(set)

# Pre-extract XLSX data for key forms
key_forms = {228, 283, 299, 300, 546}
xlsx_cache = {}
for fid in key_forms:
    ref, n = extract_xlsx_print_area(fid)
    xlsx_cache[fid] = (ref, n)
    if n:
        print(f"  Form {fid}: XLSX PA={ref} ({n} cols)")

for fid in sorted(all_clusters.keys()):
    frow = all_forms.get(fid)
    if not frow:
        continue
    designer_ver = frow[1] or ""
    regist_time = str(frow[2] or "")[:10]
    xml_data = frow[3]
    if not xml_data:
        continue
    clusters = all_clusters[fid]
    
    def ext(p, d=None):
        m = re.search(p, xml_data)
        return m.group(1) if m else d
    
    pw = float(ext(r'<width>(\d+)</width>', "612"))
    ph = float(ext(r'<height>(\d+)</height>', "792"))
    ml = float(ext(r'<marginLeft>([\d.]+)</marginLeft>', "51.02"))
    mr = float(ext(r'<marginRight>([\d.]+)</marginRight>', "51.02"))
    ch_str = ext(r'<centerH>([^<]+)</centerH>', "0")
    center_h = ch_str.lower() in ("1", "true")
    
    # Parse cluster ratios and infer print area from cell addresses
    min_col_global = 999
    max_col_global = 0
    min_l = 1.0
    min_t = 1.0
    for c in clusters:
        lr = float(c[2])
        tr = float(c[3])
        if lr < min_l: min_l = lr
        if tr < min_t: min_t = tr
        cell = c[6] or ""
        c1, c2 = parse_cell_addr_range(cell)
        if c1: min_col_global = min(min_col_global, c1)
        if c2: max_col_global = max(max_col_global, c2)
    
    min_l_pt = min_l * pw
    min_t_pt = min_t * ph
    printable_w = pw - ml - mr
    
    # Infer print area columns
    n_cols = 0
    pa_ref = ""
    if fid in xlsx_cache and xlsx_cache[fid][1]:
        n_cols = xlsx_cache[fid][1]
        pa_ref = xlsx_cache[fid][0] or ""
    elif max_col_global >= min_col_global and max_col_global > 0:
        n_cols = max_col_global - min_col_global + 1
    
    # Infer centering from stored origin
    inferred_center = center_h
    if not center_h and min_l_pt > ml + 5:
        inferred_center = True
    
    # Backward-solve column width
    stored_w = None
    if n_cols > 0 and inferred_center and printable_w > 0:
        stored_w = (printable_w - 2 * (min_l_pt - ml)) / n_cols
    
    is_default = bool(stored_w and 48 <= stored_w <= 53)
    
    # Try each column width candidate
    best_err = 999.0
    best_cand = "N/A"
    best_cw = 0
    
    if n_cols > 0 and inferred_center and is_default:
        for cname, cw in COL_WIDTHS.items():
            rw = n_cols * cw
            pred = ml + (printable_w - rw) / 2.0 if rw < printable_w else ml
            err = abs(pred - min_l_pt)
            if err < best_err:
                best_err = err
                best_cand = cname
                best_cw = cw
    
    # Classify
    if n_cols == 0:
        cls = "no_pa_data"
    elif not inferred_center:
        err_x = min_l_pt - ml
        if abs(err_x) < 0.5:
            cls = "rounding"
        elif abs(err_x) < 5:
            cls = "margin_mismatch"
        else:
            cls = "no_center_offset"
    elif not is_default:
        cls = "explicit_col_mismatch"
    elif best_err < 0.5:
        cls = "rounding"
    elif best_err < 2:
        cls = f"def_col_{best_cand}"
    elif best_err < 5:
        cls = "def_col_mismatch"
    else:
        cls = f"def_col_large_{best_err:.1f}pt"
    
    class_counts[cls] += 1
    class_forms[cls].add(fid)
    
    err_val = best_err if best_err < 999 else (min_l_pt - ml if not inferred_center else 999)
    
    results.append({
        "form_id": fid, "ver": designer_ver, "date": regist_time,
        "pw": pw, "ph": ph, "ml": round(ml,2), "mr": round(mr,2),
        "center_h": 1 if inferred_center else 0,
        "n_cols": n_cols, "n_clusters": len(clusters),
        "min_l_r": round(min_l,7), "min_l_pt": round(min_l_pt,2),
        "printable_w": round(printable_w,2),
        "stored_w": round(stored_w,3) if stored_w else 0,
        "is_default": 1 if is_default else 0,
        "best_cw": round(best_cw,3) if best_cw else 0,
        "best_err": round(best_err,3) if best_err < 999 else 999,
        "classification": cls, "err_x": round(err_val,3),
        "pa_ref": pa_ref,
    })
    errors_x.append(best_err if best_err < 999 else 999)

print(f"Processed {len(results)} forms")

# CSV
h = ["form_id","ver","date","pw","ph","ml","mr","center_h","n_cols",
     "n_clusters","min_l_r","min_l_pt","printable_w","stored_w",
     "is_default","best_cw","best_err","classification","err_x","pa_ref"]
with open("coordinate_validation.csv", "w", encoding="utf-8") as f:
    f.write(",".join(h) + "\n")
    for r in results:
        f.write(",".join(str(r.get(k,"")) for k in h) + "\n")
print("CSV: coordinate_validation.csv")

# Histogram
bins = [0, 0.5, 1, 2, 3, 5, 10, 20, 50, 100, 999]
cnts = [0]*(len(bins)-1)
for e in errors_x:
    for i in range(len(bins)-1):
        if bins[i] <= e < bins[i+1]: cnts[i] += 1; break
with open("error_histogram.csv", "w") as f:
    f.write("bin_min,bin_max,count\n")
    for i in range(len(bins)-1):
        f.write(f"{bins[i]},{bins[i+1]},{cnts[i]}\n")

# Report
with open("coordinate_validation_report.txt", "w", encoding="utf-8") as f:
    f.write("=" * 80 + "\nCOORDINATE VALIDATION REPORT\n" + "=" * 80 + "\n\n")
    f.write(f"Total forms: {len(results)}, Total clusters: {sum(r['n_clusters'] for r in results)}\n\n")
    
    f.write("CLASSIFICATION SUMMARY\n" + "-" * 50 + "\n")
    for cls in sorted(class_counts, key=lambda c: -class_counts[c]):
        cnt = class_counts[cls]
        sample = ", ".join(str(f) for f in list(class_forms[cls])[:3])
        if len(class_forms[cls]) > 3:
            sample += f" ...(+{len(class_forms[cls])-3})"
        f.write(f"{cls:<35} {cnt:<5} {100*cnt/len(results):5.1f}%  [{sample}]\n")
    
    f.write("\nERROR HISTOGRAM\n" + "-" * 50 + "\n")
    for i in range(len(bins)-1):
        label = f"{bins[i]:.1f}-{bins[i+1]:.1f}" if bins[i+1] < 999 else f"{bins[i]:.1f}+"
        f.write(f"  {label:<10} {'#'*min(cnts[i],50)} {cnts[i]}\n")
    
    ea = [abs(e) for e in errors_x if e < 999]
    if ea:
        me = sum(ea)/len(ea)
        se = sorted(ea)
        p95 = se[int(len(se)*0.95)]
        near = sum(1 for e in ea if e < 0.5)
        f.write(f"\nMean err: {me:.3f}pt, P95: {p95:.3f}pt, <0.5pt: {near}/{len(ea)} ({100*near/len(ea):.1f}%)\n")
    
    f.write("\nKEY FORMS DETAIL\n" + "-" * 80 + "\n")
    f.write(f"{'FID':<6} {'Ver':<14} {'Cols':<5} {'Ctr':<4} {'StoredL':<9} {'StoredW':<9} {'BestCW':<10} {'Err':<7} {'Class':<30}\n")
    f.write("-" * 80 + "\n")
    for fid in [228, 283, 299, 300, 546, 173, 174, 185, 311, 465, 542, 543]:
        r = next((x for x in results if x["form_id"] == fid), None)
        if r:
            f.write(f"{r['form_id']:<6} {str(r['ver'])[:14]:<14} {r['n_cols']:<5} "
                    f"{'Y' if r['center_h'] else 'N':<4} {r['min_l_pt']:<9.2f} "
                    f"{r['stored_w']:<9.3f} {r['best_cw']:<10.3f} "
                    f"{r['best_err']:<7.3f} {r['classification']:<30}\n")
    
    # Worst 10
    f.write("\nWORST 10 FORMS\n" + "-" * 80 + "\n")
    for r in sorted(results, key=lambda x: -x["best_err"])[:10]:
        f.write(f"{r['form_id']:<6} {str(r['ver'])[:14]:<14} {r['n_cols']:<5} "
                f"{'Y' if r['center_h'] else 'N':<4} {r['min_l_pt']:<9.2f} "
                f"{r['best_cw']:<10.3f} {r['best_err']:<7.3f} {r['classification']:<30}\n")
    
    # Best 10
    f.write("\nBEST 10 FORMS\n" + "-" * 80 + "\n")
    for r in sorted(results, key=lambda x: x["best_err"])[:10]:
        f.write(f"{r['form_id']:<6} {str(r['ver'])[:14]:<14} {r['n_cols']:<5} "
                f"{'Y' if r['center_h'] else 'N':<4} {r['min_l_pt']:<9.2f} "
                f"{r['best_cw']:<10.3f} {r['best_err']:<7.3f} {r['classification']:<30}\n")
    
    # Minimal changes
    f.write("\n" + "=" * 80 + "\nMINIMAL CHANGES\n" + "=" * 80 + "\n")
    no_pa = class_counts.get("no_pa_data", 0)
    total = len(results)
    with_data = total - no_pa
    
    default_ok = sum(1 for r in results if "rounding" == r["classification"])
    default_near = sum(1 for r in results if "def_col_" in r["classification"] and "large" not in r["classification"] and "mismatch" not in r["classification"])
    default_mismatch = sum(1 for r in results if "def_col_mismatch" == r["classification"])
    default_large = sum(1 for r in results if "large" in r["classification"])
    explicit = sum(1 for r in results if "explicit" in r["classification"])
    no_center = sum(1 for r in results if "no_center" in r["classification"])
    margin = sum(1 for r in results if "margin" in r["classification"])
    
    f.write(f"\nForms with analyzable data: {with_data}/{total}\n")
    f.write(f"1. Already correct (<0.5pt): {default_ok}\n")
    f.write(f"2. Fix default col width to 50.1pt: {default_near + default_mismatch}\n")
    f.write(f"3. Need explicit col widths from XLSX: {explicit}\n")
    f.write(f"4. Fix no-centering offset: {no_center}\n")
    f.write(f"5. Form 228 pattern (large residual): {default_large}\n")
    f.write(f"6. No print area data (needs XLSX extract): {no_pa}\n")
    f.write(f"\nRECOMMENDED CHANGES:\n")
    f.write(f"  A) In ExcelCaptureService, replace Range.Width with (n_cols * 50.1) for default-col forms\n")
    f.write(f"  B) For explicit-col forms, read <cols> from XLSX worksheet XML\n")
    f.write(f"  C) Apply actual PDF-rendered content width for centering calculation\n")

# Confidence JSON
with open("classification_confidence.json", "w") as f:
    json.dump({
        "total": len(results),
        "classifications": {cls: {"count": class_counts[cls], "pct": round(100*class_counts[cls]/len(results),1),
                                  "forms": list(class_forms[cls])[:10]} for cls in class_counts},
    }, f, indent=2)

print("All outputs generated.")
print(" - coordinate_validation.csv")
print(" - coordinate_validation_report.txt") 
print(" - error_histogram.csv")
print(" - classification_confidence.json")
print("\nDone.")
