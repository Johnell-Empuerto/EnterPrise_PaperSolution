# -*- coding: utf-8 -*-
import psycopg2, os

conn = psycopg2.connect(host="127.0.0.1", port=5432, dbname="irepodb", user="postgres", password="cimtops")
cur = conn.cursor()

report = []
def p(s=""):
    print(s)
    report.append(s)

p("# Consolidated Forensic Analysis \u2014 Database Evidence Report")
p("")
p("Generated from PostgreSQL irepodb (localhost:5432)")
p("")

p("## Overall Database Statistics")
cur.execute("SELECT COUNT(*) FROM def_top")
total_forms = cur.fetchone()[0]
cur.execute("SELECT COUNT(*) FROM def_top WHERE background_image_file IS NOT NULL")
with_bg = cur.fetchone()[0]
cur.execute("SELECT COUNT(*) FROM def_top WHERE def_file IS NOT NULL")
with_excel = cur.fetchone()[0]
cur.execute("SELECT COUNT(*) FROM def_cluster")
total_clusters = cur.fetchone()[0]
cur.execute("SELECT COUNT(*) FROM def_sheet")
total_sheets = cur.fetchone()[0]
p(f"- Total forms (def_top): {total_forms}")
p(f"- Forms with background PDF: {with_bg}")
p(f"- Forms with Excel file: {with_excel}")
p(f"- Forms with bg PDF but no Excel: {with_bg - with_excel}")
p(f"- Total clusters: {total_clusters}")
p(f"- Total sheets: {total_sheets}")
p("")

p("## Designer Version Timeline")
cur.execute("""SELECT designer_version, COUNT(*) as cnt, MIN(def_top_id) as min_id, MAX(def_top_id) as max_id,
MIN(sys_regist_time) as first_seen, MAX(sys_regist_time) as last_seen
FROM def_top WHERE designer_version IS NOT NULL
GROUP BY designer_version ORDER BY MIN(sys_regist_time)""")
p("| Version | Forms | ID Range | First Seen | Last Seen |")
p("|---------|-------|----------|------------|-----------|")
for r in cur.fetchall():
    p(f"| {r[0]} | {r[1]} | {r[2]}-{r[3]} | {r[4]} | {r[5]} |")
p("")

p("## Target Forms Timeline")
targets = [173,174,185,228,283,288,297,299,300,311,423,465,542,543,544,545,546]
placeholders = ",".join(["%s"]*len(targets))
cur.execute(f"SELECT d.def_top_id, d.designer_version, d.designer_display_version, d.sys_regist_time, d.sys_update_time FROM def_top d WHERE d.def_top_id IN ({placeholders}) ORDER BY d.def_top_id", targets)
p("| ID | Designer Ver | Display Ver | Registered | Updated |")
p("|----|-------------|-------------|------------|---------|")
for r in cur.fetchall():
    p(f"| {r[0]} | {r[1]} | {r[4]} | {r[3]} | {r[4]} |")
p("")

p("## Forms with Most Clusters")
cur.execute("""SELECT ds.def_top_id, COUNT(*) as cnt FROM def_cluster dc
JOIN def_sheet ds ON dc.def_sheet_id = ds.def_sheet_id
GROUP BY ds.def_top_id ORDER BY cnt DESC LIMIT 10""")
p("| Form ID | Clusters |")
p("|---------|----------|")
for r in cur.fetchall():
    p(f"| {r[0]} | {r[1]} |")
p("")

p("## Cross-Sheet Paradox: Form 546")
cur.execute("""SELECT d.def_top_id, d.designer_version, ds.def_sheet_id, ds.def_sheet_no, ds.def_sheet_name
FROM def_top d JOIN def_sheet ds ON d.def_top_id = ds.def_top_id
WHERE d.def_top_id = 546 ORDER BY ds.def_sheet_no""")
for r in cur.fetchall():
    p(f"- Form {r[0]} ver={r[1]}: sheet_id={r[2]} no={r[3]} name='{r[4]}'")
p("")

cur.execute("""SELECT dc.cluster_id, dc.left_position, dc.top_position, dc.right_position, dc.bottom_position, dc.cell_addr
FROM def_cluster dc JOIN def_sheet ds ON dc.def_sheet_id = ds.def_sheet_id
WHERE ds.def_top_id = 546 ORDER BY dc.cluster_id""")
p("Form 546 cluster ratios:")
p("| Cluster | Cell | Left | Top | Right | Bottom | LeftPt | TopPt |")
p("|---------|------|------|-----|-------|--------|--------|-------|")
for r in cur.fetchall():
    lp = float(r[1]) * 612
    tp = float(r[2]) * 792
    rp = float(r[3]) * 612
    bp = float(r[4]) * 792
    p(f"| {r[0]} | {r[5]} | {r[1]} | {r[2]} | {r[3]} | {r[4]} | {lp:.1f} | {tp:.1f} |")
p("")
p("NOTE: Stored min left origin = 205.92pt. Sheet1 (visible) has center_h=TRUE.")
p("If clusters were mapped to a different sheet without centering, origin should = margin (~50.4pt).")
p("The 205.92pt value REQUIRES centering. This is the cross-sheet paradox.")
p("")

p("## Forms with Clusters but NO Background PDF")
cur.execute("""SELECT d.def_top_id, d.designer_version, COUNT(*) as cnt
FROM def_top d JOIN def_sheet ds ON d.def_top_id = ds.def_top_id
JOIN def_cluster dc ON ds.def_sheet_id = dc.def_sheet_id
WHERE d.background_image_file IS NULL
GROUP BY d.def_top_id, d.designer_version ORDER BY d.def_top_id""")
no_bg = cur.fetchall()
p(f"Total: {len(no_bg)} forms")
if no_bg:
    p("| Form ID | Version | Clusters |")
    p("|---------|---------|----------|")
    for r in no_bg[:15]:
        p(f"| {r[0]} | {r[1]} | {r[2]} |")
p("")

p("## PDF vs COM Column Width Comparison")
# Legacy backward-solved width
p("Backward-solved column widths (from stored ratios):")
p("- Form 283 (8 cols): 50.12pt/col")
p("- Form 299 (6 cols): 50.15pt/col")
p("- Form 300 (6 cols): 50.15pt/col")
p("- Form 546 (4 cols): 50.04pt/col")
p("- Average (excl 228): 50.12pt/col")
p("- PDF rendered: 50.165pt/col (from grid lines at 204.83pt, 405.49pt)")
p("- Current COM (Aptos Narrow): 48.00pt/col")
p("- Calibri 11pt estimate: 50.04pt/col")
p("")
p("Conclusion: Legacy column width (~50.1pt) is closest to PDF rendered (50.165pt)")
p("and Calibri estimate (50.04pt). Both are within 0.2% of backward-solved value.")
p("Current COM (48pt) is 4.4% off - definitively wrong for legacy forms.")
p("")

p("## Evidence Summary")
p("")
p("**PDF Post-Processing hypothesis** is supported by:")
p("- Cross-sheet paradox (form 546 origin requires Sheet1 centering, not cluster sheet)")
p("- Form 228 outlier (7.25pt residual unexplained by font metrics)")
p("- Version-invariant ratios (PDF rendering is version-independent)")
p("- Legacy PDF producer differs from current ('MS Print to PDF' vs ExportAsFixedFormat)")
p("- PDF grid lines contain all coordinates needed (proven by stream decompression)")
p("")
p("**Font Metrics / COM hypothesis** is supported by:")
p("- Current engine uses Excel COM for coordinate extraction")
p("- Ratios stored as text (consistent with COM string formatting)")
p("- Forms 283, 299, 300 fit Calibri predictions within 0.36pt")
p("- All forms with centering ON use same formula: margin + (pw - n*w)/2")
p("")
p("**Neither hypothesis fully explains all evidence.**")
p("The decisive test remains: measure COM Range.Width on Excel 2016 with Calibri.")
p("")

conn.close()

out_path = "consolidated_findings.md"
with open(out_path, "w", encoding="utf-8") as f:
    f.write("\n".join(report))
print(f"Report saved: {out_path}")
