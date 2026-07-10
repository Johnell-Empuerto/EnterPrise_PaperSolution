# -*- coding: utf-8 -*-
import psycopg2, json, re

conn = psycopg2.connect(host="127.0.0.1", port=5432, dbname="irepodb", user="postgres", password="cimtops")
cur = conn.cursor()

out = []
out.append("Q1: Database search for PDF geometry evidence")
out.append("=" * 60)

# Search ALL columns for rendering/geometry keywords
keywords = [
    "pdf", "render", "geometry", "coord", "position", "page", "clip",
    "bbox", "rectangle", "vector", "grid", "normaliz", "cache", "temp",
    "log", "pixel", "dimension", "image_w", "image_h", "image_size",
    "resolution", "transform", "matrix", "bound", "annotation", "field"
]

cur.execute("SELECT table_name, column_name, data_type FROM information_schema.columns WHERE table_schema='public'")
all_cols = cur.fetchall()

matches = {}
for tbl, col, dtype in all_cols:
    cl = col.lower()
    tl = tbl.lower()
    for kw in keywords:
        if kw in cl or kw in tl:
            if kw not in matches:
                matches[kw] = []
            matches[kw].append(f"{tbl}.{col} ({dtype})")

for kw in sorted(matches.keys()):
    items = matches[kw]
    out.append(f"\n'{kw}' ({len(items)} matches):")
    for item in items[:25]:
        out.append(f"  {item}")

# Check specific tables for render pipeline columns
out.append("\n\n=== def_top_size ===")
cur.execute("SELECT column_name, data_type FROM information_schema.columns WHERE table_name='def_top_size' ORDER BY ordinal_position")
for r in cur.fetchall():
    out.append(f"  {r[0]} ({r[1]})")

out.append("\n=== rep_top_size ===")
cur.execute("SELECT column_name, data_type FROM information_schema.columns WHERE table_name='rep_top_size' ORDER BY ordinal_position")
for r in cur.fetchall():
    out.append(f"  {r[0]} ({r[1]})")

out.append("\n=== rep_fd_sheet ===")
cur.execute("SELECT column_name, data_type FROM information_schema.columns WHERE table_name='rep_fd_sheet' ORDER BY ordinal_position")
for r in cur.fetchall():
    out.append(f"  {r[0]} ({r[1]})")

out.append("\n=== common_document_irfd ===")
cur.execute("SELECT column_name, data_type FROM information_schema.columns WHERE table_name='common_document_irfd' ORDER BY ordinal_position")
for r in cur.fetchall():
    out.append(f"  {r[0]} ({r[1]})")

# Check PDF auto output settings
out.append("\n\n=== pdf_auto_output fields in def_top (sample) ===")
cur.execute("SELECT def_top_id, pdf_auto_output_sheet_select, pdf_auto_output_sheet_select_no FROM def_top WHERE pdf_auto_output_sheet_select IS NOT NULL LIMIT 10")
for r in cur.fetchall():
    out.append(f"  id={r[0]} sel={r[1][:50] if r[1] else None} no={r[2]}")

# Check backgrounds
out.append("\n\n=== Forms with background_image_file and thumbnail_file ===")
cur.execute("""
    SELECT def_top_id, 
           CASE WHEN background_image_file IS NOT NULL THEN 1 ELSE 0 END as has_bg,
           CASE WHEN thumbnail_file IS NOT NULL THEN 1 ELSE 0 END as has_thumb,
           CASE WHEN def_file IS NOT NULL THEN 1 ELSE 0 END as has_excel
    FROM def_top WHERE def_top_id IN (173, 174, 185, 228, 283, 299, 311, 423, 465, 542, 546)
    ORDER BY def_top_id
""")
for r in cur.fetchall():
    out.append(f"  id={r[0]} bg={'Y' if r[1] else 'N'} thumb={'Y' if r[2] else 'N'} excel={'Y' if r[3] else 'N'}")

out.append("\n\n=== conmas_operation_log columns ===")
cur.execute("SELECT column_name, data_type FROM information_schema.columns WHERE table_name='conmas_operation_log' ORDER BY ordinal_position")
for r in cur.fetchall():
    out.append(f"  {r[0]} ({r[1]})")

with open("q1_evidence.txt", "w", encoding="utf-8") as f:
    f.write("\n".join(out))
print("Q1 saved:", len(out), "lines")
conn.close()
