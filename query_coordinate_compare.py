import psycopg2
conn = psycopg2.connect(host='127.0.0.1', port=5432, dbname='irepodb', user='postgres', password='cimtops')
cur = conn.cursor()

# Get the actual background image dimensions from def_top_id=546
cur.execute("SELECT background_image_file FROM def_top WHERE def_top_id = 546")
bg_row = cur.fetchone()[0]
if bg_row:
    bg = bytes(bg_row)
    print(f"Background file size: {len(bg)} bytes")
    print(f"Header: {bg[:20]}")
    is_pdf = bg[:4] == b'%PDF'
    print(f"Is PDF: {is_pdf}")
    
    if is_pdf:
        import re
        pdf_text = bg.decode('latin-1', errors='ignore')
        mb = re.findall(r'/MediaBox\s*\[([^\]]+)\]', pdf_text)
        if mb:
            print(f"MediaBox: {mb}")
        # Check for other size hints
        rp = re.findall(r'/Resources.*?/XObject.*?/Width\s+(\d+)', pdf_text, re.DOTALL)
        if rp:
            print(f"Resource Width: {rp}")
        # Try to find page size
        ps = re.findall(r'/UserUnit\s+(\d+)', pdf_text)
        if ps:
            print(f"UserUnit: {ps}")

# Compare coordinates
print("\n\n=== COORDINATE COMPARISON: COM vs Legacy ===")

user = [
    ("A1:B2", 875, 1289.6, 400, 120),
    ("C1:D2", 1275, 1289.6, 400, 120),
    ("A3:D4", 875, 1409.6, 800, 120),
    ("A6:D7", 875, 1589.5, 800, 120),
    ("A9:D10", 875, 1769.5, 800, 120),
    ("A12", 875, 1949.4, 200, 60),
]

legacy = [
    ("A1:B2", 858.0, 1268.6, 412.5, 123.0),
    ("C1:D2", 1275.0, 1268.6, 417.0, 123.0),
    ("A3:D4", 858.0, 1396.1, 834.0, 123.0),
    ("A6:D7", 858.0, 1586.5, 834.0, 123.0),
    ("A9:D10", 858.0, 1777.0, 834.0, 123.0),
    ("A12", 858.0, 1967.4, 204.0, 61.5),
]

print(f"{'Cell':<10} {'Metric':<10} {'COM':<12} {'Legacy':<12} {'Diff':<10}")
print("-"*54)
for i in range(len(user)):
    addr = user[i][0]
    for j, metric in enumerate(["leftPx", "topPx", "widthPx", "heightPx"]):
        uv = user[i][1:][j]
        lv = legacy[i][1:][j]
        diff = uv - lv
        print(f"{addr:<10} {metric:<10} {uv:<12.1f} {lv:<12.1f} {diff:<+10.1f}")

# Analyze width discrepancy
print("\n\n=== WIDTH GAP ANALYSIS ===")
print("The COM overlay has a 1px border on each side (inside border-box).")
print("At 400px total width with 1px border + 4px padding each side:")
print("  Total border-box width: 400px")
print("  Border consumes: 2px")
print("  Padding consumes: 8px")
print("  Content width: 390px")
print("  Visual yellow area (content+padding): 398px")
print("  Cell PNG width: 400px")
print("  Gap per side: (400 - 398) / 2 = 1px")

print("\nLegacy system width for A1:B2: 412.5px")
print(f"COM width for A1:B2: 400px")
width_ratio_at_2550 = 412.5 / 2550
print(f"\nIf legacy renders at different page width:")
print(f"Legacy width ratio = {width_ratio_at_2550:.6f}")
for pw in [2550, 2480, 2400, 2200, 2000, 1700]:
    px = width_ratio_at_2550 * pw
    print(f"  At page width {pw}: legacy A1:B2 = {px:.1f}px")

# Check legacy scale
legacy_w = 412.5  # pixels at whatever resolution
# If Excel cell is 96pt wide:
cell_width_pt = 96
scale_for_legacy = legacy_w / cell_width_pt
print(f"\nLegacy scale for A1:B2: {legacy_w}px / 96pt = {scale_for_legacy:.4f}")
print(f"COM scale: {scale_for_legacy:.4f}")
print(f"Theoretical DPI from legacy scale: {scale_for_legacy * 72:.1f}")

conn.close()
