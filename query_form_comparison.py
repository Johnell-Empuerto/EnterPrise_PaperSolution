import psycopg2
conn = psycopg2.connect(host='127.0.0.1', port=5432, dbname='irepodb', user='postgres', password='cimtops')
cur = conn.cursor()

# Search for forms with many clusters (matching the user's 6-field form)
cur.execute("""
SELECT dt.def_top_id, dsh.def_sheet_id, COUNT(dc.def_cluster_id) as cluster_count
FROM def_top dt
JOIN def_sheet dsh ON dsh.def_top_id = dt.def_top_id
JOIN def_cluster dc ON dc.def_sheet_id = dsh.def_sheet_id
GROUP BY dt.def_top_id, dsh.def_sheet_id
HAVING COUNT(dc.def_cluster_id) >= 5
ORDER BY dt.def_top_id DESC
LIMIT 10
""")
print("Forms with >= 5 clusters:")
for r in cur.fetchall():
    print(f"  def_top_id={r[0]} sheet_id={r[1]} clusters={r[2]}")

# Check def_top_id=546 which is most recent
cur.execute("""
SELECT dc.cluster_id, dc.cell_addr, dc.left_position, dc.right_position, dc.top_position, dc.bottom_position
FROM def_cluster dc
JOIN def_sheet dsh ON dc.def_sheet_id = dsh.def_sheet_id
WHERE dsh.def_top_id = 546
ORDER BY dc.cluster_id
""")
print("\ndef_top_id=546 clusters:")
for r in cur.fetchall():
    cid, addr, lp, rp, tp, bp = r
    print(f"  cluster={cid} addr={addr} L={lp} R={rp} T={tp} B={bp}")
    px_left = float(lp) * 2550
    px_right = float(rp) * 2550
    px_top = float(tp) * 3299
    px_bottom = float(bp) * 3299
    print(f"    pixels: L={px_left:.1f} R={px_right:.1f} T={px_top:.1f} B={px_bottom:.1f} W={px_right-px_left:.1f} H={px_bottom-px_top:.1f}")

# Now compare coordinate calculation methods
print("\n\n=== COORDINATE COMPARISON ===")
print("User's upload (FormTest.xlsx):")
user_fields = [
    ("A1:B2", 875, 1289.6, 400, 120),
    ("C1:D2", 1275, 1289.6, 400, 120),
    ("A3:D4", 875, 1409.6, 800, 120),
    ("A6:D7", 875, 1589.5, 800, 120),
    ("A9:D10", 875, 1769.5, 800, 120),
    ("A12", 875, 1949.4, 200, 60),
]
for addr, l, t, w, h in user_fields:
    print(f"  {addr:<8} left={l:.1f} top={t:.1f} width={w:.1f} height={h:.1f}")

print("\n\nLegacy system calculation for the same cells (if they were stored):")
print("The legacy system stores normalized 0-1 ratios.")
print("Pixel = ratio * pageDimension")
print("For pageWidth=2550, pageHeight=3299:")
print("  A field at left=875 would have ratio = 875/2550 =", 875/2550)
print("  A field at top=1289.6 would have ratio = 1289.6/3299 =", 1289.6/3299)
print("  A field with width=400 would have width_ratio = 400/2550 =", 400/2550)
print("  A field with height=120 would have height_ratio = 120/3299 =", 120/3299)

# Check the user's specific concern - width mismatch
print("\n\n=== WIDTH ANALYSIS ===")
print("Field A12 (non-merged):")
print("  Pixel width from COM: 200px")
print("  Point width from COM: 48pt")
scaleX = 4.166667
recalc_w = 48 * scaleX
print(f"  Recalculated width: 48pt * {scaleX} = {recalc_w:.1f}px")
print(f"  Actual stored width: 200px")
print(f"  Difference: {200 - recalc_w:.1f}px")

print("\nField A1 (merged A1:B2):")
print("  Pixel width: 400px")
print("  Point width: 96pt")
recalc_w = 96 * scaleX
print(f"  Recalculated: 96pt * {scaleX} = {recalc_w:.1f}px")
print(f"  Actual stored: 400px")
print(f"  Difference: {400 - recalc_w:.1f}px")

# Check border effect
print("\n\n=== BORDER EFFECT ON VISUAL WIDTH ===")
print("The input has: border: 1px solid #F9A825")
print("With boxSizing: border-box, width: 100% = total 400px")
print("Content area = 400 - 2*1(border) - 2*4(padding) = 390px")
print("Visual yellow area (background) fills content + padding = 398px")
print("Cell PNG width = 400px exactly")
print("So yellow overlay appears 2px narrower than cell (1px each side)")

conn.close()
