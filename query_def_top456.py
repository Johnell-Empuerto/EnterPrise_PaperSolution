import psycopg2
conn = psycopg2.connect(host='127.0.0.1', port=5432, dbname='irepodb', user='postgres', password='cimtops')
cur = conn.cursor()

# Get all def_top records with their IDs to find forms matching "FormTest"
cur.execute("SELECT def_top_id, designer_version FROM def_top WHERE def_top_id = 456")
r = cur.fetchone()
print(f"def_top_id=456: designer_version={r[1]}")

# Get sheet dimensions and pixel dimensions
cur.execute("""
SELECT dsh.def_sheet_id, dsh.def_sheet_no, dsh.def_sheet_name,
       COALESCE(ds.left_position, '0') as left_pos,
       COALESCE(ds.right_position, '0') as right_pos,
       COALESCE(ds.top_position, '0') as top_pos,
       COALESCE(ds.bottom_position, '0') as bottom_pos,
       ds.cell_addr, ds.cluster_id
FROM def_sheet dsh
LEFT JOIN def_cluster ds ON ds.def_sheet_id = dsh.def_sheet_id
WHERE dsh.def_top_id = 456
ORDER BY dsh.def_sheet_no, ds.cluster_id
""")
rows = cur.fetchall()
print(f"\nClusters for def_top_id=456 ({len(rows)} found):")
print(f"{'sheet_id':<12} {'sheet_no':<10} {'sheet_name':<15} {'cluster':<8} {'cell_addr':<15} {'left_pos':<14} {'right_pos':<14} {'top_pos':<14} {'bottom_pos':<14}")
print("-"*120)
for r in rows:
    sid, sno, sname, lp, rp, tp, bp, addr, cid = r
    print(f"{str(sid):<12} {str(sno):<10} {str(sname or ''):<15} {str(cid):<8} {str(addr or ''):<15} {float(lp):<14.7f} {float(rp):<14.7f} {float(tp):<14.7f} {float(bp):<14.7f}")

# Convert legacy 0-1 ratios to pixels at 2550x3299 for comparison
print("\n=== Legacy ratio to Pixel conversion (page = 2550x3299) ===")
print("{:<15} {:>12} {:>12} {:>12} {:>12} {:>10} {:>10}".format("cell_addr", "leg_L_px", "leg_T_px", "leg_R_px", "leg_B_px", "leg_W", "leg_H"))
print("-"*85)
for r in rows:
    sid, sno, sname, lp, rp, tp, bp, addr, cid = r
    if addr is None:
        continue
    px_left = float(lp) * 2550
    px_right = float(rp) * 2550
    px_top = float(tp) * 3299
    px_bottom = float(bp) * 3299
    px_w = px_right - px_left
    px_h = px_bottom - px_top
    print("{:<15} {:>12.1f} {:>12.1f} {:>12.1f} {:>12.1f} {:>10.1f} {:>10.1f}".format(str(addr), px_left, px_top, px_right, px_bottom, px_w, px_h))

cur.close()
conn.close()
