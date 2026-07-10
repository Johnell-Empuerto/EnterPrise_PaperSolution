import psycopg2
conn = psycopg2.connect(host='127.0.0.1', port=5432, dbname='irepodb', user='postgres', password='cimtops')
cur = conn.cursor()

# Get column names for def_cluster
cur.execute("SELECT column_name, data_type FROM information_schema.columns WHERE table_name = 'def_cluster'")
print("def_cluster columns:")
for r in cur.fetchall():
    print(f"  {r[0]} ({r[1]})")

# Check def_top_id=546 clusters 
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
    print(f"  cluster={cid} addr={addr}")
    print(f"    L={lp} R={rp} T={tp} B={bp}")
    px_left = float(lp) * 2550
    px_right = float(rp) * 2550
    px_top = float(tp) * 3299
    px_bottom = float(bp) * 3299
    print(f"    pixels: L={px_left:.1f} R={px_right:.1f} T={px_top:.1f} B={px_bottom:.1f} W={px_right-px_left:.1f} H={px_bottom-px_top:.1f}")

# Also check sizes for 546
cur.execute("SELECT * FROM def_top_size WHERE def_top_id = 546")
print("\ndef_top_size for 546:")
cols = [desc[0] for desc in cur.description]
for r in cur.fetchall():
    for i, c in enumerate(cols):
        print(f"  {c}: {r[i]}")

# Check the legacy system: what pixel dimensions does it use?
# Get the background image to check actual rendered size
cur.execute("SELECT def_top_id, LENGTH(background_image_file), LENGTH(thumbnail_file) FROM def_top WHERE def_top_id = 546")
r = cur.fetchone()
print(f"\ndef_top_id=546: bg_file_len={r[1]}, thumb_len={r[2]}")

conn.close()
