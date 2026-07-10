import psycopg2
conn = psycopg2.connect(host='127.0.0.1', port=5432, dbname='irepodb', user='postgres', password='cimtops')
cur = conn.cursor()

# List recent def_top entries to find which one matches "FormTest"
cur.execute("""
SELECT def_top_id, sys_regist_time, designer_version 
FROM def_top 
ORDER BY def_top_id DESC 
LIMIT 20
""")
print("Recent def_top entries:")
for r in cur.fetchall():
    print(f"  ID={r[0]}  time={r[1]}  version={r[2]}")

# Check def_top_id=456 details
cur.execute("""
SELECT dsh.def_sheet_id, dsh.def_sheet_no, dsh.def_sheet_name
FROM def_sheet dsh
WHERE dsh.def_top_id = 456
""")
print("\ndef_top_id=456 sheets:")
for r in cur.fetchall():
    print(f"  sheet_id={r[0]}  sheet_no={r[1]}  name={r[2]}")

# Check cluster details for def_top_id=456
cur.execute("""
SELECT ds.cluster_id, ds.cell_addr, ds.left_position, ds.right_position, ds.top_position, ds.bottom_position
FROM def_cluster ds
JOIN def_sheet dsh ON ds.def_sheet_id = dsh.def_sheet_id
WHERE dsh.def_top_id = 456
ORDER BY ds.cluster_id
""")
print("\nAll clusters for def_top_id=456:")
for r in cur.fetchall():
    cid, addr, lp, rp, tp, bp = r
    px_left = float(lp) * 2550
    px_right = float(rp) * 2550
    px_top = float(tp) * 3299
    px_bottom = float(bp) * 3299
    print(f"  cluster={cid} addr={addr}")
    print(f"    ratios: L={lp} R={rp} T={tp} B={bp}")
    print(f"    pixels (2550x3299): L={px_left:.1f} R={px_right:.1f} T={px_top:.1f} B={px_bottom:.1f}")
    print(f"    size: W={px_right-px_left:.1f} H={px_bottom-px_top:.1f}")

# Also check the XML to see stored pixel coordinates
cur.execute("""
SELECT xml_data::text 
FROM def_top 
WHERE def_top_id = 456
""")
xml = cur.fetchone()[0]

# Extract cluster coordinates from XML
import re
clusters_xml = re.findall(r'<cluster[^>]*>(.*?)</cluster>', xml, re.DOTALL)
print(f"\nXML clusters found: {len(clusters_xml)}")
for i, cx in enumerate(clusters_xml):
    print(f"\n  Cluster {i}:")
    # Extract various coordinate fields
    for tag in ['cluster_id', 'cell_address', 'left', 'right', 'top', 'bottom', 'left_pt', 'top_pt', 'width_pt', 'height_pt']:
        m = re.search(r'<' + tag + r'>(.*?)</' + tag + r'>', cx)
        if m:
            print(f"    {tag}: {m.group(1)}")

cur.close()
conn.close()
