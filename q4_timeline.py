import psycopg2
conn = psycopg2.connect(host="127.0.0.1", port=5432, dbname="irepodb", user="postgres", password="cimtops")
cur = conn.cursor()

# Q4: Timeline - get all forms with designer version and timestamps
cur.execute("SELECT def_top_id, designer_version, sys_regist_time, sys_update_time FROM def_top WHERE designer_version IS NOT NULL ORDER BY def_top_id")
rows = cur.fetchall()
print(f"Q4: {len(rows)} forms with designer_version")
print("=" * 80)

# Group by designer version
versions = {}
for r in rows:
    v = r[1]
    if v not in versions:
        versions[v] = []
    versions[v].append(r[0])

print("\nBy designer version:")
for v in sorted(versions.keys(), key=lambda x: [int(y) for y in x.split(".")]):
    ids = versions[v]
    sample = [r for r in rows if r[0] in ids[:3]]
    print(f"  v{v}: {len(ids)} forms - ids: {min(ids)}-{max(ids)}")
    for s in sample:
        print(f"    id={s[0]} sys_regist={s[2]} sys_update={s[3]}")

# Target forms detailed timeline
targets = [173, 174, 185, 228, 283, 288, 297, 299, 300, 311, 423, 465, 542, 543, 544, 545, 546]
cur.execute("SELECT def_top_id, designer_version, sys_regist_time, sys_update_time, updated_at FROM def_top WHERE def_top_id IN (%s) ORDER BY def_top_id" % ",".join("?"*len(targets)).replace("?", "%s"), targets)
print("\n\nTarget forms timeline:")
for r in cur.fetchall():
    print(f"  id={r[0]} ver={r[1]} regist={r[2]} update={r[3]} updated_at={r[4]}")

# Font analysis - check designer version distribution
print("\n\nAll distinct designer versions:")
cur.execute("SELECT designer_version, COUNT(*) as cnt FROM def_top WHERE designer_version IS NOT NULL GROUP BY designer_version ORDER BY designer_version")
for r in cur.fetchall():
    print(f"  v{r[0]}: {r[1]} forms")

conn.close()
