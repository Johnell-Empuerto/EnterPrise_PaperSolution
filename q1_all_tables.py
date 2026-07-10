# -*- coding: utf-8 -*-
"""
Q1: Search every database table for PDF geometry evidence.
"""
import psycopg2
conn = psycopg2.connect(host="127.0.0.1", port=5432, dbname="irepodb", user="postgres", password="cimtops")
cur = conn.cursor()

# Get ALL table names
cur.execute("SELECT table_name FROM information_schema.tables WHERE table_schema='public' ORDER BY table_name")
tables = [r[0] for r in cur.fetchall()]

# Search for columns related to PDF, render, geometry, coordinates
keywords = ["pdf", "render", "geometry", "coordinate", "position", "page", "clip", "bbox", 
            "rectangle", "vector", "grid", "normalize", "cache", "temp", "log", "pixel", "image_dim", "dimension"]

print("TABLE COLUMN SEARCH")
print("-" * 80)

# Exclude def_cluster (known) and log/audit tables
for tbl in tables:
    cur.execute("SELECT column_name, data_type FROM information_schema.columns WHERE table_name=%s ORDER BY ordinal_position", (tbl,))
    cols = cur.fetchall()
    matches = []
    for cname, ctype in cols:
        cl = cname.lower()
        if any(k in cl for k in ["pdf", "render", "clip", "coord", "position", "normaliz",
                                  "cache", "temp", "vector", "bound", "rectangle", "pixel",
                                  "dimension", "image_w", "image_h", "image_size", "resolution"]):
            matches.append((cname, ctype))
    if matches:
        print(f"\n{tbl}:")
        for c, t in matches:
            print(f"  {c} ({t})")

# Look for tables with "pdf" in name
print("\n\n=== TABLES WITH 'pdf' IN NAME ===")
for tbl in tables:
    if "pdf" in tbl.lower():
        print(f"  {tbl}")

# Look for processing_log or temp tables
print("\n\n=== PROCESSING/TEMP/CACHE TABLES ===")
for tbl in tables:
    if any(k in tbl.lower() for k in ["temp", "cache", "log", "queue", "batch", "process", "background"]):
        print(f"  {tbl}")

# Check pg_catalog for any custom types or enums related to coordinate generation
cur.execute("SELECT typname, typcategory FROM pg_type WHERE typname LIKE '%coord%' OR typname LIKE '%pos%' OR typname LIKE '%pdf%'")
print("\n\n=== CUSTOM TYPES ===")
for r in cur.fetchall():
    print(f"  {r[0]} ({r[1]})")

conn.close()
print("\n\nDone.")
