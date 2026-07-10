import psycopg2
conn = psycopg2.connect(host="127.0.0.1", port=5432, dbname="irepodb", user="postgres", password="cimtops")
cur = conn.cursor()
cur.execute("SELECT column_name, data_type FROM information_schema.columns WHERE table_name='def_top' AND data_type LIKE '%time%' OR data_type LIKE '%date%' OR column_name LIKE '%date%' OR column_name LIKE '%time%' OR column_name IN ('designer_version','updated_at','designer_display_version')")
for r in cur.fetchall():
    print(f"{r[0]} ({r[1]})")
conn.close()
