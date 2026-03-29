import sqlite3

with sqlite3.connect("job_audit.db") as conn:
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    cur.execute("SELECT * FROM audit_log ORDER BY job_id ASC")

    for row in cur:
        print(dict(row),"\n")













'''
import time
while True:
    with sqlite3.connect("audit.db") as conn:
        cur = conn.cursor()
        cur.execute("""
            SELECT *
            FROM audit_log
            ORDER BY job_id ASC
        """)

        columns = [c[0] for c in cur.description]


        print(cur.fetchone())
        time.sleep(0.1)
       '''