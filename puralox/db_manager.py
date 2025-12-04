import sqlite3

class DatabaseManager:

    def __init__(self, db_path):
        self.db_path = db_path
        self._ensure()

    def _ensure(self):
        conn = sqlite3.connect(self.db_path)
        conn.execute("PRAGMA foreign_keys = ON")
        conn.close()

    def connect(self):
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        return conn

    # -----------------------------
    # FIXED: Now supports parameters
    # -----------------------------
    def fetchall_dict(self, sql, params=None):
        conn = self.connect()
        cur = conn.cursor()
        if params:
            cur.execute(sql, params)
        else:
            cur.execute(sql)
        rows = cur.fetchall()
        conn.close()
        return [dict(r) for r in rows]

    # Same fix for fetchone if needed later
    def fetchone_dict(self, sql, params=None):
        conn = self.connect()
        cur = conn.cursor()
        if params:
            cur.execute(sql, params)
        else:
            cur.execute(sql)
        row = cur.fetchone()
        conn.close()
        return dict(row) if row else None

    # Insert/update/delete
    def execute(self, sql, params=None):
        conn = self.connect()
        cur = conn.cursor()
        if params:
            cur.execute(sql, params)
        else:
            cur.execute(sql)
        conn.commit()
        last_id = cur.lastrowid
        conn.close()
        return last_id
