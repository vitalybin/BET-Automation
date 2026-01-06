import sqlite3
import logging

logger = logging.getLogger(__name__)


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
    # Fetch all rows as dicts
    # -----------------------------
    def fetchall_dict(self, sql, params=None):
        conn = self.connect()
        try:
            cur = conn.cursor()
            if params:
                cur.execute(sql, params)
            else:
                cur.execute(sql)
            rows = cur.fetchall()
            return [dict(r) for r in rows]
        except sqlite3.OperationalError as e:
            logger.exception("SQLite OperationalError in fetchall_dict: %s", e)
            raise
        except Exception as e:
            logger.exception("Database error in fetchall_dict: %s", e)
            raise
        finally:
            conn.close()

    # Single row
    def fetchone_dict(self, sql, params=None):
        conn = self.connect()
        try:
            cur = conn.cursor()
            if params:
                cur.execute(sql, params)
            else:
                cur.execute(sql)
            row = cur.fetchone()
            return dict(row) if row else None
        except sqlite3.OperationalError as e:
            logger.exception("SQLite OperationalError in fetchone_dict: %s", e)
            raise
        except Exception as e:
            logger.exception("Database error in fetchone_dict: %s", e)
            raise
        finally:
            conn.close()

    # Insert/update/delete
    def execute(self, sql, params=None):
        conn = self.connect()
        try:
            cur = conn.cursor()
            if params:
                cur.execute(sql, params)
            else:
                cur.execute(sql)
            conn.commit()
            last_id = cur.lastrowid
            return last_id
        except sqlite3.OperationalError as e:
            logger.exception("SQLite OperationalError in execute: %s", e)
            raise
        except Exception as e:
            logger.exception("Database error in execute: %s", e)
            raise
        finally:
            conn.close()
