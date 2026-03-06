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
            if params is not None:
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
            if params is not None:
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
            if params is not None:
                cur.execute(sql, params)
            else:
                cur.execute(sql)
            conn.commit()
            return cur.lastrowid
        except sqlite3.OperationalError as e:
            logger.exception("SQLite OperationalError in execute: %s", e)
            raise
        except Exception as e:
            logger.exception("Database error in execute: %s", e)
            raise
        finally:
            conn.close()

    # -----------------------------
    # Bulk helpers (needed by PdfProcessor and others)
    # -----------------------------
    def execute_returning_id(self, sql, params=None):
        # execute() already returns lastrowid for INSERTs.
        return self.execute(sql, params)

    # -----------------------------
    # Schema helpers
    # -----------------------------
    def table_exists(self, name: str) -> bool:
        """Return True if a table with the given name exists in the database."""
        rows = self.fetchall_dict(
            "SELECT name FROM sqlite_master WHERE type='table' AND name=?",
            (name,)
        )
        return bool(rows)

    def executemany(self, sql, seq_of_params):
        conn = self.connect()
        try:
            cur = conn.cursor()
            cur.executemany(sql, seq_of_params)
            conn.commit()
            return cur.rowcount
        except sqlite3.OperationalError as e:
            logger.exception("SQLite OperationalError in executemany: %s", e)
            raise
        except Exception as e:
            logger.exception("Database error in executemany: %s", e)
            raise
        finally:
            conn.close()
