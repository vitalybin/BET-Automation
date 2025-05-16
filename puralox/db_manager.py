import sqlite3

class DatabaseManager:
    def __init__(self, db_name):
        self.db_name = db_name
        self.create_db()

    def create_db(self):
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS file_info (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_name TEXT, date_of_measurement TEXT, time_of_measurement TEXT,
                comment1 TEXT, comment2 TEXT, comment3 TEXT, comment4 TEXT,
                serial_number TEXT, version TEXT
            )
        ''')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS bet_parameters (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_info_id INTEGER, sample_weight REAL, standard_volume REAL,
                dead_volume REAL, equilibrium_time REAL, adsorptive TEXT,
                apparatus_temperature REAL, adsorption_temperature REAL,
                starting_point INTEGER, end_point INTEGER, slore REAL,
                intercept REAL, correlation_coefficient REAL, vm REAL,
                as_bet REAL, c_value REAL, total_pore_volume REAL,
                average_pore_diameter REAL,
                FOREIGN KEY (file_info_id) REFERENCES file_info(id)
            )
        ''')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS bet_plot_columns (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_info_id INTEGER, col_index INTEGER, col_name TEXT,
                FOREIGN KEY (file_info_id) REFERENCES file_info(id)
            )
        ''')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS bet_data_points (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_info_id INTEGER, no INTEGER, p_p0 REAL, p_va_p0_p REAL,
                FOREIGN KEY (file_info_id) REFERENCES file_info(id)
            )
        ''')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS technical_info (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_info_id INTEGER, saturated_vapor_pressure REAL,
                adsorption_cross_section REAL, wall_adsorption_correction1 TEXT,
                wall_adsorption_correction2 TEXT, num_adsorption_points INTEGER,
                num_desorption_points INTEGER,
                FOREIGN KEY (file_info_id) REFERENCES file_info(id)
            )
        ''')
        conn.commit()
        conn.close()

    def execute(self, query, params=None):
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()
        if params:
            cursor.execute(query, params)
        else:
            cursor.execute(query)
        conn.commit()
        row_id = cursor.lastrowid
        conn.close()
        return row_id

    def fetchall_dict(self, query):
        conn = sqlite3.connect(self.db_name)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        cursor.execute(query)
        results = [dict(row) for row in cursor.fetchall()]
        conn.close()
        return results
