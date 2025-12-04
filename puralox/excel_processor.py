import pandas as pd

class ExcelProcessor:
    def __init__(self, db_manager):
        self.db = db_manager

    @staticmethod
    def to_float(val):
        try:
            return float(val)
        except:
            return None

    @staticmethod
    def to_int(val):
        try:
            return int(float(val))
        except:
            return None

    def process_file(self, filepath):
        df = pd.read_excel(filepath, sheet_name='BET', header=None)

        # --- File Info ---
        file_info = {
            'file_name': str(df.iloc[1, 2]),
            'date_of_measurement': str(df.iloc[2, 2]),
            'time_of_measurement': str(df.iloc[3, 2]),
            'comment1': str(df.iloc[4, 2]),
            'comment2': str(df.iloc[5, 2]),
            'comment3': str(df.iloc[6, 2]),
            'comment4': str(df.iloc[7, 2]),
            'serial_number': str(df.iloc[8, 2]),
            'version': str(df.iloc[9, 2])
        }
        # insert into file_info and capture its new id
        fid = self.db.execute(
            '''INSERT INTO file_info 
               (file_name, date_of_measurement, time_of_measurement,
                comment1, comment2, comment3, comment4,
                serial_number, version)
               VALUES (:file_name, :date_of_measurement, :time_of_measurement,
                       :comment1, :comment2, :comment3, :comment4,
                       :serial_number, :version)''',
            file_info
        )

        # --- BET Parameters ---
        params = {
            'file_info_id': fid,
            'sample_weight':           self.to_float(df.iloc[11, 2]),
            'standard_volume':         self.to_float(df.iloc[12, 2]),
            'dead_volume':             self.to_float(df.iloc[13, 2]),
            'equilibrium_time':        self.to_float(df.iloc[14, 2]),
            'adsorptive':              str(df.iloc[15, 2]),
            'apparatus_temperature':   self.to_float(df.iloc[16, 2]),
            'adsorption_temperature':  self.to_float(df.iloc[17, 2]),
            'starting_point':          self.to_int(df.iloc[19, 3]),
            'end_point':               self.to_int(df.iloc[20, 3]),
            'slore':                   self.to_float(df.iloc[21, 3]),
            'intercept':               self.to_float(df.iloc[22, 3]),
            'correlation_coefficient': self.to_float(df.iloc[23, 3]),
            'vm':                      self.to_float(df.iloc[24, 3]),
            'as_bet':                  self.to_float(df.iloc[25, 3]),
            'c_value':                 self.to_float(df.iloc[26, 3]),
            'total_pore_volume':       self.to_float(df.iloc[27, 3]),
            'average_pore_diameter':   self.to_float(df.iloc[28, 3])
        }
        cols = ', '.join(params.keys())
        placeholders = ', '.join(':'+k for k in params)
        self.db.execute(
            f'INSERT INTO bet_parameters ({cols}) VALUES ({placeholders})',
            params
        )

        # --- Technical Info ---
        tech = {
            'file_info_id':                fid,
            'saturated_vapor_pressure':    self.to_float(df.iloc[11, 7]),
            'adsorption_cross_section':    self.to_float(df.iloc[12, 7]),
            'wall_adsorption_correction1': str(df.iloc[13, 4]),
            'wall_adsorption_correction2': str(df.iloc[14, 4]),
            'num_adsorption_points':       self.to_int(df.iloc[15, 4]),
            'num_desorption_points':       self.to_int(df.iloc[16, 4])
        }
        tcols = ', '.join(tech.keys())
        tph   = ', '.join(':'+k for k in tech)
        self.db.execute(
            f'INSERT INTO technical_info ({tcols}) VALUES ({tph})',
            tech
        )

        # --- Plot Columns ---
        header_row = df.iloc[30]
        for idx, col in header_row.items():
            if pd.notna(col):
                self.db.execute(
                    'INSERT INTO bet_plot_columns (file_info_id, col_index, col_name) VALUES (?, ?, ?)',
                    (fid, int(idx), str(col))
                )

        # --- Data Points ---
        for i in range(31, len(df)):
            row = df.iloc[i]
            no   = self.to_int(row[0])
            p_p0 = self.to_float(row[1])
            p_va = self.to_float(row[2])
            if None not in (no, p_p0, p_va):
                self.db.execute(
                    'INSERT INTO bet_data_points (file_info_id, no, p_p0, p_va_p0_p) VALUES (?, ?, ?, ?)',
                    (fid, no, p_p0, p_va)
                )

        # ─── return the new file_info ID ────────────────────────────────────
        return fid
