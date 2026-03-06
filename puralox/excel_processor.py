import pandas as pd
from abc import ABC, abstractmethod
from typing import Optional

from .nomenclature import build_measurement_id


class BaseImporter(ABC):
    """
    Abstract importer contract so Excel + PDF share a single interface.
    This is what gives you real UML inheritance (generalization).
    """

    @abstractmethod
    def import_file(self, path: str, original_filename: Optional[str] = None) -> int:
        """Import file into DB and return file_info_id."""
        raise NotImplementedError


class ExcelProcessor(BaseImporter):
    def __init__(self, db_manager):
        self.db = db_manager

    # Polymorphic entrypoint (keeps existing behavior)
    def import_file(self, path: str, original_filename: Optional[str] = None) -> int:
        return self.process_file(path)

    @staticmethod
    def to_float(val):
        try:
            return float(val)
        except Exception:
            return None

    @staticmethod
    def to_int(val):
        try:
            return int(float(val))
        except Exception:
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
            'version': str(df.iloc[9, 2]),
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

        # --- auto-generate Measurement ID and store in comment5 ---
        measurement_id = build_measurement_id(
            file_id=fid,
            file_name=file_info['file_name'],
            date_of_measurement=file_info['date_of_measurement'],
            time_of_measurement=file_info['time_of_measurement'],
            operator=file_info['comment2'],
            instrument=file_info['comment4'],
            serial_number=file_info['serial_number'],
            comment1=file_info['comment1'],
            comment3=file_info['comment3'],
        )
        self.db.execute(
            "UPDATE file_info SET comment5=? WHERE id=?",
            (measurement_id, fid)
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
            'starting_point':          self.to_int(df.iloc[18, 2]),
            'end_point':               self.to_int(df.iloc[19, 2]),
            'slore':                   self.to_float(df.iloc[20, 2]),
            'intercept':               self.to_float(df.iloc[21, 2]),
            'correlation_coefficient': self.to_float(df.iloc[22, 2]),
            'vm':                      self.to_float(df.iloc[23, 2]),
            'as_bet':                  self.to_float(df.iloc[24, 2]),
            'c_value':                 self.to_float(df.iloc[25, 2]),
            'total_pore_volume':       self.to_float(df.iloc[26, 2]),
            'average_pore_diameter':   self.to_float(df.iloc[27, 2]),
        }
        self.db.execute(
            '''INSERT INTO bet_parameters
               (file_info_id, sample_weight, standard_volume, dead_volume,
                equilibrium_time, adsorptive, apparatus_temperature,
                adsorption_temperature, starting_point, end_point, slore,
                intercept, correlation_coefficient, vm, as_bet, c_value,
                total_pore_volume, average_pore_diameter)
               VALUES (:file_info_id, :sample_weight, :standard_volume, :dead_volume,
                       :equilibrium_time, :adsorptive, :apparatus_temperature,
                       :adsorption_temperature, :starting_point, :end_point, :slore,
                       :intercept, :correlation_coefficient, :vm, :as_bet, :c_value,
                       :total_pore_volume, :average_pore_diameter)''',
            params
        )

        # --- Technical Info ---
        tech = {
            'file_info_id': fid,
            'saturated_vapor_pressure':     self.to_float(df.iloc[29, 2]),
            'adsorption_cross_section':     self.to_float(df.iloc[30, 2]),
            'wall_adsorption_correction1':  str(df.iloc[31, 2]),
            'wall_adsorption_correction2':  str(df.iloc[32, 2]),
            'num_adsorption_points':        self.to_int(df.iloc[33, 2]),
            'num_desorption_points':        self.to_int(df.iloc[34, 2]),
            'mass':                         self.to_float(df.iloc[35, 2]) if df.shape[0] > 35 else None,
            'internal_device_id':           str(df.iloc[36, 2]) if df.shape[0] > 36 else None,
        }
        self.db.execute(
            '''INSERT INTO technical_info
               (file_info_id, saturated_vapor_pressure, adsorption_cross_section,
                wall_adsorption_correction1, wall_adsorption_correction2,
                num_adsorption_points, num_desorption_points, mass, internal_device_id)
               VALUES (:file_info_id, :saturated_vapor_pressure, :adsorption_cross_section,
                       :wall_adsorption_correction1, :wall_adsorption_correction2,
                       :num_adsorption_points, :num_desorption_points, :mass, :internal_device_id)''',
            tech
        )

        # --- Plot Columns ---
        for col_index in range(1, 5):
            col_name = str(df.iloc[38, col_index]) if df.shape[1] > col_index else None
            if col_name and col_name.lower() != "nan":
                self.db.execute(
                    "INSERT INTO bet_plot_columns (file_info_id, col_index, col_name) VALUES (?, ?, ?)",
                    (fid, col_index, col_name)
                )

        # --- Data Points ---
        start_row = 39
        no = 1
        while start_row + no < df.shape[0]:
            p_p0 = self.to_float(df.iloc[start_row + no, 1]) if df.shape[1] > 1 else None
            p_va = self.to_float(df.iloc[start_row + no, 2]) if df.shape[1] > 2 else None
            if p_p0 is None and p_va is None:
                break
            self.db.execute(
                "INSERT INTO bet_data_points (file_info_id, no, p_p0, p_va_p0_p) VALUES (?, ?, ?, ?)",
                (fid, no, p_p0, p_va)
            )
            no += 1

        return fid
