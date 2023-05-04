import pandas as pd
import time


class ExcelSheetAutomator:
    def __init__(self):
        self.path = 'Ama_table__lts.json'
        self.file_name = 'lts_ama_table.xlsx'

    def json_reader(self):
        """ Function used for reading the json files into a dataframe"""
        df = pd.read_json(self.path)
        # to find the columns that have 'unnamed', then drop those columns.
        df.drop(df.columns[df.columns.str.contains('unnamed', case=False)], axis=1, inplace=True)
        return df

    @staticmethod
    def decorator_function(original_function):
        start_time = time.time()

        def wrapper_function(*args, **kwargs):
            original_function(*args, **kwargs)
            end_time = time.time()
            print(f'total time taken by {original_function.__name__} is {end_time - start_time} seconds')

        return wrapper_function
