import pandas as pd
from datetime import datetime

class ExcelFileHandler:
    def __init__(self, sheet_name="", file_path=""):
        self.file_path = file_path
        self.sheet_name = sheet_name

    def datetime_iso_stamp(self):
        year = str(datetime.now().year)
        month = str(datetime.now().month) 
        day = str(datetime.now().day)
        hour = str(datetime.now().hour)
        minute = str(datetime.now().minute)
        second = str(datetime.now().second)

        if len(month) == 1: month = "0" + month
        if len(day) == 1: day = "0" + day
        if len(hour) == 1: hour = "0" + hour
        if len(minute) == 1: minute = "0" + minute
        if len(second) == 1: second = "0" + second
        return f"{year}-{month}-{day} {hour}.{minute}.{second}"

    def read_excel(self):
        """Reads an Excel file and returns a DataFrame."""
        try:
            self.file_path = self.file_path if self.file_path.endswith('.xlsx') else  f"{self.file_path}.xlsx"
            df = pd.read_excel(self.file_path, sheet_name=self.sheet_name)
            print(f"💚 Excel file read successfully from {self.file_path}.")
            return df
        except Exception as e:
            print(f"❌ Error reading Excel file: {e}")
            return None


    def write_excel(self, data, filename="ExcelFile", sheet_name='Sheet1', index=False):
        is_saved = False
        """Writes a DataFrame to an Excel file."""
        df: pd.DataFrame = pd.DataFrame()
        if data is not None:
            try:
                df = pd.DataFrame(data)
            except:
                df = data
            finally:
                filename = filename + "_" + self.datetime_iso_stamp()
                fullpath = self.file_path + filename if filename.endswith('.xlsx') else f"{self.file_path}{filename}.xlsx"
                with pd.ExcelWriter(fullpath, engine='openpyxl', mode='a' if pd.io.common.file_exists(fullpath) else 'w') as writer:
                    df.to_excel(writer, sheet_name=sheet_name, index=index, freeze_panes=(1, 0))
                print(f"💚 DataFrame written to {fullpath} successfully.")
                is_saved = True
        else:
            is_saved = False
            print("❌ No data to write to Excel.")
        
        return is_saved