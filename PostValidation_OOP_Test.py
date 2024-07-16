import os
import win32com.client
import pandas as pd
import time

class DataValidation:
    def __init__(self, table, mdg_path, atlas_path, scope_path, dvp_path):
        self.table = table.upper()
        self.mdg_path = mdg_path
        self.atlas_path = atlas_path
        self.scope_path = scope_path
        self.dv_path = os.path.join(dvp_path, f"{self.table}_DataValidation.xlsx")
        self.start_time = time.time()
        self.mismatched_dataframes = {}

    def check_file_exists(self, file_path):
        if os.path.isfile(file_path):
            print(f"The file {file_path} exists.")
        else:
            print(f"The file {file_path} does not exist.........!")

    def create_validation_file(self):
        df = pd.DataFrame()
        df.to_excel(self.dv_path)
        scope_df = pd.read_excel(self.scope_path, sheet_name=self.table)
        scope_df.to_excel(self.dv_path, sheet_name='Sheet1', index=False)
        print(f"{self.table} Data Validation File created at - {self.dv_path}")

    def copy_sheet(self, system_name):
        excel_app = win32com.client.Dispatch("Excel.Application")
        try:
            source_wb = excel_app.Workbooks.Open(system_name)
            source_sheet = source_wb.Sheets('Sheet1')
            dest_wb = excel_app.Workbooks.Open(self.dv_path)
            source_sheet.Copy(dest_wb.Sheets(1))
            print(f"{os.path.basename(system_name)} File copied to the DataValidation file Successfully!")
            scope_sheet = dest_wb.Sheets('Sheet1')
            dest_wb.Sheets(1).Name = os.path.basename(system_name)
            dest_sheet = dest_wb.Sheets(os.path.basename(system_name))

            if self.table != 'MARA':
                dest_sheet.Columns(1).Insert()
                dest_sheet.Cells(1, 1).Value = 'KEY_' + os.path.basename(system_name)
                dest_sheet.Range("A1").Interior.Color = 65535
                dest_sheet.Range("A1").Font.Bold = True
                print("Key Column created!")

                row = 2
                while dest_sheet.Cells(row, 2).Value is not None and dest_sheet.Cells(row, 3).Value is not None:
                    if self.table in ['MVKE', 'MARD']:
                        dest_sheet.Cells(row, 1).Value = str(dest_sheet.Cells(row, 2).Value) + str(dest_sheet.Cells(row, 3).Value) + str(dest_sheet.Cells(row, 4).Value)
                    else:
                        dest_sheet.Cells(row, 1).Value = str(dest_sheet.Cells(row, 2).Value) + str(dest_sheet.Cells(row, 3).Value)
                    row += 1
                print("Key creation done Successfully!")
                print(f'Count of {os.path.basename(system_name)}: {row-1}')

            print("<-----Starting to mark Scope Fields------>")
            for row in range(2, scope_sheet.usedRange.Rows.Count + 1):
                if scope_sheet.Cells(row, 2).Value == 'Scope':
                    technical_name = scope_sheet.Cells(row, 1).Value
                    for col in range(1, dest_sheet.UsedRange.Columns.Count + 1):
                        if dest_sheet.Cells(1, col).Value == technical_name:
                            dest_sheet.Cells(1, col).Interior.ColorIndex = 4
                            break

            print("Scope fields marked with Green Successfully!")
            dest_wb.Save()
            dest_wb.Close()
            source_wb.Close()
        except Exception as e:
            print(f"An error occurred: {e}")
        finally:
            excel_app.Quit()
            time.sleep(5)

    def validate_data(self):
        df1 = pd.read_excel(self.dv_path, sheet_name=os.path.basename(self.mdg_path))
        df2 = pd.read_excel(self.dv_path, sheet_name=os.path.basename(self.atlas_path))

        if self.table == 'MARA':
            df1.set_index('MATNR', inplace=True)
            df2.set_index('MATNR', inplace=True)
        else:
            df1.set_index('KEY_' + os.path.basename(self.mdg_path), inplace=True)
            df2.set_index('KEY_' + os.path.basename(self.atlas_path), inplace=True)

        common_indices = df1.index.intersection(df2.index)
        df3 = pd.read_excel(self.dv_path, sheet_name="Sheet1")
        df_scope = df3[df3['Comment'] == 'Scope']
        scope_field = df_scope['Technical Name']

        for col_name in scope_field:
            if col_name not in df1.columns:
                print(f"{col_name} --------------Not Found")
            else:
                print(col_name)
                comparison = df1.loc[common_indices, col_name].eq(df2.loc[common_indices, col_name]) | (df1.loc[common_indices, col_name].isna() & df2.loc[common_indices, col_name].isna())
                if not comparison.all():
                    mismatched_indices = common_indices[~comparison]
                    df_mismatch = self.create_mismatch_dataframe(df1, df2, mismatched_indices, col_name)
                    self.mismatched_dataframes[col_name] = df_mismatch

        print("Validation Done, Saving all mismatch in the Excel file kindly wait..............")
        self.save_mismatched_dataframes()
        end_time = time.time()
        execution_time = end_time - self.start_time
        print(f"The program took {execution_time} seconds to execute.")

    def create_mismatch_dataframe(self, df1, df2, mismatched_indices, col_name):
        if self.table == 'MARA':
            return pd.DataFrame({
                'Matched': pd.Series(['no'] * len(mismatched_indices), index=mismatched_indices),
                'MDG': df1.loc[mismatched_indices, col_name].squeeze(),
                'Atlas': df2.loc[mismatched_indices, col_name].squeeze(),
            })
        elif self.table == 'MARC':
            return pd.DataFrame({
                'Matched': pd.Series(['no'] * len(mismatched_indices), index=mismatched_indices),
                'Material no': df2.loc[mismatched_indices, "MATNR"].squeeze(),
                'Plant': df2.loc[mismatched_indices, "WERKS"].squeeze(),
                'MDG': df1.loc[mismatched_indices, col_name].squeeze(),
                'Atlas': df2.loc[mismatched_indices, col_name].squeeze(),
            })
        elif self.table == 'MBEW': 
                df_mismatch = pd.DataFrame({
                'Matched': pd.Series(['no']*len(mismatched_indices), index=mismatched_indices),
                'Material no': df2.loc[mismatched_indices, "MATNR"].squeeze(),
                'Plant': df2.loc[mismatched_indices, "BWKEY"].squeeze(),
                'MDG': df1.loc[mismatched_indices, col_name].squeeze(),
                'Atlas': df2.loc[mismatched_indices, col_name].squeeze(),
                })
        elif self.table == 'MLGN': 
            df_mismatch = pd.DataFrame({
            'Matched': pd.Series(['no']*len(mismatched_indices), index=mismatched_indices),
            'Material no': df2.loc[mismatched_indices, "MATNR"].squeeze(),
            'LGNUM': df2.loc[mismatched_indices, "LGNUM"].squeeze(),
            'MDG': df1.loc[mismatched_indices, col_name].squeeze(),
            'Atlas': df2.loc[mismatched_indices, col_name].squeeze(),
            })
        elif self.table == 'MLAN': 
            df_mismatch = pd.DataFrame({
            'Matched': pd.Series(['no']*len(mismatched_indices), index=mismatched_indices),
            'Material no': df2.loc[mismatched_indices, "MATNR"].squeeze(),
            'ALAND': df2.loc[mismatched_indices, "ALAND"].squeeze(),
            'MDG': df1.loc[mismatched_indices, col_name].squeeze(),
            'Atlas': df2.loc[mismatched_indices, col_name].squeeze(),
            })
        elif self.table == 'MARD': 
            df_mismatch = pd.DataFrame({
            'Matched': pd.Series(['no']*len(mismatched_indices), index=mismatched_indices),
            'Material no': df2.loc[mismatched_indices, "MATNR"].squeeze(),
            'LGORT': df2.loc[mismatched_indices, "LGORT"].squeeze(),
            'Plant': df2.loc[mismatched_indices, "WERKS"].squeeze(),
            'MDG': df1.loc[mismatched_indices, col_name].squeeze(),
            'Atlas': df2.loc[mismatched_indices, col_name].squeeze(),
            })
        elif self.table == 'MVKE': 
            df_mismatch = pd.DataFrame({
            'Matched': pd.Series(['no']*len(mismatched_indices), index=mismatched_indices),
            'Material no': df2.loc[mismatched_indices, "MATNR"].squeeze(),
            'VKORG': df2.loc[mismatched_indices, "VKORG"].squeeze(),
            'VKWEG': df2.loc[mismatched_indices, "VKWEG"].squeeze(),
            'MDG': df1.loc[mismatched_indices, col_name].squeeze(),
            'Atlas': df2.loc[mismatched_indices, col_name].squeeze(),
            })

        # Add other table-specific conditions here...

    def save_mismatched_dataframes(self):
        with pd.ExcelWriter(self.dv_path, engine='openpyxl', mode='a') as writer:
            for sheet_name, df_mismatch in self.mismatched_dataframes.items():
                df_mismatch.to_excel(writer, sheet_name=sheet_name.replace("/", ""))

if __name__ == "__main__":
    table = input("TABLE NAME: ").upper()
    mdg_path = input(f"Please give MDG {table} file path: ").strip('"')
    atlas_path = input(f"Please give ATLAS {table} file path: ").strip('"')
    scope_path = input("Field Scope File ------> ").strip('"')
    dvp_path = input("Path to store the DataValidation file ------> ").strip('"')

    validator = DataValidation(table, mdg_path, atlas_path, scope_path, dvp_path)
    validator.check_file_exists(mdg_path)
    validator.check_file_exists(atlas_path)
    validator.create_validation_file()
    validator.copy_sheet(mdg_path)
    validator.copy_sheet(atlas_path)
    validator.validate_data()
