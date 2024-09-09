import os
import win32com.client
import pandas as pd
try:
    import dask.dataframe as dd
except ImportError:
    os.system('pip install dask[complete]')
    import dask.dataframe as dd
import time
print("Welcome to My Program ............!")
TABLE = input("TABLE NAME : ").upper()
print("\n")

ValidationFileReady = (input("Data validation file ready?   :  ")).upper()
if ValidationFileReady in ["Y", "YES"]:
    DataValidationPath = input("Please enter Data validation file Path  :  ").strip('\"')
    print("\n")
    start_time = time.time()
else:
    MDG = input("Please give MDG " + TABLE + " file path ").strip('\"')
    if os.path.isfile(MDG):
        print("The file exists.")
    else:
        print("The file does not exist.........!")

    ATLAS = input("Please give ATLAS " + TABLE + " file path ").strip('\"')
    if os.path.isfile(ATLAS):
        print("The file exists.")
    else:
        print("The file does not exist.........!")

    SCOPE = input("Field Scope File ------> ").strip('\"')
    DVP = input("Path to store the Datavalidation file  ------> ").strip('\"')
    DataValidationPath = DVP + '/' + TABLE + '_DataValidation.xlsx'

    start_time = time.time()

    df = dd.from_pandas(pd.DataFrame(), npartitions=1)
    df.to_excel(DataValidationPath)
    scope_df = dd.read_excel(SCOPE, sheet_name=TABLE)
    scope_df.to_excel(DataValidationPath, sheet_name='Sheet1', index=False)

    print("\n")
    print(TABLE + " Data Validation File created at - " + DataValidationPath)
    print("\n")

    def fun_cp(system_name, DataValidationPath, TABLE):
        excel_app = win32com.client.Dispatch("Excel.Application")
        source_wb = excel_app.Workbooks.Open(system_name)
        source_sheet = source_wb.Sheets('Sheet1')

        dest_wb = excel_app.Workbooks.Open(DataValidationPath)
        source_sheet.Copy(dest_wb.Sheets(1))
        print(os.path.basename(system_name), " File copied to the DataValidation file Successfully!")
        scope_sheet = dest_wb.Sheets('Sheet1')
        dest_wb.Sheets(1).Name = os.path.basename(system_name)
        dest_sheet = dest_wb.Sheets(os.path.basename(system_name))
        dest_wb.Save()
        dest_wb.Close()
        source_wb.Close()
        excel_app.Quit()
        time.sleep(5)

    fun_cp(MDG, DataValidationPath, TABLE)
    fun_cp(ATLAS, DataValidationPath, TABLE)

if ValidationFileReady in ["Y", "YES"]:
    MDG = input("MDG Sheet Name  :  ")
    ATLAS = input("ATLAS Sheet Name  :  ")
    df1 = dd.read_excel(DataValidationPath, sheet_name=MDG, index_col='KEY_' + MDG)
    df2 = dd.read_excel(DataValidationPath, sheet_name=ATLAS, index_col='KEY_' + ATLAS)
else:
    df1 = dd.read_excel(DataValidationPath, sheet_name=os.path.basename(MDG))
    df1 = df1.assign(KEY_=df1['MATNR'].astype(str) + df1['WERKS'].astype(str) + df1['LGORT'].astype(str))
    with pd.ExcelWriter(DataValidationPath, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df1.to_excel(writer, sheet_name=os.path.basename(MDG), index=False)

    df2 = dd.read_excel(DataValidationPath, sheet_name=os.path.basename(ATLAS))
    df2 = df2.assign(KEY_=df2['MATNR'].astype(str) + df2['WERKS'].astype(str) + df2['LGORT'].astype(str))
    with pd.ExcelWriter(DataValidationPath, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df2.to_excel(writer, sheet_name=os.path.basename(ATLAS), index=False)

    if TABLE == 'MARA':
        df1 = df1.set_index('MATNR')
        df2 = df2.set_index('MATNR')
    else:
        df1 = df1.set_index('KEY_' + os.path.basename(MDG))
        df2 = df2.set_index('KEY_' + os.path.basename(ATLAS))

common_indices = df2.index.intersection(df1.index)

df3 = dd.read_excel(DataValidationPath, sheet_name="Sheet1")
df_scope = df3[df3['Comment'] == 'Scope']
scope_field = df_scope['Technical Name']

mismatched_dataframes = {}

for col_name in scope_field:
    if col_name not in df1.columns:
        print(col_name, "--------------Not Found")
    else:
        print(col_name)
        comparison = (df2.loc[common_indices, col_name].isna() & df1.loc[common_indices, col_name].isna()) | df2.loc[common_indices, col_name].eq(df1.loc[common_indices, col_name])

        if not comparison.all():
            mismatched_indices = common_indices[~comparison]

            if TABLE == 'MARA':
                df_mismatch = dd.from_pandas(pd.DataFrame({
                    'Matched': pd.Series(['no']*len(mismatched_indices), index=mismatched_indices),
                    'MDG': df1.loc[mismatched_indices, col_name].squeeze(),
                    'Atlas': df2.loc[mismatched_indices, col_name].squeeze(),
                }), npartitions=1)
            elif TABLE == 'MARC':
                df_mismatch = dd.from_pandas(pd.DataFrame({
                    'Matched': pd.Series(['no']*len(mismatched_indices), index=mismatched_indices),
                    'Material no': df2.loc[mismatched_indices, "MATNR"].squeeze(),
                    'Plant': df2.loc[mismatched_indices, "WERKS"].squeeze(),
                    'MDG': df1.loc[mismatched_indices, col_name].squeeze(),
                    'Atlas': df2.loc[mismatched_indices, col_name].squeeze(),
                }), npartitions=1)
            elif TABLE == 'MBEW':
                df_mismatch = dd.from_pandas(pd.DataFrame({
                    'Matched': pd.Series(['no']*len(mismatched_indices), index=mismatched_indices),
                    'Material no': df2.loc[mismatched_indices, "MATNR"].squeeze(),
                    'Plant': df2.loc[mismatched_indices, "BWKEY"].squeeze(),
                    'MDG': df1.loc[mismatched_indices, col_name].squeeze(),
                    'Atlas': df2.loc[mismatched_indices, col_name].squeeze(),
                }), npartitions=1)
            elif TABLE == 'MLGN':
                df_mismatch = dd.from_pandas(pd.DataFrame({
                    'Matched': pd.Series(['no']*len(mismatched_indices), index=mismatched_indices),
                    'Material no': df2.loc[mismatched_indices, "MATNR"].squeeze(),
                    'LGNUM': df2.loc[mismatched_indices, "LGNUM"].squeeze(),
                    'MDG': df1.loc[mismatched_indices, col_name].squeeze(),
                    'Atlas': df2.loc[mismatched_indices, col_name].squeeze(),
                }), npartitions=1)
            elif TABLE == 'MLAN':
                df_mismatch = dd.from_pandas(pd.DataFrame({
                    'Matched': pd.Series(['no']*len(mismatched_indices), index=mismatched_indices),
                    'Material no': df2.loc[mismatched_indices, "MATNR"].squeeze(),
                    'ALAND': df2.loc[mismatched_indices, "ALAND"].squeeze(),
                    'MDG': df1.loc[mismatched_indices, col_name].squeeze(),
                    'Atlas': df2.loc[mismatched_indices, col_name].squeeze(),
                }), npartitions=1)
            elif TABLE == 'MARD':
                df_mismatch = dd.from_pandas(pd.DataFrame({
                    'Matched': pd.Series(['no']*len(mismatched_indices), index=mismatched_indices),
                    'Material no': df2.loc[mismatched_indices, "MATNR"].squeeze(),
                    'LGORT': df2.loc[mismatched_indices, "LGORT"].squeeze(),
                    'Plant': df2.loc[mismatched_indices, "WERKS"].squeeze(),
                    'MDG': df1.loc[mismatched_indices, col_name].squeeze(),
                    'Atlas': df2.loc[mismatched_indices, col_name].squeeze(),
                }), npartitions=1)
            elif TABLE == 'MVKE':
                df_mismatch = dd.from_pandas(pd.DataFrame({
                    'Matched': pd.Series(['no']*len(mismatched_indices), index=mismatched_indices),
                    'Material no': df2.loc[mismatched_indices, "MATNR"].squeeze(),
                    'VKORG': df2.loc[mismatched_indices, "VKORG"].squeeze(),
                    'VKWEG': df2.loc[mismatched_indices, "VKWEG"].squeeze(),
                    'MDG': df1.loc[mismatched_indices, col_name].squeeze(),
                    'Atlas': df2.loc[mismatched_indices, col_name].squeeze(),
                }), npartitions=1)
            mismatched_dataframes[col_name] = df_mismatch
