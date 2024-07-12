import os
import win32com.client
import pandas as pd
import time

TABLE = input("TABLE NAME : ").upper()
print("\n")

# MDG = 'C:/Users/2095421/Downloads/Migration/MBEW_M1Q_Ext.xlsx'
# ATLAS = 'C:/Users/2095421/Downloads/Migration/MBEW_APB_Ext.xlsx'

MDG = input(" Please give MDG "+ TABLE + " file path ").strip('\"')
if os.path.isfile(MDG):
    print("The file exists.")
else:
    print("The file does not exist.........!")


ATLAS = input(" Please give ATLAS "+ TABLE + " file path ").strip('\"')
if os.path.isfile(ATLAS):
    print("The file exists.")
else:
    print("The file does not exist.........!")

start_time = time.time()

SCOPE = 'C:/Users/2095421/Downloads/Migration/GRD/Field_Scope_Data_Migration.xlsx'
DataValidationPath = 'C:/Users/2095421/Downloads/Migration/' + TABLE + '_DataValidation.xlsx'

df = pd.DataFrame()
df.to_excel(DataValidationPath)
scope_df = pd.read_excel(SCOPE, sheet_name=TABLE)
scope_df.to_excel(DataValidationPath, sheet_name='Sheet1', index=False)

print("\n")
print(TABLE + " Data Validation File created at - " + DataValidationPath )
print("\n")

def fun_cp(system_name, DataValidationPath, TABLE):
    excel_app = win32com.client.Dispatch("Excel.Application")

    # excel_app.DisplayAlerts = False #To stop the Alert in Excel

    source_wb = excel_app.Workbooks.Open(system_name)
    source_sheet = source_wb.Sheets('Sheet1')

    dest_wb = excel_app.Workbooks.Open(DataValidationPath)
    source_sheet.Copy(dest_wb.Sheets(1))    
    print(os.path.basename(system_name), " File copied to the DataValidation file Successfully!" )
    scope_sheet = dest_wb.Sheets('Sheet1')
    dest_wb.Sheets(1).Name=os.path.basename(system_name)
    dest_sheet = dest_wb.Sheets(os.path.basename(system_name))

    if TABLE != 'MARA':
        
        dest_sheet.Columns(1).Insert()
        dest_sheet.Cells(1, 1).Value = 'KEY_' + os.path.basename(system_name)
        dest_sheet.Range("A1").Interior.Color = 65535
        dest_sheet.Range("A1").Font.Bold = True
        print("Key Column created!")
        
        # All table consideration
        row = 2
        while dest_sheet.Cells(row, 2).Value is not None and dest_sheet.Cells(row, 3).Value is not None:
            # if TABLE == 'MVKE' or TABLE == 'MARD':
            if TABLE in ['MVKE', 'MARD']:
                dest_sheet.Cells(row, 1).Value = str(dest_sheet.Cells(row, 2).Value) + str(dest_sheet.Cells(row, 3).Value) + str(dest_sheet.Cells(row, 4).Value)
            else:
                dest_sheet.Cells(row, 1).Value = str(dest_sheet.Cells(row, 2).Value) + str(dest_sheet.Cells(row, 3).Value)
            row += 1
        print("Key creation done Successfully!")    

        print('Count of ', os.path.basename(system_name) + ": ", row-1)

    for row in range(2, scope_sheet.usedRange.Rows.Count+1):
        if scope_sheet.Cells(row, 2).Value == 'Scope':
            technical_name = scope_sheet.Cells(row, 1).Value
            for col in range(1, dest_sheet.UsedRange.Columns.Count+1):
                if dest_sheet.Cells(1, col).Value == technical_name:
                    dest_sheet.Cells(1, col).Interior.ColorIndex = 4
                    break

    print("Scope fields marked with Green Successfully!")

    dest_wb.Save()
    dest_wb.Close()
    source_wb.Close()
    excel_app.Quit()

    time.sleep(5)

    print("\n")


fun_cp(MDG, DataValidationPath, TABLE)
fun_cp(ATLAS, DataValidationPath, TABLE)



# df1 = pd.read_excel('C:/Users/2095421/Downloads/Migration/MBEW_DataValidation.xlsx', sheet_name='MBEW_M1Q_Ext.xlsx')
df1 = pd.read_excel(DataValidationPath, sheet_name=os.path.basename(MDG))
# df2 = pd.read_excel('C:/Users/2095421/Downloads/Migration/MBEW_DataValidation.xlsx', sheet_name='MBEW_APB_Ext.xlsx')
df2 = pd.read_excel(DataValidationPath, sheet_name=os.path.basename(ATLAS))

if TABLE == 'MARA':
    df1.set_index('MATNR', inplace=True)
    df2.set_index('MATNR', inplace=True)
else:
    df1.set_index('KEY_' + os.path.basename(MDG), inplace=True)
    df2.set_index('KEY_' + os.path.basename(ATLAS), inplace=True)

common_indices = df1.index.intersection(df2.index)

# df3 = pd.read_excel('C:/Users/2095421/Downloads/Migration/Field_Scope_Data_Migration.xlsx', sheet_name=TABLE)
df3 = pd.read_excel(DataValidationPath, sheet_name="Sheet1")
df_scope = df3[df3['Comment'] == 'Scope']
scope_field = df_scope['Technical Name']

# scope_field = ['/VSO/R_BOT_IND']

mismatched_dataframes = {}

for col_name in scope_field:
    
    # if col_name not in ['BWKEY', 'VKWEG'] :
    if col_name not in df1.columns:
        print(col_name, "--------------Not Found")
    else:
        print(col_name)
        # Compare the FIELD values of df1 and df2 for the common indices
        # comparison = df1.loc[common_indices, col_name] == df2.loc[common_indices, col_name]
        comparison = df1.loc[common_indices, col_name].eq(df2.loc[common_indices, col_name]) | (df1.loc[common_indices, col_name].isna() & df2.loc[common_indices, col_name].isna())
        # If they don't match, create a new dataframe with the mismatched indices
        if not comparison.all():
            mismatched_indices = common_indices[~comparison]

            if TABLE == 'MARA':
                df_mismatch = pd.DataFrame({
                'Matched': pd.Series(['no']*len(mismatched_indices), index=mismatched_indices),
                'MDG': df1.loc[mismatched_indices, col_name].squeeze(),
                'Atlas': df2.loc[mismatched_indices, col_name].squeeze(),
                })
            elif TABLE == 'MARC': 
                df_mismatch = pd.DataFrame({
                'Matched': pd.Series(['no']*len(mismatched_indices), index=mismatched_indices),
                'Material no': df2.loc[mismatched_indices, "MATNR"].squeeze(),
                'Plant': df2.loc[mismatched_indices, "WERKS"].squeeze(),
                'MDG': df1.loc[mismatched_indices, col_name].squeeze(),
                'Atlas': df2.loc[mismatched_indices, col_name].squeeze(),
                })
            elif TABLE == 'MBEW': 
                df_mismatch = pd.DataFrame({
                'Matched': pd.Series(['no']*len(mismatched_indices), index=mismatched_indices),
                'Material no': df2.loc[mismatched_indices, "MATNR"].squeeze(),
                'Plant': df2.loc[mismatched_indices, "BWKEY"].squeeze(),
                'MDG': df1.loc[mismatched_indices, col_name].squeeze(),
                'Atlas': df2.loc[mismatched_indices, col_name].squeeze(),
                })
            elif TABLE == 'MLGN': 
                df_mismatch = pd.DataFrame({
                'Matched': pd.Series(['no']*len(mismatched_indices), index=mismatched_indices),
                'Material no': df2.loc[mismatched_indices, "MATNR"].squeeze(),
                'LGNUM': df2.loc[mismatched_indices, "LGNUM"].squeeze(),
                'MDG': df1.loc[mismatched_indices, col_name].squeeze(),
                'Atlas': df2.loc[mismatched_indices, col_name].squeeze(),
                })
            elif TABLE == 'MLAN': 
                df_mismatch = pd.DataFrame({
                'Matched': pd.Series(['no']*len(mismatched_indices), index=mismatched_indices),
                'Material no': df2.loc[mismatched_indices, "MATNR"].squeeze(),
                'ALAND': df2.loc[mismatched_indices, "ALAND"].squeeze(),
                'MDG': df1.loc[mismatched_indices, col_name].squeeze(),
                'Atlas': df2.loc[mismatched_indices, col_name].squeeze(),
                })
            elif TABLE == 'MARD': 
                df_mismatch = pd.DataFrame({
                'Matched': pd.Series(['no']*len(mismatched_indices), index=mismatched_indices),
                'Material no': df2.loc[mismatched_indices, "MATNR"].squeeze(),
                'LGORT': df2.loc[mismatched_indices, "LGORT"].squeeze(),
                'Plant': df2.loc[mismatched_indices, "WERKS"].squeeze(),
                'MDG': df1.loc[mismatched_indices, col_name].squeeze(),
                'Atlas': df2.loc[mismatched_indices, col_name].squeeze(),
                })
            elif TABLE == 'MVKE': 
                df_mismatch = pd.DataFrame({
                'Matched': pd.Series(['no']*len(mismatched_indices), index=mismatched_indices),
                'Material no': df2.loc[mismatched_indices, "MATNR"].squeeze(),
                'VKORG': df2.loc[mismatched_indices, "VKORG"].squeeze(),
                'VKWEG': df2.loc[mismatched_indices, "VKWEG"].squeeze(),
                'MDG': df1.loc[mismatched_indices, col_name].squeeze(),
                'Atlas': df2.loc[mismatched_indices, col_name].squeeze(),
                })
            mismatched_dataframes[col_name] = df_mismatch



# Save all the mismatched dataframes to new sheets in the excel file
with pd.ExcelWriter(DataValidationPath, engine='openpyxl', mode='a') as writer:
    for sheet_name, df_mismatch in mismatched_dataframes.items():
        df_mismatch.to_excel(writer, sheet_name=sheet_name.replace("/",""))

# Print a success message
print("All mismatched dataframes have been written to the excel file.")

end_time = time.time()
execution_time = end_time - start_time
print(f"The program took {execution_time} seconds to execute.")