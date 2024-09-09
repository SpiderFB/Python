import os
import win32com.client
try:
    import pandas as pd
except ImportError:
    os.system('pip install pandas')
    import pandas as pd
import time
print("Welcome to My Program ............!")
TABLE = input("TABLE NAME : ").upper()
print("\n")

ValidationFileReady = (input("Data validation file ready?   :  ")).upper()
if ValidationFileReady in ["Y", "YES"]:
    DataValidationPath = input("Please enter Data validation file Path  :  ").strip('\"')
    # SCOPE = input("Field Scope File ------> ").strip('\"')
    print("\n")
    start_time = time.time()

    # df = pd.DataFrame()
    # df.to_excel(DataValidationPath)
    # scope_df = pd.read_excel(SCOPE, sheet_name=TABLE)
    # scope_df.to_excel(DataValidationPath, sheet_name='Sheet1', index=False)

else:

    # MDG = 'C:/Users/2095421/Downloads/Migration/MBEW_M1Q_Ext.xlsx'
    # ATLAS = 'C:/Users/2095421/Downloads/Migration/MBEW_APB_Ext.xlsx'

    MDG = input("Please give MDG "+ TABLE + " file path ").strip('\"')
    if os.path.isfile(MDG):
        print("The file exists.")
    else:
        print("The file does not exist.........!")


    ATLAS = input("Please give ATLAS "+ TABLE + " file path ").strip('\"')
    if os.path.isfile(ATLAS):
        print("The file exists.")
    else:
        print("The file does not exist.........!")

    SCOPE = input("Field Scope File ------> ").strip('\"')
    # SCOPE = 'C:/Users/2095421/Downloads/Migration/GRD/Field_Scope_Data_Migration.xlsx'

    DVP = input("Path to store the Datavalidation file  ------> ").strip('\"')
    DataValidationPath = DVP + '/' + TABLE + '_DataValidation.xlsx'


    start_time = time.time()

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
        dest_wb.Save()
        dest_wb.Close()
        source_wb.Close()
        excel_app.Quit()
        time.sleep(5)

        # if TABLE != 'MARA':
            
        #     dest_sheet.Columns(1).Insert()
        #     dest_sheet.Cells(1, 1).Value = 'KEY_' + os.path.basename(system_name)
        #     dest_sheet.Range("A1").Interior.Color = 65535
        #     dest_sheet.Range("A1").Font.Bold = True
        #     print("Key Column created!")
            
        #     # All table consideration
        #     row = 2
        #     while dest_sheet.Cells(row, 2).Value is not None and dest_sheet.Cells(row, 3).Value is not None:
        #         # if TABLE == 'MVKE' or TABLE == 'MARD':
        #         if TABLE in ['MVKE', 'MARD']:
        #             dest_sheet.Cells(row, 1).Value = str(dest_sheet.Cells(row, 2).Value) + str(dest_sheet.Cells(row, 3).Value) + str(dest_sheet.Cells(row, 4).Value)
        #         else:
        #             dest_sheet.Cells(row, 1).Value = str(dest_sheet.Cells(row, 2).Value) + str(dest_sheet.Cells(row, 3).Value)
        #         row += 1
        #     print("Key creation done Successfully!")    

        #     print('Count of ', os.path.basename(system_name) + ": ", row-1)
            
        # print("\n")


    fun_cp(MDG, DataValidationPath, TABLE)
    fun_cp(ATLAS, DataValidationPath, TABLE)



# df1 = pd.read_excel('C:/Users/2095421/Downloads/Migration/MBEW_DataValidation.xlsx', sheet_name='MBEW_M1Q_Ext.xlsx')
# df2 = pd.read_excel('C:/Users/2095421/Downloads/Migration/MBEW_DataValidation.xlsx', sheet_name='MBEW_APB_Ext.xlsx')

if ValidationFileReady in ["Y", "YES"]:
    MDG = input("MDG Sheet Name  :  ")
    ATLAS = input("ATLAS Sheet Name  :  ")
    df1 = pd.read_excel(DataValidationPath, sheet_name = MDG, index_col='KEY_' + MDG) #Except MARA
    df2 = pd.read_excel(DataValidationPath, sheet_name = ATLAS, index_col='KEY_' + ATLAS) #Except MARA
else:
    df1 = pd.read_excel(DataValidationPath, sheet_name=os.path.basename(MDG))
    df1.insert(0, 'KEY_' + os.path.basename(MDG), df1['MATNR'].astype(str) + df1['WERKS'].astype(str) + df1['LGORT'].astype(str))
    # df1.to_excel(DataValidationPath,sheet_name=os.path.basename(MDG), index=False)
    with pd.ExcelWriter(DataValidationPath, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df1.to_excel(writer, sheet_name=os.path.basename(MDG), index=False)


    # df1['KEY_' + os.path.basename(MDG)] = df1['MATNR'].astype(str) + df['WERKS'].astype(str) + df['LGORT'].astype(str)
    # # Move the new column to the first position
    # cols = df.columns.tolist()
    # cols.insert(0, cols.pop(cols.index('new_column')))
    # df = df[cols]


    df2 = pd.read_excel(DataValidationPath, sheet_name=os.path.basename(ATLAS))
    df2.insert(0, 'KEY_' + os.path.basename(ATLAS), df2['MATNR'].astype(str) + df2['WERKS'].astype(str) + df2['LGORT'].astype(str))
    # df2.to_excel(DataValidationPath,sheet_name=os.path.basename(ATLAS), index=False)
    with pd.ExcelWriter(DataValidationPath, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df2.to_excel(writer, sheet_name=os.path.basename(ATLAS), index=False)


    if TABLE == 'MARA':
        df1.set_index('MATNR', inplace=True)
        df2.set_index('MATNR', inplace=True)
    else:
        df1.set_index('KEY_' + os.path.basename(MDG), inplace=True)
        df2.set_index('KEY_' + os.path.basename(ATLAS), inplace=True)

common_indices = df2.index.intersection(df1.index)
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

        # del df1
        # del df2

        # df1 = pd.read_excel(DataValidationPath, sheet_name=os.path.basename(MDG), usecols=['KEY_' + os.path.basename(MDG), col_name ])
        # df1.set_index('KEY_' + os.path.basename(MDG), inplace=True)
            
        # df2 = pd.read_excel(DataValidationPath, sheet_name=os.path.basename(ATLAS), usecols=['KEY_' + os.path.basename(ATLAS), col_name])
        # df2.set_index('KEY_' + os.path.basename(ATLAS), inplace=True)

        # Compare the FIELD values of df1 and df2 for the common indices
        comparison = (df2.loc[common_indices, col_name].isna() & df1.loc[common_indices, col_name].isna()) | df2.loc[common_indices, col_name].eq(df1.loc[common_indices, col_name])
        
        # print(f"Size of common_indices: {len(common_indices)}")
        # print(f"Size of comparison: {len(comparison)}")


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
    with pd.ExcelWriter(DataValidationPath, engine='openpyxl', mode='a') as writer:
        for sheet_name, df_mismatch in mismatched_dataframes.items():
            df_mismatch.to_excel(writer, sheet_name=sheet_name.replace("/",""))

print("Validation Done, all saved           ^_^             ")

# Save all the mismatched dataframes to new sheets in the excel file
# with pd.ExcelWriter(DataValidationPath, engine='openpyxl', mode='a') as writer:
#     for sheet_name, df_mismatch in mismatched_dataframes.items():
#         df_mismatch.to_excel(writer, sheet_name=sheet_name.replace("/",""))

# Print a success message
print("All mismatched dataframes have been written to the excel file.")

end_time = time.time()
execution_time = end_time - start_time
print(f"The program took {execution_time} seconds to execute.")