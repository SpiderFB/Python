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

CLUSTER = input("Cluster Name  :  ")

# ValidationFileReady = (input("Data validation file ready?   :  ")).upper()
ValidationFileReady  = "N"
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

    # SCOPE = input("Field Scope File ------> ").strip('\"')
    SCOPE = 'C:/Users/2095421/Downloads/Migration/Field_Scope_Data_Migration.xlsx'

    DVP = input("Path to store the Datavalidation file  ------> ").strip('\"')

    # MUCCA = "C:\\Users\\2095421\\Downloads\\Migration\\Prakhar\\MUCCA\\PROD"
    # DVP = MUCCA.strip('\"')
    
    DataValidationPath = DVP + '/'+ CLUSTER + '_' + TABLE + '_DataValidation.xlsx'


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
        source_sheet = source_wb.Sheets('Sheet1') #Make "Sheet1 if normal download is done"
        # source_sheet = source_wb.Sheets('MVKE') #Make "Sheet1" if normal download is done"

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
             
            
            # All table consideration
            row = 2
            last_row = dest_sheet.Cells(dest_sheet.Rows.Count, 2).End(-4162).Row #Not sure about the End(-4162)
            # while dest_sheet.Cells(row, 2).Value is not None and dest_sheet.Cells(row, 3).Value is not None:
            #     # if TABLE == 'MVKE' or TABLE == 'MARD':
            #     if TABLE in ['MVKE', 'MARD']:
            #         dest_sheet.Cells(row, 1).Value = str(dest_sheet.Cells(row, 2).Value) + str(dest_sheet.Cells(row, 3).Value) + str(dest_sheet.Cells(row, 4).Value)
            #     else:
            #         dest_sheet.Cells(row, 1).Value = str(dest_sheet.Cells(row, 2).Value) + str(dest_sheet.Cells(row, 3).Value)
            #     row += 1


            if TABLE in ['MVKE', 'MARD']:
                # dest_sheet.Range("A2:A"+str(last_row)).FormulaR1C1 = "=RC[1]&RC[2]&RC[3]"
                # dest_sheet.Range("A2:A"+str(last_row)).FormulaR1C1 = "=CONCAT(RC[1],RC[2],RC[3])"
                # dest_sheet.Range("A2:A"+str(last_row)).FormulaR1C1 = "=TEXT(RC[1], \"0\") & TEXT(RC[2], \"0\") & TEXT(RC[3], \"0\")"
                dest_sheet.Range("A2:A" + str(last_row)).FormulaR1C1 = "=RC[1] & RC[2] & RC[3]"


            else:
                dest_sheet.Range("A2:A" + str(last_row)).FormulaR1C1 = "=RC[1] & RC[2]"

            # Convert formulas to values
            dest_sheet.Range("A2:A"+str(last_row)).Value = dest_sheet.Range("A2:A"+str(last_row)).Value
            # dest_sheet.Range("A2:A"+str(last_row)).NumberFormat = "@"




            print("Key creation done Successfully!")    

            # print('Count of ', os.path.basename(system_name) + ": ", row-1)

        # cc = input("Do you need Colour Code the Fields? ")
        # if cc == 'y':
        #     print("<-----Starting to mark Scope Fields------> ")
        #     for row in range(2, scope_sheet.usedRange.Rows.Count+1):
        #         if scope_sheet.Cells(row, 2).Value == 'Scope':
        #             technical_name = scope_sheet.Cells(row, 1).Value
        #             for col in range(1, dest_sheet.UsedRange.Columns.Count+1):
        #                 if dest_sheet.Cells(1, col).Value == technical_name:
        #                     dest_sheet.Cells(1, col).Interior.ColorIndex = 4
        #                     break

        #     print("Scope fields marked with Green Successfully!")
        # else:
        #     print("Fields not coulr coded")

        dest_wb.Save()
        dest_wb.Close()
        source_wb.Close()
        excel_app.Quit()

        time.sleep(5)

        print("\n")


    fun_cp(MDG, DataValidationPath, TABLE)
    fun_cp(ATLAS, DataValidationPath, TABLE)



# df1 = pd.read_excel('C:/Users/2095421/Downloads/Migration/MBEW_DataValidation.xlsx', sheet_name='MBEW_M1Q_Ext.xlsx')
# df2 = pd.read_excel('C:/Users/2095421/Downloads/Migration/MBEW_DataValidation.xlsx', sheet_name='MBEW_APB_Ext.xlsx')

if ValidationFileReady in ["Y", "YES"]:
    MDG = input("MDG Sheet Name  :  ")
    ATLAS = input("ATLAS Sheet Name  :  ")
    df1 = pd.read_excel(DataValidationPath, sheet_name = MDG, index_col='KEY_' + MDG)
    df2 = pd.read_excel(DataValidationPath, sheet_name = ATLAS, index_col='KEY_' + ATLAS)
else:
    df1 = pd.read_excel(DataValidationPath, sheet_name=os.path.basename(MDG))
    df2 = pd.read_excel(DataValidationPath, sheet_name=os.path.basename(ATLAS))

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
    
    # if col_name not in ['BWKEY', 'VTWEG'] :
    if col_name not in df1.columns:
        print(col_name, "--------------Not Found")
    else:
        print(col_name)
        # Compare the FIELD values of df1 and df2 for the common indices
        comparison = (df2.loc[common_indices, col_name].isna() & df1.loc[common_indices, col_name].isna()) | df2.loc[common_indices, col_name].eq(df1.loc[common_indices, col_name])
        # comparison = (df2.loc[common_indices, col_name].isna() & df1.loc[common_indices, col_name].isna()) | (df2.loc[common_indices, col_name].astype(str) == df1.loc[common_indices, col_name].astype(str))
        # comparison = (df2.loc[common_indices, col_name].isna() & df1.loc[common_indices, col_name].isna()) | (df2.loc[common_indices, col_name].str.strip("'") == df1.loc[common_indices, col_name].str.strip("'"))

        # print(f"Size of common_indices: {len(common_indices)}")
        # print(f"Size of comparison: {len(comparison)}")


        # If they don't match, create a new dataframe with the mismatched indices
        if not comparison.all():
            mismatched_indices = common_indices[~comparison]

            if TABLE == 'MARA':
                df_mismatch = pd.DataFrame({
                'Matched': pd.Series(['no']*len(mismatched_indices), index=mismatched_indices),
                'Value in MDG': df1.loc[mismatched_indices, col_name].squeeze(),
                'Value in Atlas': df2.loc[mismatched_indices, col_name].squeeze(),
                })
            elif TABLE == 'MARC': 
                df_mismatch = pd.DataFrame({
                'Matched': pd.Series(['no']*len(mismatched_indices), index=mismatched_indices),
                'Material no': df2.loc[mismatched_indices, "MATNR"].squeeze(),
                'Plant': df2.loc[mismatched_indices, "WERKS"].squeeze(),
                'Value in MDG': df1.loc[mismatched_indices, col_name].squeeze(),
                'Value in Atlas': df2.loc[mismatched_indices, col_name].squeeze(),
                })
            elif TABLE == 'MBEW': 
                df_mismatch = pd.DataFrame({
                'Matched': pd.Series(['no']*len(mismatched_indices), index=mismatched_indices),
                'Material no': df2.loc[mismatched_indices, "MATNR"].squeeze(),
                'Plant': df2.loc[mismatched_indices, "BWKEY"].squeeze(),
                'Value in MDG': df1.loc[mismatched_indices, col_name].squeeze(),
                'Value in Atlas': df2.loc[mismatched_indices, col_name].squeeze(),
                })
            elif TABLE == 'MLGN': 
                df_mismatch = pd.DataFrame({
                'Matched': pd.Series(['no']*len(mismatched_indices), index=mismatched_indices),
                'Material no': df2.loc[mismatched_indices, "MATNR"].squeeze(),
                'LGNUM': df2.loc[mismatched_indices, "LGNUM"].squeeze(),
                'Value in MDG': df1.loc[mismatched_indices, col_name].squeeze(),
                'Value in Atlas': df2.loc[mismatched_indices, col_name].squeeze(),
                })
            elif TABLE == 'MLAN': 
                df_mismatch = pd.DataFrame({
                'Matched': pd.Series(['no']*len(mismatched_indices), index=mismatched_indices),
                'Material no': df2.loc[mismatched_indices, "MATNR"].squeeze(),
                'ALAND': df2.loc[mismatched_indices, "ALAND"].squeeze(),
                'Value in MDG': df1.loc[mismatched_indices, col_name].squeeze(),
                'Value in Atlas': df2.loc[mismatched_indices, col_name].squeeze(),
                })
            elif TABLE == 'MARD': 
                df_mismatch = pd.DataFrame({
                'Matched': pd.Series(['no']*len(mismatched_indices), index=mismatched_indices),
                'Material no': df2.loc[mismatched_indices, "MATNR"].squeeze(),
                'LGORT': df2.loc[mismatched_indices, "LGORT"].squeeze(),
                'Plant': df2.loc[mismatched_indices, "WERKS"].squeeze(),
                'Value in MDG': df1.loc[mismatched_indices, col_name].squeeze(),
                'Value in Atlas': df2.loc[mismatched_indices, col_name].squeeze(),
                })
            elif TABLE == 'MVKE': 
                df_mismatch = pd.DataFrame({
                'Matched': pd.Series(['no']*len(mismatched_indices), index=mismatched_indices),
                'Material no': df2.loc[mismatched_indices, "MATNR"].squeeze(),
                'VKORG': df2.loc[mismatched_indices, "VKORG"].squeeze(),
                'VTWEG': df2.loc[mismatched_indices, "VTWEG"].squeeze(),
                'Value in MDG': df1.loc[mismatched_indices, col_name].squeeze(),
                'Value in Atlas': df2.loc[mismatched_indices, col_name].squeeze(),
                })
            mismatched_dataframes[col_name] = df_mismatch

print("Validation Done, Saving all mismatch in the Excel file kindly wait..............")

# Save all the mismatched dataframes to new sheets in the excel file
with pd.ExcelWriter(DataValidationPath, engine='openpyxl', mode='a') as writer:
    for sheet_name, df_mismatch in mismatched_dataframes.items():
        df_mismatch.to_excel(writer, sheet_name=sheet_name.replace("/",""))

# Print a success message
print("All mismatched dataframes have been written to the excel file.")

end_time = time.time()
execution_time = end_time - start_time
print(f"The program took {execution_time} seconds to execute.")
from playsound import playsound


