import os
import win32com.client
try:
    import pandas as pd
except ImportError:
    os.system('pip install pandas')
    import pandas as pd

def FileExist(Path):
    if os.path.isfile(Path):
        print("The file exists.")
    else:
        print("The file does not exist.........!")

def key_generator(full_file_path, dest_sheet):
    dest_sheet.Columns(1).Insert()
    dest_sheet.Cells(1, 1).Value = 'KEY_' + os.path.basename(full_file_path)
    dest_sheet.Range("A1").Interior.Color = 65535
    dest_sheet.Range("A1").Font.Bold = True
    row = 2
    last_row = dest_sheet.Cells(dest_sheet.Rows.Count, 2).End(-4162).Row
    dest_sheet.Range("A2:A" + str(last_row)).FormulaR1C1 = "=RC[1] & RC[2]"
    dest_sheet.Range("A2:A"+str(last_row)).Value = dest_sheet.Range("A2:A"+str(last_row)).Value
    print("Key creation done Successfully!")

def Fun_CP(full_file_path, PreValidationFilePath, TABLE):
    excel_app = win32com.client.Dispatch("Excel.Application")
    source_wb = excel_app.Workbooks.Open(full_file_path)
    source_sheet = source_wb.Sheets('Sheet1')
    dest_wb = excel_app.Workbooks.Open(PreValidationFilePath)
    source_sheet.Copy(dest_wb.Sheets(1))
    print(os.path.basename(full_file_path), " File copied to the PreValidationFilePath file Successfully!" )
    dest_wb.Sheets(1).Name=os.path.basename(full_file_path)
    dest_sheet = dest_wb.Sheets(os.path.basename(full_file_path))

    if TABLE == 'MARA':
        key_generator(full_file_path, dest_sheet)

    dest_wb.Save()
    dest_wb.Close()
    source_wb.Close()
    excel_app.Quit()

ClusterName = input("Please give Cluster Name in CAPS:  ")
TABLE  = input("Table Name --->  ")
PVFP = input("Give file path where to create Comapre File and save:   ").strip('\"')
PreValidationFilePath = PVFP + "/" + ClusterName  + "_" + TABLE + "_PreValidation.xlsx"
df = pd.DataFrame()
df.to_excel(PreValidationFilePath, index = False)
print(f"CompareFile Excel created at ------> {PreValidationFilePath}.")

if TABLE == 'MARA':
    # MDG_FILE_PATH = "C:/Users/2095421/Downloads/Migration/PreValidation/MARA_M1Q_Ext.XLSX"
    MDG_FILE_PATH = input("Enter MDG_FILE_PATH path: ").strip('\"')
    FileExist(MDG_FILE_PATH)

    # ATLAS_FILE_PATH = "C:/Users/2095421/Downloads/Migration/PreValidation/MARA_APB_Ext.XLSX"
    ATLAS_FILE_PATH = input("Enter ATLAS_FILE_PATH path: ").strip('\"')
    FileExist(ATLAS_FILE_PATH)

    # GRD_FILE_PATH = "C:/Users/2095421/Downloads/Migration/PreValidation/MARA_AMB_Ext.XLSX"
    GRD_FILE_PATH = input("Enter GRD_FILE_PATH path: ").strip('\"')
    FileExist(GRD_FILE_PATH)

    Fun_CP(MDG_FILE_PATH, PreValidationFilePath, TABLE)
    Fun_CP(ATLAS_FILE_PATH, PreValidationFilePath, TABLE)
    Fun_CP(GRD_FILE_PATH, PreValidationFilePath, TABLE)

    df1 = pd.read_excel(PreValidationFilePath, sheet_name=os.path.basename(MDG_FILE_PATH), index_col='KEY_' + os.path.basename(MDG_FILE_PATH))
    df2 = pd.read_excel(PreValidationFilePath, sheet_name=os.path.basename(ATLAS_FILE_PATH), index_col='KEY_' + os.path.basename(ATLAS_FILE_PATH))
    df3 = pd.read_excel(PreValidationFilePath, sheet_name=os.path.basename(GRD_FILE_PATH), index_col='KEY_' + os.path.basename(GRD_FILE_PATH))

    mismatched_dataframes = {}

    DIFF12 = df1.index.difference(df2.index)# Key Present in df1 but not in df2
    if not DIFF12.empty:
        DIFF12_df = pd.DataFrame(DIFF12, columns=['Difference_Index'])# Save the difference indices to a new sheet
        with pd.ExcelWriter(PreValidationFilePath, engine='openpyxl', mode='a') as writer:
            DIFF12_df.to_excel(writer, sheet_name='MDG_ATLAS_ Diff', index=False)

    DIFF13 = df1.index.difference(df3.index)# Key Present in df1 but not in df3
    if not DIFF13.empty:
        DIFF13_df = pd.DataFrame(DIFF13, columns=['Difference_Index'])# Save the difference indices to a new sheet
        with pd.ExcelWriter(PreValidationFilePath, engine='openpyxl', mode='a') as writer:
            DIFF13_df.to_excel(writer, sheet_name='MDG_GRD Diff', index=False)

    DIFF23 = df2.index.difference(df3.index)# Key Present in df2 but not in df3
    if not DIFF23.empty:
        DIFF23_df = pd.DataFrame(DIFF23, columns=['Difference_Index'])# Save the difference indices to a new sheet
        with pd.ExcelWriter(PreValidationFilePath, engine='openpyxl', mode='a') as writer:
            DIFF23_df.to_excel(writer, sheet_name='ATLAS_GRD Diff', index=False)

    print("Difference indices checked successfully!")


    def MARS_Item_Roles(row):
        columns_with_x = [col[4:] for col in row.index if row[col] == 'X']
        return '_'.join(columns_with_x)

    df2['MARS_Item_ROLES'] = df2.apply(MARS_Item_Roles, axis=1)
    with pd.ExcelWriter(PreValidationFilePath, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        df2.to_excel(writer, sheet_name=os.path.basename(ATLAS_FILE_PATH), index=True)
    print("MARS_Item_R updated successfully!")

    def MSTAE_90_check(df, SheetName90):
        filtered_df = df[df['MSTAE'] == 90]
        if filtered_df.empty:
            print(f"Sheet Name: {SheetName90} - Filter DF is Empty")
            # break
        else:
            # Get the index and cell values of the column 'MATNR'
            matnr_data = filtered_df[['MATNR']].reset_index()
            # Save the data to a new sheet named 'MDG'
            with pd.ExcelWriter(PreValidationFilePath, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                matnr_data.to_excel(writer, sheet_name=SheetName90, index=False)
    MSTAE_90_check(df1, 'MDG90')
    MSTAE_90_check(df2, 'ATLAS90')
    MSTAE_90_check(df3, 'GRD90')


# else:
#     print('WOrking on MARC & MVKE------------> ')
#     MARC_File_Path = "C:/Users/2095421/Downloads/Migration/CHD/Quality/Scope Identitification/CHD FR25 AEB MARC.XLSX"
#     # MARC_file_path = input("Enter MARC File Path path: ").strip('\"')
#     FileExist(MARC_File_Path)
#     MVKE_File_Path = "C:/Users/2095421\Downloads/Migration/CHD/Quality/Scope Identitification/CHD 107 99 AEB MVKE.XLSX"
#     # MARC_file_path = input("Enter MVKE File Path path: ").strip('\"')
#     Fun_CP(MARC_File_Path, PreValidationFilePath, MARC)
#     Fun_CP(MVKE_File_Path, PreValidationFilePath, MVKE)