#Prerequisitories
#MATNR, WREKS, VKWEG, LGORT  colums must be in 2,3,4 colums as needed
#MDG / Atlas File must be starting with the Colums that will be used for "Key" creation(For MBEW MATNR & BWKEY)



import os
import win32com.client
import pandas as pd
# import openpyxl

TABLE = input(" TABLE NAME : ").upper()

MDG = 'C:/Users/2095421/Downloads/Migration/MBEW_M1Q_Ext.xlsx'
ATLAS = 'C:/Users/2095421/Downloads/Migration/MBEW_APB_Ext.xlsx'



# MDG = input(" Please give MDG "+ TABLE + " file path ").strip('\"')
# if os.path.isfile(MDG):
#     print("The file exists.")
# else:
#     print("The file does not exist.")
# 
# ATLAS = input(" Please give ATLAS "+ TABLE + " file path ").strip('\"')
# if os.path.isfile(ATLAS):
#     print("The file exists.")
# else:
#     print("The file does not exist.")




df = pd.DataFrame()

DataValidationPath = 'C:/Users/2095421/Downloads/Migration/' + TABLE + '_DataVALIDATION.xlsx'

df.to_excel(DataValidationPath)

scope_df = pd.read_excel('C:/Users/2095421/Downloads/Migration/Field_Scope_Data_Migration.xlsx', sheet_name=TABLE)
scope_df.to_excel(DataValidationPath, sheet_name='Sheet1', index=False)


print(TABLE + " Data Validation File created at - " + DataValidationPath )
print("\n")



def fun_cp(system_name, DataValidationPath, TABLE):
    excel_app = win32com.client.Dispatch("Excel.Application")
    source_wb = excel_app.Workbooks.Open(system_name)
    source_sheet = source_wb.Sheets('Sheet1')

    dest_wb = excel_app.Workbooks.Open(DataValidationPath)
    
    # source_sheet.Copy(Before = dest_wb.Sheets(1))
    source_sheet.Copy(dest_wb.Sheets(1))    #Here Sheet(1) means in destination file before the 1st sheet the new sheet will get created
    print(os.path.basename(system_name), " File copied to the DataValidation file Successfully" )

    scope_sheet = dest_wb.Sheets('Sheet1')

    dest_wb.Sheets(1).Name=os.path.basename(system_name)
    dest_sheet = dest_wb.Sheets(os.path.basename(system_name))
    dest_sheet.Columns(1).Insert()
    dest_sheet.Cells(1, 1).Value = 'KEY_' + os.path.basename(system_name)
    # dest_sheet.Cells(1, 1).Value = 'KEY_' + system_name
    dest_sheet.Range("A1").Interior.Color = 65535
    dest_sheet.Range("A1").Font.Bold = True
    print("Key creation done Successfully!")


    # All table consideration
    row = 2  # start from the second row assuming the first row contains headers
    while dest_sheet.Cells(row, 2).Value is not None and dest_sheet.Cells(row, 3).Value is not None:
        if TABLE == 'MVKE' or TABLE == 'MARD':
            dest_sheet.Cells(row, 1).Value = str(dest_sheet.Cells(row, 2).Value) + str(dest_sheet.Cells(row, 3).Value) + str(dest_sheet.Cells(row, 4).Value)
        else:    
            dest_sheet.Cells(row, 1).Value = str(dest_sheet.Cells(row, 2).Value) + str(dest_sheet.Cells(row, 3).Value)
        row += 1

    print('Count of ', os.path.basename(system_name) + ": ", row-1)

    for row in range(2, scope_sheet.usedRange.Rows.Count+1):
        if scope_sheet.Cells(row, 2).Value == 'Scope':
            technical_name = scope_sheet.Cells(row, 1).Value

            for col in range(1, dest_sheet.UsedRange.Columns.Count+1):
                if dest_sheet.Cells(1, col).Value == technical_name:
                    # If found, highlighting the cell with Green colour(4)
                    dest_sheet.Cells(1, col).Interior.ColorIndex = 4
                    break
                # else:
                #     dest_sheet.Cells(1, col).Interior.ColorIndex = 3
    
    print("Scope fields marked with Green Successfully!")

    dest_wb.Save()
    dest_wb.Close()
    source_wb.Close()
    excel_app.Quit()

    print("\n")

fun_cp(MDG, DataValidationPath, TABLE)
fun_cp(ATLAS, DataValidationPath, TABLE)




# Loading the sheets into dataframes
df1 = pd.read_excel(DataValidationPath, sheet_name=os.path.basename(MDG))
df2 = pd.read_excel(DataValidationPath, sheet_name=os.path.basename(ATLAS))

# Set the index to be the key for easier comparison
df1.set_index('KEY_' + os.path.basename(MDG), inplace=True)
df2.set_index('KEY_' + os.path.basename(ATLAS), inplace=True)

# Get the column of names in iteration
df3 = pd.read_excel(DataValidationPath, sheet_name='Sheet1')
df_scope = df3[df3['Comment'] == 'Scope']
scope_field = df_scope['Technical Name']


for col_name in scope_field:
    print(col_name)
    # if col_name != 'MATNR' and col_name != 'BWKEY' :
    if col_name not in ['MATNR', 'BWKEY', 'VKWEG', 'LGORT'] :

        # col_name = 'PEINH'
        result_df = pd.DataFrame(columns=['Key','Value-Matched?','Material','Plant', 'ATLAS', 'MDG'])
        # print(type(result_df))

        for index in df2.index:
            # Check if the index is in df2
            if index in df1.index:
                # Check if the values are the same
                if df2.loc[index, col_name] == df1.loc[index, col_name] or (pd.isna(df2.loc[index, col_name]) and pd.isna(df1.loc[index, col_name])):
                    # print(f'Key {index}: Yes')
                    pass
                else:
                    # print(f'Key {index}: No')
                    result_df.loc[len(result_df)] = [index, 'No',df2.loc[index, 'MATNR'],df2.loc[index, 'BWKEY'], df1.loc[index, col_name], df1.loc[index, col_name]]
            else:
                # print(f'Key {index}: Key not found')
                result_df.loc[len(result_df)] = [index, 'Key Not Found','','','','']

        if not result_df.empty:
            writer = pd.ExcelWriter(DataValidationPath, engine='openpyxl', mode='a')
            result_df.to_excel(writer, sheet_name=col_name)
            writer.close()