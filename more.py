# import win32com.client
# from openpyxl import Workbook
# import pandas as pd

# TABLE = 'MBEW'

# df = pd.DataFrame()

# DataValidationPath = 'C:/Users/2095421/Downloads/Migration/' + TABLE + '_DataVALIDATION.xlsx'

# df.to_excel(DataValidationPath)

# scope_df = pd.read_excel('C:/Users/2095421/Downloads/Migration/Field_Scope_Data_Migration.xlsx', sheet_name=TABLE)
# scope_df.to_excel(DataValidationPath, sheet_name='Sheet1', index=False)



# excel_app = win32com.client.Dispatch("Excel.Application")
# excel_app.Quit()



# MDG = 'C:/Users/2095421/Downloads/Migration/MBEW_M1Q_Ext.xlsx'
# ATLAS = 'C:/Users/2095421/Downloads/Migration/MBEW_APB_Ext.xlsx'
# TABLE = "MARC"
# MDG = input(" Please give MDG "+ TABLE + " file path").strip('\"')
# if os.path.isfile(MDG):
#     print("The file exists.")
# else:
#     print("The file does not exist.")

# excel_app = win32com.client.Dispatch("Excel.Application")
# source_wb = excel_app.Workbooks.Open(MDG)

# source_sheet = source_wb.Sheets('Sheet2')
# dest_wb = excel_app.Workbooks.Open('C:/Users/2095421/Downloads/Migration/MBEW_DataValidation.xlsx')
# # source_sheet.Copy(Before = dest_wb.Sheets(1)) #for postion of the sheet

# source_sheet.Copy(dest_wb.Sheets(1))
# print(os.path.basename(MDG), " File copied to the DataValidation file Successfully" )
# dest_wb.Save()
# dest_wb.Close()
# source_wb.Close()
# excel_app.Quit()

# import time

# # Start the timer
# start_time = time.time()

# # Your program starts here
# print("Hello, World!")
# # Your program ends here

# # End the timer
# end_time = time.time()

# # Calculate the execution time
# execution_time = end_time - start_time

# # Print the execution time
# print(f"The program took {execution_time} seconds to execute.")

















