import os
import win32com.client

excel_app = win32com.client.Dispatch("Excel.Application")
excel_app.Quit()