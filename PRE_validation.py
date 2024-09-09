import os
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

MDG = input("Enter MDG path: ").strip('\"')
FileExist(MDG)

# ATLAS = input("Enter ATLAS path: ").strip('\"')
# FileExist(ATLAS)

# GRD = input("Enter GRD path: ").strip('\"')
# FileExist(GRD)

CompareFilePath = input("Give file path where to create Comapre File and save:   ").strip('\"')
CompareFile = pd.DataFrame()
CompareFile.to_excel(CompareFilePath + "/PreValidation.xlsx", index = False)
print(f"CompareFile Excel created at ------> {CompareFilePath}.")

# df1 = pd.read_excel(MDG, sheet_name = "Sheet1")
# df2 = pd.read_excel()