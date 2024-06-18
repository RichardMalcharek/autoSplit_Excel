import pandas as pd
import os
from datetime import datetime

#Asking for Folder
strPath = str(input("Folder path (e.g. C:\\Users\\Dokumente ) :"))

#Asking for filename
strFile = str(input("Name of the file (e.g. Source.xlsx) :"))

#Asking for relevant sheet
strWorksheet = str(input("Name of the worksheet :"))

#Asking for column where the responsible Person is in
strColumn = str(input("Name of the Column/Headline (e.g. Responsible) case sensitive!! :"))

#create file path
strFile_Path = str(strPath + "\\" + strFile)

#open excel file
df = pd.read_excel(strFile_Path, sheet_name=strWorksheet)

#Save all responsibles in variable
arrResponsible = list(set(df[strColumn].tolist()))

#Create new folders for files
strNew_Folder = strPath + "\\" + datetime.now().strftime('%Y-%m-%d_%H.%M.%S')
os.makedirs(strNew_Folder)

#Creates sub-folder for each responsible
for strShort in arrResponsible:
    os.makedirs(strNew_Folder+"\\"+strShort)
    filtered_df = df[df[strColumn] == strShort]
    filtered_df.to_excel(strNew_Folder+"\\"+strShort+"\\"+strShort+datetime.now().strftime('%Y-%m-%d_%H.%M.%S')+".xlsx", index=False)


