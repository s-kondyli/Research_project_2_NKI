# script that takes as input an excel workbook and randomizes rows of a certain datasheet
# #based on the number of samples/patients the user defines
# ideal for randomizing a cohort

import pandas as pd
import os
import time
path = os.getcwd() #get the absolute path to the current working directory
files = os.listdir(path) #list of files in the cwd
excel_files = [f for f in files if f[-4:] == 'xlsx'] #only excel files in the cwd
timestr = time.strftime("%Y%m%d%H%M") #time string for the output file

#request data
while True:
    try:
        workbook = input("Enter the excel workbook:")
        if workbook in excel_files:
            break
        else:
            print(workbook,'is not a valid name.\nValid files are:', excel_files)
    except Exception:
        print(workbook,'is not a valid name.\nValid files are:', excel_files)

sheet_dict = pd.read_excel(workbook, sheet_name=None)
while True:
    sheetname = input("Enter sheetname containing data for randomization:")
    if sheetname not in sheet_dict.keys():
        print("Sheetname not valid.\nValid sheetnames are:", sheet_dict.keys())
    else:
        break

#input is by default str but you have to convert it to str for the sample method
number = int(input("Enter the desired number of patients in the randomized cohort:"))

df = sheet_dict.get(sheetname)
df_randomized = df.sample(n=number)
#store your excel file in the cwd with the desired name
df_randomized.to_excel("Randomized_cohort_"+ timestr+"_.xlsx", sheet_name=sheetname, index=False)
print("All done :)")
