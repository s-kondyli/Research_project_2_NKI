#this script takes as input: data to be filtered in excel file,excel file with filters in sheets
#and gives as output an excel file with the filtered data in the current working directory
import pandas as pd
import time
import os
path = os.getcwd() #get the absolute path to the current working directory
files = os.listdir(path) #list of files in the cwd
excel_files = [f for f in files if f[-4:] == 'xlsx'] #only excel files in the cwd
timestr = time.strftime("%Y%m%d%H%M") #time string for the output file

#defining a function that takes data from excel files and applies conditional filtering to columns based on a list of dictionaries
#the keys are the columns you apply a filter on and the values are threshold/filter
#2 dictionaries in the list: 1)for columns to be filtered for greater than/equal to a value and 2)for columns to be filtered for less than/equal to a value
def filter(data,dict_list):
    greater = dict_list[0]
    lower = dict_list[1]
    for column, value in greater.items():
        data = data[data[column] >= value]
    for column, value in lower.items():
        filtered_data = data[data[column] <= value]
    return filtered_data

#input the excel workbook and catch errors if the user gave wrong filename
while True:
    try:
        workbook = input("Enter the excel workbook:")
        if workbook in excel_files:
            break
        else:
            print(workbook,'is not a valid name.\nValid files are:', excel_files)
    except Exception:
        print(workbook,'is not a valid name.\nValid files are:', excel_files)
#create a dictionary of the imported workbook with keys the sheetnames and values the dataframes in each sheet
sheet_dict = pd.read_excel(workbook, sheet_name=None)
#input the sheetname for the desired data to be analysed and catch for errors if the user gave wrong sheetname
while True:
    sheetname = input("Enter sheetname that containts the data to be filtered:")
    if sheetname not in sheet_dict.keys():
        print("Sheetname not valid.\nValid sheetnames are:", sheet_dict.keys())
    else:
        break

#assign the desired df of the workbook based on the given sheetname to a data variable
data = sheet_dict.get(sheetname)
# Clean up data
data.dropna(how='all', inplace=True) # drop the rows where in all columns of the df the values are NaN, so keep rows that NaN are only in some columns
data.fillna(value=0, inplace=True)
data.fillna(value=0, inplace=True)
while True:
    try:
        workbook2 = input("Enter the excel with the filters:")
        if workbook2 in excel_files:
            break
        else:
            print(workbook2,'is not a valid name.\nValid files are:', excel_files)
    except Exception:
        print(workbook2,'is not a valid name.\nValid files are:', excel_files)

filter_dict = pd.read_excel(workbook2, sheet_name=None, names=['column_name', 'filter']) #import our filters
dict_list = [] #create an empty list for the dictionaries

#transform the dataframes with the filters to dictionaries
for key, value in filter_dict.items():
    value.set_index('column_name', inplace=True)
    value = value.to_dict()['filter'] #really important to have the filter because nested dictionary otherwise
    dict_list.append(value)

#apply our function
filtered_data = filter(data, dict_list)
#save the filtered data to an excel file to the current working directory
filtered_data.to_excel(excel_writer='filtered_data'+ timestr+'.xlsx', index=False, sheet_name=sheetname)
print('All done!')
