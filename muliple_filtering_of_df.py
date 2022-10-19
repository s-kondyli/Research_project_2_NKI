import pandas as pd
import sys
#this script takes as input: data to be filtered in excel file,excel file with filters in sheets
#and gives as output an excel file with the filtered data in the current working directory

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

data = pd.read_excel(sys.argv[1]) #import our data
data.fillna(value=0, inplace=True)
sheet_dict = pd.read_excel(sys.argv[2], sheet_name=None, names=['column_name', 'filter']) #import our filters
dict_list = [] #create an empty list for the dictionaries

#transform the dataframes with the filters to dictionaries
for key, value in sheet_dict.items():
    value.set_index('column_name', inplace=True)
    value = value.to_dict()['filter'] #really important to have the filter because nested dictionary otherwise
    dict_list.append(value)

#apply our function
filtered_data = filter(data, dict_list)
#save the filtered data to an excel file to the current working directory 
filtered_data.to_excel(excel_writer='filtered_data.xlsx', index=False, sheet_name='Data')
