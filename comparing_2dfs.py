import pandas as pd
import sys
import time
# this script compares a desired column of two dataframes
# mutual values are stored in a excel sheet marked as mutual and the non-mutual values of the 2dfs are stored in 2 separate sheets
timestr = time.strftime("%Y%m%d") #stores current date(year,month,day) as a str
#input data
df1 = pd.read_excel(sys.argv[1],  parse_dates=['DOB','DOA'])
df2 = pd.read_excel(sys.argv[2])
# clean up NaN values
df1.dropna(how='all', inplace=True) # drop the rows where in all columns of the df the values are NaN, so keep rows that NaN are only in some columns
df2.dropna(how='all', inplace=True)

df1.fillna(value=0, inplace=True)
df2.fillna(value=0, inplace=True)

initial_df1 = df1['HiX#'].size
initial_df2 = df2['patients'].size
#compare the dfs
mutual = df1[df1['HiX#'].isin(df2['patients'])] # mutual patients
df1only = df1[~df1['HiX#'].isin(df2['patients'])] # patients unique in dataframe1
df2only = df2[~df2['patients'].isin(df1['HiX#'])] # patients unique in dataframe2

mutual_patients = mutual['HiX#'].size
rest_df1 = initial_df1 - mutual_patients
rest_df2 = initial_df2 - mutual_patients
# print some valuable info
print("Initial patients df1:",initial_df1,"\nInitial patients df2:",initial_df2,"\nMutual patients:",mutual_patients,"\n Rest df1:", rest_df1, "\n Rest df2:", rest_df2)
# store the dfs into a dictionary
frames = {'Mutual_patients': mutual, 'general_bs': df1only, '1-year_prior_bs': df2only}

#loop thru the dict of dfs and put each on a specific sheet of an excel workbook
#in the filename we include the date the file was created
with pd.ExcelWriter('Baseline_samples_'+ timestr +'_.xlsx') as writer:
    for sheet, frame in frames.items():  # .use .items for python 3.X
        frame.to_excel(writer, sheet_name=sheet, index=False)


