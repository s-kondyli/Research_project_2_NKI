import pandas as pd
import sys
#we want to find the mutual values from 2 dataframes

data = pd.read_excel(sys.argv[1]) 
df = pd.read_excel(sys.argv[2], names=['patient'])
data.fillna(value=0, inplace=True)

check = df['patient'].tolist()
data = data[data['HiX#'].isin(check)] #we keep only the values of data['HiX#'] that are the in the list check 


