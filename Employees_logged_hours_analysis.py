
#This allows you to use pandas functions by simply typing pd. function_name rather than pandas.
#pandasused for (data manipulation and analysis)

import pandas as pd

# add file path to python 
path = r"C:\Users\Hazem\Desktop\VALEO\Final PPT\Employees KPIs per Department\Intial Employees Data.xlsx"

#read excel sheet 


sheet1 = pd.read_excel(path, sheet_name="Hours")  
sheet2 = pd.read_excel(path, sheet_name="Engineering") 

#check that each column i want to remove has same fixed values acroos all rows
 
#sheet1.drop_duplicates(subset=['Project Code'], keep= False, inplace=True)
#sheet1.drop_duplicates(subset=['Calendar'], keep= False, inplace=True)
#sheet1.drop_duplicates(subset=['Period_To'], keep= False, inplace=True)
#sheet1.drop_duplicates(subset=['Period_From'], keep= False, inplace=True)
#sheet1.drop_duplicates(subset=["Calendar", "Period_To","Period_From"], keep=False,inplace=True)

#after checking that they all have the same values across all rows , we delet them
del sheet1 ['Calendar']
del sheet1 ['Period_To']
del sheet1 ['Period_From']
del sheet1 ['Cost Rate']


#Create and insert 3 new cloumns to replace the Calendar & (period to & period from)

#new_cloumn= [*range (1,2)]
#print(new_cloumn)

#sheet1.insert(13,"new column",new_cloumn) {error list values must be equal },so i tried a different approach

sheet1['Fixed Duration'] = pd.Series([" Week 24:42"])
sheet1['Fixed Calender'] = pd.Series(["EG_02"])
sheet1['Fixed Cost'] = pd.Series(["33.32"])

#instead of filling the remaning cells with NaN , just leave them emty blank 
sheet1 = sheet1.fillna("")



print(sheet1)



from datetime import datetime

#calculate the differnce bteween end date and starting date to know time taken to close tickets
#create variable and assign differnce in it 
date_diff=(sheet2["End Date"]-sheet2["Start Date"]).dt.days
#insert new variable to our table 
sheet2.insert(4,"Date differnce",date_diff)

#print (sheet2)

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('New Use case (3).xlsx', engine='xlsxwriter')

# Write each dataframe to a different worksheet.
sheet1.to_excel(writer, sheet_name='Sheet1')
sheet2.to_excel(writer, sheet_name='Sheet2')

# Close the Pandas Excel writer and output the Excel file.
writer.save()

#sheet1.to_excel("New Demo.xlsx",sheet_name="sheet1")
#sheet2.to_excel("New Demo.xlsx",sheet_name="sheet2")




















