import pandas as pd
import matplotlib.pyplot as plt

excel_file_1 = 'shift-data.xlsx'
excel_file_2 = 'third-shift-data.xlsx'

# The sheet_name attribute refers to the sheet you want to access
# in an excel file. If none is specified, it defaults to the first sheet
df_first_shift = pd.read_excel(excel_file_1, sheet_name='first')
df_second_shift = pd.read_excel(excel_file_1, sheet_name='second')
df_third_shift = pd.read_excel(excel_file_2)

#Prints first sheet from first file
print(df_first_shift)

#Prints 'Product' column from first file
print(df_first_shift['Product'])

# Combines all dataframes into a single dataframe
# Joins all the values where the column heading is the same
df_all = pd.concat([df_first_shift, df_second_shift, df_third_shift])
print(df_all)

# 
pivot = df_all.groupby(['Shift']).mean()
shift_productivity = pivot.loc[:, "Production Run Time (Min)":"Products Produced (Units)"]
print(shift_productivity)

# Graph it (bar graph)
shift_productivity.plot(kind='bar')
plt.show()

# Output data to new excel workbook
df_all.to_excel("output.xlsx")