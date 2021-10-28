#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on 27 Oct 2021

@author: mohammedalbatati
"""
import pandas as pd
import os
from rich import print
from helper_function import sum_legends,clean_df_na_set_index, merge_two_dataframe

'''
This code is directed to open last month crew time sheet and picks up the time sheet from 26 to 31 of last month and
past it in 4 sheets fom the ops stat excel sheet last it opens the current month excel and copy time sheet from 1st to
25th of the current month and then paste it to the ops stat sheet in the correct sheet and position
The code is very dirty and repetitive and is not factored but it do the task faster than I normally do
- I download the files from OnDrive and rename it to lastMonth and currentMonth and keep a copy of the opsstat sheet
in the same directory.
The code will use the same template sheet opsStat.xlsx so no need to save it or even its better to save as different file
This code is the same one used with openpyxl but using xlwings which is much faster due to the use of copy and paste option which is
more faster then iterating over every single line
Created on : 22 Oct 2021

'''
package_dir = os.path.dirname(os.path.abspath(__file__))

time_sheet = os.path.join(package_dir, 'Book1.xlsx')
crew_cost = os.path.join(package_dir, "Crew cost.xlsx")

# Load the data frames from the time sheet 
df_consultant_last = pd.read_excel(time_sheet, sheet_name='Last consultants',header=None)
df_consultant_current = pd.read_excel(time_sheet, sheet_name='Current consultants',header=None)
df_field_crew = pd.read_excel(time_sheet, sheet_name='Field',header=None)
df_overhead_crew = pd.read_excel(time_sheet, sheet_name='Overhead',header=None)

# Load the data for the crew cost excel sheet
cost_df = pd.read_excel(crew_cost, sheet_name='Crew cost')
# Convert the ID to integer and make it the index
cost_df['employee #'] = cost_df['Employee #'].astype(int)
cost_df = cost_df.set_index("Employee #")
# in case you need to remove some columns then uncomment the below line
# cost_df.drop(columns=['BL', 'COVID' ,'Sub BL', 'Employee Name', 'Type', 'Function Description','International Type'], axis=1, inplace=True)

# Use the function clean def to remove the na and clean it
df_consultant_last = clean_df_na_set_index(df_consultant_last,0)
df_consultant_current = clean_df_na_set_index(df_consultant_current, 0)
df_field_crew = clean_df_na_set_index(df_field_crew,0)
df_overhead_crew = clean_df_na_set_index(df_overhead_crew,0)

# Calculate the summation of the daily time sheet KB's, TB's etc
df_consultant_current = sum_legends(df_consultant_current)
df_consultant_last = sum_legends(df_consultant_last)
df_field_crew = sum_legends(df_field_crew)
df_overhead_crew = sum_legends(df_overhead_crew)

# Merge the data frames with the cost dataframe 
df_consultant_current_comp = merge_two_dataframe(df_consultant_current,cost_df)
df_consultant_last_comp = merge_two_dataframe(df_consultant_last,cost_df)
df_field_crew = merge_two_dataframe(df_field_crew, cost_df)
df_overhead_crew = merge_two_dataframe(df_overhead_crew, cost_df)

# calculate the current month for the consultants 1 - 14th 
df_consultant_current_comp['consultant_salary'] = (
    ((df_consultant_current_comp["KB"]) * df_consultant_current_comp["Base Rate (Daily Rate)"])
    + ((df_consultant_current_comp["KB"]) * df_consultant_current_comp["Meal Allowance"])
    + (df_consultant_current_comp["TB1"] * df_consultant_current_comp["Wellsite Rate (J1)"])
    + (df_consultant_current_comp["TB2"] * df_consultant_current_comp["Wellsite Rate (J2)"])
)

# calculate the last month for the consultants 15 - 31st 
df_consultant_last_comp['consultant_salary'] = (
    ((df_consultant_last_comp["KB"]) * df_consultant_last_comp["Base Rate (Daily Rate)"])
    + ((df_consultant_last_comp["KB"]) * df_consultant_last_comp["Meal Allowance"])
    + (df_consultant_last_comp["TB1"] * df_consultant_last_comp["Wellsite Rate (J1)"])
    + (df_consultant_last_comp["TB2"] * df_consultant_last_comp["Wellsite Rate (J2)"])
)

# Calculate the salary for the field crew employess SWT and SLS
df_field_crew['field_salary'] = (
    (
        (df_field_crew["KB"] + df_field_crew["D"] + df_field_crew["TB1"]
         + df_field_crew["TB2"]) * df_field_crew["Base Rate (Daily Rate)"]
    )
    + (df_field_crew["TB1"] * df_field_crew["Wellsite Rate (J1)"])
    + (df_field_crew["TB2"] * df_field_crew["Wellsite Rate (J2)"])
)

# Calculate the Over head salary
df_overhead_crew['overhead_salary'] = (
    (
     (df_overhead_crew["KB"] + df_overhead_crew["D"] + df_overhead_crew["TB1"]
      + df_overhead_crew["TB2"]) * df_overhead_crew["Base Rate (Daily Rate)"]
    )
    + (df_overhead_crew["TB1"] * df_overhead_crew["Wellsite Rate (J1)"])
    + (df_overhead_crew["TB2"] * df_overhead_crew["Wellsite Rate (J2)"])
)

# Calculate the total sum of all crew compensation
total_com = df_overhead_crew['overhead_salary'].sum() + df_consultant_current_comp['consultant_salary'].sum() + df_consultant_last_comp['consultant_salary'].sum()+df_field_crew['field_salary'].sum()

# print('consultants current:   ', df_consultant_current_comp['consultant_salary'].sum())
print('consultants current:   ', df_consultant_current_comp['consultant_salary'].agg(['sum','count']))
# print('consultants last   :   ', df_consultant_last_comp['consultant_salary'].sum())
print('consultants last   :   ', df_consultant_last_comp['consultant_salary'].agg(['sum','count']))
print('field Salary       :   ', df_field_crew['field_salary'].sum())
print('Overhead Salary    :   ', df_overhead_crew['overhead_salary'].sum())
print(f"Total compensation for this month is {total_com}")
print('================ Analysis for field crew ===================')
# print(df_field_crew.groupby(1).sum()['field_salary'])
print(df_field_crew.groupby(1).agg(['sum', 'count'])['field_salary'])
print('================ Analysis for field crew split==============')
print(df_field_crew.groupby([1, 'International Type']).agg(['sum', 'count'])['field_salary'])
# print(df_overhead_crew.columns)


