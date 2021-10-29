#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on 27 Oct 2021

@author: mohammedalbatati
"""
import pandas as pd
import xlwings as xw
import os
from rich import print

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

try:
    lastMonth = os.path.join(package_dir, 'lastMonth.xlsx')
    currentMonth = os.path.join(package_dir, 'currentMonth.xlsx')
    # ops_file = os.path.join(package_dir, '/Book*.xlsx')
except FileNotFoundError as error:
    print(error)


comp_book = xw.Book()
comp_book.app.visible = False
comp_curr_cons_sht = comp_book.sheets.add('Current consultants')
comp_last_cons_sht = comp_book.sheets.add('Last consultants')
comp_wtc_sht = comp_book.sheets.add('Field')
comp_over_sht = comp_book.sheets.add('Overhead')

current_book = xw.Book(currentMonth)
current_book.app.visible = False
try:
    current_sht_cons = current_book.sheets['Consultnat']
    current_sht_wtc = current_book.sheets['WTC']
    current_sht_over = current_book.sheets['Timesheet']
except Exception as error:
    print(error)

last_book = xw.Book(lastMonth)
current_book.app.visible = False
try:
    last_sht_cons = last_book.sheets['Consultnat']
    last_sht_wtc = last_book.sheets['WTC']
    last_sht_over = last_book.sheets['Timesheet']
except Exception as error:
    print(error)


##########################
# update the current month sheet
##########################
# No 1 - Consultants values from current month (from 1st to 15th)
con_names = current_sht_cons.range('I3:I60').options(ndim=2).value
con_id = current_sht_cons.range('A3:A60').options(ndim=2).value
con_curr_val = current_sht_cons.range('AD3:AQ60').value
# Moving the data to the ops stat sheet
comp_curr_cons_sht.range('D1').value = con_curr_val
comp_curr_cons_sht.range('A1').value = con_id
comp_curr_cons_sht.range('C1:C60').value = 'Consultant'
comp_curr_cons_sht.range('B1').value = con_names

# No 2 - Consultants values from last month (from 15th to 30th)
con_names = current_sht_cons.range('I3:I60').options(ndim=2).value
con_id = current_sht_cons.range('A3:A60').options(ndim=2).value
con_curr_val = current_sht_cons.range('AR3:BH60').value
# Moving the data to the ops stat sheet
comp_last_cons_sht.range('D1').value = con_curr_val
comp_last_cons_sht.range('A1').value = con_id
comp_last_cons_sht.range('C1:C60').value = 'Consultant'
comp_last_cons_sht.range('B1').value = con_names

# No 2 - SWT & SLS details (change the row number TODO )
# will not calculate the current month at this stage TODO
wtc_names = last_sht_wtc.range('I3:I37').options(ndim=2).value
wtc_bl = last_sht_wtc.range('B3:B37').options(ndim=2).value
wtc_id = last_sht_wtc.range('A3:A37').options(ndim=2).value
wtc_curr_val = last_sht_wtc.range('AD3:BH37').value

comp_wtc_sht.range('D1').value = wtc_curr_val
comp_wtc_sht.range('A1').value = wtc_id
comp_wtc_sht.range('C1:C60').value = wtc_names
comp_wtc_sht.range('B1').value = wtc_bl

# No 3 - overhead details:
over_names = current_sht_over.range('B14:B22').options(ndim=2).value
over_id = current_sht_over.range('A14:A22').options(ndim=2).value
over_curr_val = current_sht_over.range('E14:AI22').value

comp_over_sht.range('A1').value = over_id
comp_over_sht.range('B1:B6').value = 'SWT'
comp_over_sht.range('C1').value = over_names
comp_over_sht.range('D1').value = over_curr_val

comp_book.save('Book1.xlsx')

comp_book.close()
current_book.close()
last_book.close()


