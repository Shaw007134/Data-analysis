from datetime import datetime,date
from xlrd import open_workbook, xldate_as_tuple
import sys
import os
import csv

input_file = "test_date.xlsx"
workbook = open_workbook(input_file)
worksheet = workbook.sheet_by_index(0)
header = worksheet.row_values(0)
print(header)

cell_type = worksheet.cell_type(1,0)
print(cell_type)

#date_value = datetime.date(datetime.strptime(worksheet.cell(1,0).value,'%m/%d/%Y'))
#print(date_value)
#date_value = date_value.strftime('%Y/%m/%d')
#print(date_value)

#cell_value = worksheet.cell(1,0).value = date_value



#print(cell_type)

cell_value = xldate_as_tuple(worksheet.cell(1,0).value,workbook.datemode)
print(cell_value)



print(datetime(*cell_value))
print(date(*cell_value[:3]))
print(*cell_value[3:6])
print(*cell_value[0:3])
print(cell_value[0:3])

value = worksheet.cell(2,0).value
print("type of value: {}".format(type(value)))
print("value is {}".format(value))
value = str(value).split('.')[0]
print('\n value is :{}'.format(value))



