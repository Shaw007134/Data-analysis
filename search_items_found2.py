import csv
import glob
import os
import sys
from datetime import date,datetime
from xlrd import open_workbook, xldate_as_tuple

item_numbers_file = sys.argv[1]
path_to_folder = sys.argv[2]
output_file = sys.argv[3]

item_to_find = []
#读取时newline=''使用universal mode,返回untranslated
with open(item_numbers_file, 'r', newline='') as item_csv_file:
  filereader = csv.reader(item_csv_file)
  for row in filereader:
    item_to_find.append(row[0])
  print(item_to_find)
#写入时newline=''不转换，返回untranslated
filewriter = csv.writer(open(output_file,'a',newline=''))

file_counter = 0
line_counter = 0
count_of_items = 0
for input_file in glob.glob(os.path.join(path_to_folder,'*.*')):
  file_counter +=1
  if input_file.split('.')[1] == 'csv':
    with open(input_file,'r',newline='') as csv_in_file:
      filereader = csv.reader(csv_in_file)
      header = next(filereader)
      for row in filereader:
        row_of_output = []
        for col in range(len(header)):
          if col < 3:
            cell_value = str(row[col]).strip()
          elif col == 3:
            cell_value = str(row[col]).lstrip('$').\
            replace(',','').split('.')[0].strip()
          else:
            #以下strptime将string转换为datetime.date类型。
            #第二个参数表示string的日期格式
            cell_value = datetime.date(\
            datetime.strptime(str(row[col]),'%m/%d/%Y'))
            #以下strftime将datetime.date格式化后转换为字符串。
            #参数表示需要格式化成何种日期格式
            cell_value = cell_value.strftime('%Y/%m/%d')
          row_of_output.append(cell_value)
        row_of_output.append(os.path.basename(input_file))
        if row[0] in item_to_find:
          filewriter.writerow(row_of_output)
          count_of_items += 1
        line_counter += 1
  if input_file.split('.')[1] in ['xls','xlsx']:
    workbook = open_workbook(input_file)
    #可以通过workbook.sheet_by_index(number)进行访问某个worksheet
    for worksheet in workbook.sheets():
      try:
        header = worksheet.row_values(0)
      except IndexError:
        pass
      for row in range(1,worksheet.nrows):
        row_of_output = []
        for col in range(len(header)):
          if col < 3:
            cell_value = str(worksheet.cell_value(row,col)).strip()
          elif col == 3:
            #以下对值进行转换，excel中$xxx,xxx.xxx可通过str将浮点数转换为字符串
            #并通过split取小数点前的字符
            cell_value = str(worksheet.cell_value(row,col)).split('.')[0].strip()
          else:
            #xldate_as_tuple函数将日期单元格转换为(year,month,day,x,x,x)的元组
            #通过cell_type=3来确认单元格是日期格式
            cell_value = xldate_as_tuple(worksheet.cell(row,col).value,\
            workbook.datemode)
            #以下*将元组解包成一系列位置参数，通过[:3]来取前三个值，即年月日
            #传入date后，利用strftime将其date类型转换为特定的字符串
            cell_value = date(*cell_value[:3]).strftime('%Y/%m/%d')
          row_of_output.append(cell_value)
        row_of_output.append(os.path.basename(input_file))
        row_of_output.append(worksheet.name)
        if str(worksheet.cell(row,0).value).split('.')[0].strip() in item_to_find:
          filewriter.writerow(row_of_output)
          count_of_items += 1
        line_counter += 1
print('Number of files: {}'.format(file_counter))
print('Number of lines: {}'.format(line_counter))
print('Number of item numbers: {}'.format(count_of_items))
