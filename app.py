import os
import glob
import re
import sys
import xlsxwriter

try:  
  # loop through txt files in folder
  for file_path in glob.glob(os.path.join('data/', '*.txt')):
    with open(file_path,'r') as f:
      file_name = os.path.basename(file_path)
      file_name_wo_ext = os.path.splitext(file_name)[0]
      
      # open file and read first row
      first_row = f.readline()

      # remove \n character
      first_row = first_row.strip()

      # populate list with column headers
      header_list = first_row.split('\t')
      
      ws_lnd_header = (['Source','DB Name','Table Name','Column Name','Data Type'])
      ws_raw_header = (['Source','DB Name','Table Name','Column Name','Data Type'])
      ws_lnd_data = []
      ws_raw_data = []
      lnd_source = ''
      lnd_db_name = 'lnd_db'
      lnd_table_name = ''
      raw_source = ''
      raw_db_name = 'raw_db'
      raw_table_name = ''
      
      # loop through column header
      for idx, val in enumerate(header_list):
        if re.search('_num', val):
          ws_lnd_data.append([file_name, lnd_db_name, file_name_wo_ext, val, ' numeric(13,2)'])
          ws_raw_data.append([file_name, raw_db_name, file_name_wo_ext, val, ' numeric(13,2)'])
        elif re.search('_int', val):
          ws_lnd_data.append([file_name, lnd_db_name, file_name_wo_ext, val, ' int'])
          ws_raw_data.append([file_name, raw_db_name, file_name_wo_ext, val, ' int'])
        elif re.search('_dt', val):
          ws_lnd_data.append([file_name, lnd_db_name, file_name_wo_ext, val, ' date'])
          ws_raw_data.append([file_name, raw_db_name, file_name_wo_ext, val, ' date'])
        else:
          ws_lnd_data.append([file_name, lnd_db_name, file_name_wo_ext, val, ' text'])
          ws_raw_data.append([file_name, raw_db_name, file_name_wo_ext, val, ' text'])
      
      # create xlsx workbook, worksheets
      wb = xlsxwriter.Workbook(file_name_wo_ext + '.xlsx')
      ws_lnd = wb.add_worksheet('LND')
      ws_raw = wb.add_worksheet('RAW')

      # create header in lnd
      row = 0
      col = 0
      
      for i in (ws_lnd_header):
        ws_lnd.write(row,col,i)
        col += 1
      
      # write lnd data      
      for row_num, row_data in enumerate(ws_lnd_data):
        for col_num, col_data in enumerate(row_data):
          ws_lnd.write(row_num, col_num, col_data)
      
      # create header in raw
      row = 0
      col = 0
      
      for i in (ws_raw_header):
        ws_raw.write(row,col,i)
        col += 1
      
      # write raw data      
      for row_num, row_data in enumerate(ws_raw_data):
        for col_num, col_data in enumerate(row_data):
          ws_raw.write(row_num, col_num, col_data)
      
      wb.close()
      
except: 
  print("An error occurred...") 
  wb.close()

  sys.exit(1) 

finally:
  sys.exit(0)