# -*- encoding: utf-8 -*-
import re
import sys
import os
import xlrd
import xlwt

writeBook = xlwt.Workbook()
writeSheet = writeBook.add_sheet('union')
pattern = re.compile('^.*\.xlsx')
root_path = sys.argv[1]
row_num_sum = 0
for file_path in os.listdir(root_path):
    if pattern.match(file_path):
        readBook = xlrd.open_workbook(os.path.join(root_path, file_path))
        readSheet = readBook.sheet_by_index(0)
        for row_num in xrange(1, readSheet.nrows):
            row = readSheet.row_values(row_num)
            for col, cell in enumerate(row):
                writeSheet.write(row_num_sum, col, cell)
            row_num_sum += 1
writeBook.save('union.xlsx')
