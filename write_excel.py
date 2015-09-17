
from xlwt import *

#formatting excel
#http://www.youlikeprogramming.com/2011/04/examples-generating-excel-documents-using-pythons-xlwt/
#91px = 3333 units = 1 inch

def set_by_row(sheet, data, start_row = 0, start_col = 0, s_type=XFStyle()):
	for x in xrange(len(data)):
		row = sheet.row(x + start_row)
		for y in xrange(len(data[x])):
			row.write(y + start_col, data[x][y], style=s_type)
		
	return x, y
			
def set_row(sheet, line, row = 0, start_col = 0, s_type=XFStyle()):
	for x in xrange(len(line)):
		sheet.write(row, start_col + x, line[x], style=s_type)
				
def set_by_col(sheet, data, col = 0, s_type=XFStyle(), start_row = 0):
	for x in xrange(len(data)):
		sheet.write(x + start_row, col, data[x], style=s_type)
		
def set_individual_cell(sheet, val, row = 0, col = 0, s_type=XFStyle()):
	sheet.write(row, col, label=val, style=s_type)
		
def save(book, name):
	book.save(name)
	