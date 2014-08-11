
def set_by_row(sheet, data, start_row = 0, start_col = 0):
	for x in xrange(len(data)):
		row = sheet.row(x + start_row)
		for y in xrange(len(data[x])):
			row.write(y + start_col, data[x][y])
				
def set_by_col(sheet, data, col):
	for x in xrange(len(data)):
		sheet.write(x, col, data[x])
		
def save(book, name):
	book.save(name)
	