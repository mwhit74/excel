
#For get_loadcases() function
import sys
sys.path.append(r'C:\Users\mwhitten\brgusers\Python\loads')
from force import Force
from loadcase import LoadCase
from xlrd import open_workbook

def get_by_row(sheet, start_row, start_col, end_row = None, 
				end_col = None):
	if end_row == None:
		end_row = sheet.nrows
	if end_col == None:
		end_col = sheet.ncols
	data_set = []
	for row in xrange(start_row, end_row):
		row_list = []
		for col in xrange(start_col, end_col):
			val = sheet.cell(row, col).value
			if isinstance(val, float):
				row_list.append(val)
			else:
				row_list.append(val.encode('utf-8'))
		data_set.append(row_list)
	return data_set
	
def get_by_col(sheet, start_row, start_col):
	data_set = []
	for col in xrange(sheet.ncols - start_col):
		col_list = []
		for row in xrange(sheet.nrows - start_row):
			col_list.append(sheet.cell(start_row + row, start_col + col).value)
		data_set.append(row_list)
	return data_set

def get_column(sheet, start_row, end_row, column):
	data = []
	for row in xrange(start_row, end_row):
		data.append(get_individual_cell(sheet, row, column))
	return data
		

def get_individual_cell(sheet, row, col):
	val = sheet.cell(row, col).value
	if isinstance(val, float):
		return val
	else:
		return val.encode('utf-8')

def get_loadcases(sheet, start_cell_row, start_cell_col):
	counter = start_cell_row

	load_cases = []
	
	#first cell in set to read
	#initializes the while loop
	cell_name = sheet.cell(start_cell_row, start_cell_col)
	
	while cell_name.value != '':
		
		#clean up temporary arrays after each loop so the entries
		# don't compound
		forces = []
		
		forces.append(sheet.cell(counter, start_cell_col + 1).value)
		forces.append(sheet.cell(counter, start_cell_col + 2).value)
		forces.append(sheet.cell(counter, start_cell_col + 3).value)
		forces.append(sheet.cell(counter, start_cell_col + 4).value)
		forces.append(sheet.cell(counter, start_cell_col + 5).value)
		forces.append(sheet.cell(counter, start_cell_col + 6).value)
		
		new_force = Force(forces = forces)
		new_load_case = LoadCase(name = cell_name.value, forces = new_force)
		load_cases.append(new_load_case)
		
		#for testing purposes
		#if counter == start_cell_row:
		#	print_loads(cell_name, sheet, counter, start_cell_col)
		
		counter = counter + 1
		
		cell_name = sheet.cell(counter, start_cell_col)
		#for testing purposes
		#print_loads(cell_name, sheet, counter, start_cell_col)
		
		
	return load_cases