import win32com.client
import os
import traceback
import sys
import easygui

def gui():
	msg = ""
	title = "ETP"
	fieldNames = ["Path", "Worksheet(s)", "Print Area(s)", "Paper Size(s)"]
	fieldValues = [".txt or .xlsx", "1,3-5,6,7,11-9", "A1:G57,B4:Z57,A2:G56,A2:G56,A5:J72", "1,2,1,1,2"]
	fieldValues = easygui.multenterbox(msg, title, fieldNames, fieldValues)
	return fieldValues
	
def get_pathlist(path):

	filetype = os.path.splitext(path)[1]
		
	if filetype == ".txt":
		pathlist = [file for file in open(path).readlines]
	elif filetype == ".xlsx" or filetype == ".xls":
		pathlist = path
	else:
		#error
		pass
		
	return pathlist
	
def get_wslist(worksheets):

	try:
		spl = worksheets.split(',')
	except:
		raise()
	
	p1 = re.compile([0-9]+)
	p2 = re.compile([0-9]+[-][0-9]+)
	
	ws_list = []
		
	for ws in spl:
		if "-" not in ws:
			if p1.match(ws) != None:
				ws_list.append(ws)
		elif "-" in ws:
			if p2.match(ws) != None:
				f = ws.split('-')[0]
				l = ws.split('-')[:-1]
				if f < l:
					for n in range(f,l,1):
						ws_list.append(n)
				elif f > l:
					for n in range(l,f,-1):
						ws_list.append(n)
				elif f == l:
					ws_list.append(ws)
				else:
					raise()
		else:
			raise()

def get_palist(print_areas):

	try:
		spl = print_areas.split(',')
	except:
		raise()
		
	p = re.compile(r'[A-z]+[0-9]+[:][A-z]+[0-9]+', re.IGNORECASE)
		
	for pa in spl:
		if p.match(pa) == None:
			raise()
			
	return spl

def get_pslist(page_sizes):
	
	try:
		spl = page_sizes.split(',')
	except:
		raise()
		
	for ps in spl:
		if ps != "1" and ps != "2":
			raise()
		
	return spl
		
def open_wb(wb_path):
	o = win32com.client.Dispatch("Excel.Application")
	o.Visible = False
	wb = o.Workbooks.Open(wb_path)
	return wb
		

	
def print_setup(ws, print_area):
	ws.PageSetup.Zoom = False
	ws.PageSetup.FitToPagesTall = 1
	ws.PageSetup.FitToPagesWide = 1
	ws.PageSetup.PrintArea = print_area
	
	#ws.PageSetup.XlPaper11x17
	#ws.PageSetup.XlPaperLetter
	
	#ws.PageSetup.XlLandscape
	#ws.PageSetup.XlPortrait
	
# def print_single_ws(wb, wb_path, ws_index, print_area = None, pdf_path = None):
	# if pdf_path is None:
		# pdf_path = os.path.dirname(wb_path) + "\\" + os.path.splitext(os.path.basename(wb_path))[0]
		
	# if print_area is not None:
		# print_setup(ws, print_area)
		
	# ws = wb.Worksheets[ws_index]
	# ws.ExportAsFixedFormat(0, pdf_path)
	
def print_ws(wb, wb_path, ws_index_list, print_area_list, pdf_path = None):
	if pdf_path is None:
		pdf_path = os.path.dirname(wb_path) + "\\" + os.path.splitext(os.path.basename(wb_path))[0]
		
	if print_area is not None:
		for index in ws_index_list:
			ws = wb.Worksheets[index-1]
			print_setup(ws, print_area, page_size)
	
	wb.WorkSheets(ws_index_list).Select()
	wb.ActiveSheet.ExportAsFixedFormat(0, pdf_path)
	wb.WorkSheets(1).Select()
	
def close(wb):
	wb.Close(True)
	
def manager():

	input_values = gui()
	path = input_values[0]
	worksheets = input_values[1]
	print_areas = input_values[2]
	page_sizes = input_values[3]
	
	pathlist = get_pathlist(path)
	ws_list = get_wslist(worksheets)
	pa_list = get_palist(print_areas)
	ps_list = get_pslist(page_sizes)	
	
	for path in pathlist:
		wb = open_wb(path)
		print_ws(wb, path, ws_list, pa_list, ps_list)
	

if __name__ == "__main__":

	manager()
	
"""


"""