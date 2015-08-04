import win32com.client
import os
import traceback
import sys

def open(wb_path):
	o = win32com.client.Dispatch("Excel.Application")
	o.Visible = False
	wb = o.Workbooks.Open(wb_path)
	return wb
	
def print_setup(ws, print_area):
	ws.PageSetup.Zoom = False
	ws.PageSetup.FitToPagesTall = 1
	ws.PageSetup.FitToPagesWide = 1
	ws.PageSetup.PrintArea = print_area
	
def print_single_ws(wb, wb_path, ws_index, print_area = None, pdf_path = None):
	if pdf_path is None:
		pdf_path = os.path.dirname(wb_path) + "\\" + os.path.splitext(os.path.basename(wb_path))[0]
		
	if print_area is not None:
		print_setup(ws, print_area)
		
	ws = wb.Worksheets[ws_index]
	ws.ExportAsFixedFormat(0, pdf_path)
	
def print_multiple_ws(wb, wb_path, ws_index_list, print_area = None, pdf_path = None):
	if pdf_path is None:
		pdf_path = os.path.dirname(wb_path) + "\\" + os.path.splitext(os.path.basename(wb_path))[0]
		
	if print_area is not None:
		for index in ws_index_list:
			ws = wb.Worksheets[index-1]
			print_setup(ws, print_area)
	
	wb.WorkSheets(ws_index_list).Select()
	wb.ActiveSheet.ExportAsFixedFormat(0, pdf_path)
	wb.WorkSheets(1).Select()
	
def close(wb):
	wb.Close(True)
	
def manager(wb_path, print_area, ws_index_list):

	wb = open(wb_path)
	try:
		print_multiple_ws(wb, wb_path, ws_index_list, print_area)
	except Exception as e:
		print e
	close(wb)

if __name__ == "__main__":
	wb_path = r'\\kcow00\Jobs3\50790\Bridges\Ratings\Task Order 46\Br 39.02\Rating Calculation\TO46_39.02.xls'
	print_area = 'A1:N65'
	
	wb = open(wb_path)
	try:
		print_single_ws(wb, wb_path, 16, print_area)
	except Exception, e:
		exc_type, exc_value, exc_traceback = sys.exc_info()
		traceback.print_exception(exc_type, exc_value, exc_traceback)
	close(wb)