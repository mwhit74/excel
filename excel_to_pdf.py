import win32com.client
import os
import traceback
import sys

def open(wb_path):
	o = win32com.client.Dispatch("Excel.Application")
	o.Visible = False
	wb = o.Workbooks.Open(wb_path)
	return o, wb
	
def print_setup(ws, print_area):
	ws.PageSetup.Zoom = False
	ws.PageSetup.FitToPagesTall = 1
	ws.PageSetup.FitToPagesWide = 1
	ws.PageSetup.PrintArea = print_area
	
def print_single_ws(wb, wb_path, ws_index, print_area, pdf_path = None):
	if pdf_path is None:
		pdf_path = os.path.dirname(wb_path) + "\\" + os.path.splitext(os.path.basename(wb_path))[0]
		
	ws = wb.Worksheets[ws_index]
	print_setup(ws, print_area)
	ws.ExportAsFixedFormat(0, pdf_path)
	
def print_multiple_ws(wb, wb_path, ws_index_list, print_area, pdf_path = None):
	if pdf_path is None:
		pdf_path = os.path.dirname(wb_path) + "\\" + os.path.splitext(os.path.basename(wb_path))[0]

	for index in ws_index_list:
		ws = wb.Worksheets[index]
		print_setup(ws, print_area)
	
	wb.WorkSheets(ws_index_list).Select()
	wb.ActiveSheet.ExportAsFixedFormat(0, pdf_path)
	
def close(o, wb):
	wb.Close(True)
	o.Quit()

if __name__ == "__main__":
	wb_path = r'\\kcow00\Jobs3\50790\Bridges\Ratings\Task Order 46\zScratch\MLW\TO46_Template.xls'
	print_area = 'A1:N65'
	
	o, wb = open(wb_path)
	try:
		print_multiple_ws(wb, wb_path, [1,4,5], print_area)
	except Exception, e:
		exc_type, exc_value, exc_traceback = sys.exc_info()
		traceback.print_exception(exc_type, exc_value, exc_traceback)
	close(o, wb)