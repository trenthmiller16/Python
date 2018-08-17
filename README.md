import schedule
import time
import win32com.client
# Start an instance of Excel
xlapp = win32com.client.DispatchEx("Excel.Application")

# Open the workbook in said instance of Excel
wb = xlapp.workbooks.open("I:\\share\\OPS_ANALYSTS\\Excel Tools\\Daily SLA Report\\DailyOpenCaseSLAReport_Updated.xlsx")


wb.RefreshAll()
time.sleep(300)
count = wb.Sheets.Count
for i in range(count):
  ws = wb.Worksheets[i]
  pivotCount = ws.PivotTables().Count
  for j in range(1, pivotCount+1):
    ws.PivotTables(j).PivotCache().Refresh()

wb.Save()


# Quit
xlapp.Quit()
