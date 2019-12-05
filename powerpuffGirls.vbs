Set obj = createobject("Excel.Application")

obj.visible = True

sourcePath ="C:\Users\s.sarvesh.basutkar\Desktop\Invoices_exception.xlsx"

set sourceWorkBook = obj.Workbooks.open(sourcePath)

set sourceWorksheet = sourceWorkBook.Worksheets("Invoices")

'Removes all the applied filters from the sheet
Sub removeFilter()
	If sourceWorksheet.AutoFilterMode = True Then
		sourceWorksheet.ShowAllData
	End If
End Sub

'Copy/Pastes data from Invoices to other sheets based on filters
Sub filterData(col, sheetName, colNum, condition)
	'Applies filter on the sheet based on condition
	With sourceWorksheet
		.Range(col).AutoFilter colNum, condition
	End With
	sourceWorkBook.Sheets("Invoices").UsedRange.Copy
	sourceWorkBook.Sheets(sheetName).Select
	sourceWorkBook.Sheets(sheetName).Range("A1").Select
	sourceWorkBook.Sheets(sheetName).Paste
End Sub

'appends the data from every sheet to "Output Invoices" sheet 
Sub appendData(sheetName)
	nUsedRows = sourceWorkBook.Worksheets(sheetName).UsedRange.Rows.Count
	sourceWorkBook.Worksheets(sheetName).Range("A2:S"&nUsedRows).Copy
	sourceWorkBook.Worksheets("Output Invoices").Range(st).PasteSpecial Paste =xlValues
	rCount = sourceWorkBook.Worksheets("Output Invoices").UsedRange.rows.count
	rCount = rCount + 1
	st = "A"&rCount
End Sub

call filterData("K1","Exceptions", 11, "=0")
call removeFilter()
call filterData("Q1", "ROAD", 17, "=ROAD")
call removeFilter()
call filterData("Q1","PRIORITY", 17, "=PRIORITY")
call removeFilter()
call filterData("Q1", "CTNLOCRED", 17, "=CTNLOCRED")
call removeFilter()
call filterData("Q1", "RTNCGC", 17, "=RTNCGC")
call removeFilter()
call filterData("Q1", "NXTFL", 17, "=NXTFL")
call removeFilter()

st = "A1"
sourceWorkBook.Worksheets("ROAD").UsedRange.Copy
sourceWorkBook.Worksheets("Output Invoices").Range(st).PasteSpecial Paste =xlValues
rCount = sourceWorkBook.Worksheets("Output Invoices").UsedRange.rows.count
rCount = rCount + 1
st = "A"&rCount

call appendData("PRIORITY")
call appendData("CTNLOCRED")
call appendData("RTNCGC")
call appendData("NXTFL")

'Deletes the rows copied to "Exception" sheet from the "Invoices" sheet 
nUsedCols = sourceWorksheet.UsedRange.Columns.Count
nUsedRows = sourceWorksheet.UsedRange.Rows.Count
For j=1 to nUsedCols	
    If sourceWorkSheet.cells(1, j).value = "Item" Then		
        For i=1 to nUsedRows			
            If sourceWorkSheet.Cells(i, j).value = "0" Then
                'converts the integer i to string
                CStr(i)               
                Set objRange = sourceWorkSheet.Range("A" & i).EntireRow			
                objRange.Delete				
				i=i-1			
            End If		
        Next	
    End If	
Next