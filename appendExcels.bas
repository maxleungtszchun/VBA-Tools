Attribute VB_Name = "appendExcels"
Option Explicit

Sub appendExcels()
	Dim cover As Worksheet
	Dim appended As Worksheet
	Dim path As String

	Dim fso As Object
	Dim folder As Object
	Dim file As Object

	Dim i As Long
	Dim wb As Workbook
	Dim ws As Worksheet
	Dim lc As Long
	Dim firstHeader As Variant
	Dim header As Variant
	Dim j As Long
	Dim appended_lr As Long
	Dim k As Long

	On Error GoTo Handler

	Set cover = ThisWorkbook.Worksheets("cover")
	Set appended = ThisWorkbook.Worksheets("appended")

	appended.Cells.Delete

	path = cover.Cells(1, 1).Value2
	If path = "" Then
		Err.Raise Number:=vbObjectError + 10001, Description:="no path"
	End If

	Set fso = CreateObject("Scripting.FileSystemObject") 'late binding

	If Not fso.folderexists(path) Then
		Err.Raise Number:=vbObjectError + 10002, Description:="no folder"
	End If

	Set folder = fso.getfolder(path)

	i = 1
	For Each file In folder.Files
		Set wb = Workbooks.Open(Filename:=path & "\" & file.Name)
		Set ws = wb.Worksheets(1)
		
		lc = lastCol(ws)
		header = ws.Range(ws.Cells(1, 1), ws.Cells(1, lc))
		
		If i = 1 Then
			firstHeader = header
		Else
			If UBound(firstHeader, 2) <> UBound(header, 2) Then
				Err.Raise Number:=vbObjectError + 10003, Description:="not same header"
			End If
			
			For j = LBound(firstHeader, 2) To UBound(firstHeader, 2)
				 If firstHeader(1, j) <> header(1, j) Then
					Err.Raise Number:=vbObjectError + 10003, Description:="not same header"
				 End If
			Next j
		End If
		
		ws.Cells(1, 1).CurrentRegion.Copy
		appended_lr = lastRow(appended)
		appended.Cells(appended_lr + 1, 1).PasteSpecial Paste:=xlPasteValues
		wb.Close
		
		i = i + 1
	Next file

	appended.Cells(1, 1).AutoFilter Field:=1, Criteria1:=firstHeader(1, 1)
	appended.AutoFilter.Range.SpecialCells(xlCellTypeVisible).EntireRow.Delete

	appended.Rows(1).Insert Shift:=xlShiftDown
	For k = LBound(firstHeader, 2) To UBound(firstHeader, 2)
		appended.Cells(1, k).Value2 = firstHeader(1, k)
	Next k

	appended.Cells(1, 1).Select

	Exit Sub

	Handler:
		MsgBox Err.Description & ", " & Err.Number
End Sub

Function lastRow(ws) As Long
    Dim lr As Long
    
    lr = ws.Cells(ws.Cells.Rows.Count, 1).End(xlUp).Row
    If lr = 1 And ws.Cells(1, 1).Value2 = "" Then
        lr = 0
    End If
    
    lastRow = lr
End Function

Function lastCol(ws) As Long
    lastCol = ws.Cells(1, ws.Cells.Columns.Count).End(xlToLeft).Column
End Function
