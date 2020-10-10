Attribute VB_Name = "splitData"
Option Explicit

Sub splitData_sql()
	Dim cover As Worksheet
	Dim dataSheet As Worksheet
	Dim keyFieldName As String
	Dim lr As Long
	Dim arr As Variant
	Dim dict As Object
	Dim i As Long
	Dim conn As Object
	Dim newWb As Workbook
	Dim key As Variant
	Dim sql As String
	Dim rs As Object
	Dim newWs As Worksheet
	Dim j As Long
	Dim field As Object

	On Error GoTo Handler

	With ThisWorkbook
		Set cover = .Worksheets("cover")
		Set dataSheet = .Worksheets("dataSheet")
	End With

	keyFieldName = dataSheet.Cells(1, 1).Value2

	If keyFieldName = "" Then
		Err.Raise Number:=vbObjectError + 10001, Description:="no data at all"
	End If

	lr = lastRow(dataSheet)
	arr = dataSheet.Range(dataSheet.Cells(2, 1), dataSheet.Cells(lr, 1))
	Set dict = CreateObject("Scripting.Dictionary")

	For i = LBound(arr, 1) To UBound(arr, 1)
		If Not dict.exists(arr(i, 1)) Then
			dict.Add arr(i, 1), vbNullString
		End If
	Next i

	Set conn = CreateObject("Adodb.Connection")
	conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
			"Data Source=" & ThisWorkbook.FullName & ";" & _
			"Extended Properties=""Excel 12.0;HDR=Yes;"";"
	Set newWb = Workbooks.Add

	For Each key In dict.Keys
		sql = "SELECT * FROM [" & dataSheet.Name & "$] WHERE " & keyFieldName & " = '" & key & "';"
		
		Set rs = CreateObject("Adodb.RecordSet")
		rs.Open sql, conn
		
		With newWb
			Set newWs = .Worksheets.Add(After:=.Worksheets(.Worksheets.Count))
		End With
		newWs.Name = key
		newWs.Cells(2, 1).CopyFromRecordset rs
		
		j = 1
		For Each field In rs.Fields
			newWs.Cells(1, j).Value2 = field.Name
			j = j + 1
		Next field
		
		rs.Close
		Set rs = Nothing
	Next key

	conn.Close
	Set conn = Nothing
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
