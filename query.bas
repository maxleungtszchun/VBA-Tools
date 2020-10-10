Attribute VB_Name = "query"
Option Explicit 

Function Query(tbl As ListObject, sql1 As String, Optional sql2 As String = "") As Variant
    Dim conn As Object
    Dim sql As String
    Dim rs As Object
    Dim i As Long, j As Long
    Dim matrix As Variant
    
    Set conn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
            "Data Source=" & ThisWorkbook.FullName & ";" & _
            "Extended Properties=""Excel 12.0;HDR=Yes;"";"
    
    sql = sql1 & " FROM " & getName(tbl) & sql2 & ";"
    
    With rs
        .ActiveConnection = conn
        .Source = sql
        .CursorLocation = 3
        .LockType = 3
        .Open
        .MoveLast
        .MoveFirst
        ReDim matrix(1 To .RecordCount + 1, 1 To .Fields.Count)
        For j = 1 To .Fields.Count
            matrix(1, j) = .Fields(j - 1).Name
        Next j
        For i = 2 To .RecordCount + 1
            For j = 1 To .Fields.Count
                  matrix(i, j) = .Fields(j - 1).Value
            Next j
            .MoveNext
        Next i
    End With
    
    rs.Close
    conn.Close
    
    Set rs = Nothing
    Set conn = Nothing
    
    Query = matrix 'return matrix dimension starts from 1
End Function

Private Function getName(table As ListObject) As String
    getName = "[" & table.Parent.Name & "$" & table.Range.Address(False, False) & "] AS [" & table.Name & "]"
End Function

Sub PasteFromMatrix(matrix As Variant, targetCell As Range)
    Dim matrixRange As Range
    Set matrixRange = targetCell.Resize(UBound(matrix, 1), UBound(matrix, 2))
    matrixRange.Value2 = matrix
End Sub