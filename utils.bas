Attribute VB_Name = "utils"
Option Explicit

Public Function getUniqueValue(arr As Variant) As Variant
    Dim i As Long
    Dim dict As Object
    Dim new_arr As Variant
    Dim j As Long
    Dim key As Variant
    
    Set dict = CreateObject("scripting.dictionary")
    
    j = 0
    For i = LBound(arr) To UBound(arr)
        If Not dict.exists(arr(i)) Then
            dict.Add arr(i), vbNullString
            j = j + 1
        End If
    Next i
    
    ReDim new_arr(0 To j - 1)
    i = 0
    For Each key In dict.keys
        new_arr(i) = key
        i = i + 1
    Next key
    
    getUniqueValue = new_arr
End Function

Public Function getUniqueValueNoBlank(arr As Variant) As Variant
    arr = getUniqueValue(arr)
    arr = noBlankOrEmptyStrInArr(arr)
    getUniqueValueNoBlank = arr
End Function

Public Function noBlankOrEmptyStrInArr(arr As Variant)
    Dim i As Long, j As Long
    Dim new_arr As Variant
    
    j = 0
    For i = LBound(arr) To UBound(arr)
        If arr(i) = "" Then
            j = j + 1
        End If
    Next i
    
    ReDim new_arr(0 To UBound(arr) - LBound(arr) + 1 - j - 1)
    
    j = 0
    For i = LBound(arr) To UBound(arr)
        If arr(i) <> "" Then
            new_arr(j) = arr(i)
            j = j + 1
        End If
    Next i
    
    noBlankOrEmptyStrInArr = new_arr
End Function

Public Function twoD2oneD(twoD_array As Variant, dimension) As Variant
    Dim d As Integer
    Dim new_array As Variant
    Dim i As Long
    
    Select Case dimension
        Case "row"
            d = 2
        Case "column"
            d = 1
        Case Else
            Err.Raise Number:=vbObjectError + 666, Description:="only row or column are accepted for dimension"
    End Select
    
    If d = 2 Then
        If Not UBound(twoD_array, 1) = 1 Or Not LBound(twoD_array, 1) = 1 Then
            Err.Raise Number:=vbObjectError + 777, Description:="it is not a row array"
        End If
    Else
        If Not UBound(twoD_array, 2) = 1 Or Not LBound(twoD_array, 2) = 1 Then
            Err.Raise Number:=vbObjectError + 888, Description:="it is not a column array"
        End If
    End If
    
    ReDim new_array(0 To UBound(twoD_array, d) - 1)
    
    For i = LBound(twoD_array, d) To UBound(twoD_array, d)
        If d = 2 Then
            new_array(i - 1) = twoD_array(1, i)
        Else
            new_array(i - 1) = twoD_array(i, 1)
        End If
    Next i
    
    twoD2oneD = new_array 'return array dimension starts from 0
End Function

Public Function getColIndex(ws As Worksheet, colName As String, Optional headerRow As Long = 1) As Long
    Dim lc As Long
    Dim arr As Variant
    Dim i As Long
    
    lc = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    arr = ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, lc))
    
    For i = LBound(arr, 2) To UBound(arr, 2)
        If arr(1, i) = colName Then
            getColIndex = i
            Exit For
        End If
    Next i
    
    If getColIndex = 0 Then
        Err.Raise Number:=vbObjectError + 999, Description:="colName not found"
    End If
End Function

Public Function getLastRow(ws As Worksheet) As Long
    On Error Resume Next
    getLastRow = ws.Cells.Find(What:="*", _
                                After:=ws.Range("A1"), _
                                Lookat:=xlPart, _
                                LookIn:=xlFormulas, _
                                SearchOrder:=xlByRows, _
                                SearchDirection:=xlPrevious, _
                                MatchCase:=False).Row
    On Error GoTo 0
End Function

Public Function getLastColumn(ws As Worksheet) As Long
    On Error Resume Next
    getLastColumn = ws.Cells.Find(What:="*", _
                                    After:=ws.Range("A1"), _
                                    Lookat:=xlPart, _
                                    LookIn:=xlFormulas, _
                                    SearchOrder:=xlByColumns, _
                                    SearchDirection:=xlPrevious, _
                                    MatchCase:=False).Column
    On Error GoTo 0
End Function
