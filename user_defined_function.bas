Attribute VB_Name = "user_defined_function"
Option Explicit

Function distinctCountNoBlank(rng As Range, dimension As String) As Long
    Application.Volatile
    Dim arr As Variant
    
    If Application.WorksheetFunction.Concat(rng) = "" Then
        distinctCountNoBlank = 0
        Exit Function
    End If
    
    Select Case dimension
        Case "column"
            arr = Application.WorksheetFunction.Unique(rng)
            arr = twoD2oneD(arr, dimension)
        Case "row"
            arr = Application.WorksheetFunction.Unique(rng, True, False)
        Case Else
            distinctCountNoBlank = CVErr(xlErrValue)
    End Select
    
    arr = noBlankOrEmptyStrInArr(arr)
    distinctCountNoBlank = UBound(arr) + 1
End Function
