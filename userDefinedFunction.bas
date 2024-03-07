Attribute VB_Name = "userDefinedFunction"
Option Explicit

' =COUNTA(UNIQUE(IF(ISBLANK(yourRange),"",yourRange)))-IF(COUNTBLANK(yourRange)>0,1,0)
' the above Excel function can do the same thing for the column case
' COUNTBLANK() counts empty string as blank, which is different from ISBLANK()
Function distinctCountNoBlank(rng As Range, dimension As String) As Long
    Application.Volatile
    Dim arr As Variant

    If Application.WorksheetFunction.Concat(rng) = "" Then
        distinctCountNoBlank = 0
        Exit Function
    End If

    Select Case dimension
        Case "column"
            arr = Application.WorksheetFunction.Unique(rng) ' the function returns 2d column array index starts from 1
            arr = twoD2oneD(arr, dimension) ' the function returns 1d array index starts from 0
        Case "row"
            arr = Application.WorksheetFunction.Unique(rng, True, False) ' the function returns 1d array index starts from 1
        Case Else
            distinctCountNoBlank = CVErr(xlErrValue)
    End Select

    arr = noBlankOrEmptyStrInArr(arr) ' the function returns 1d array index starts from 0
    distinctCountNoBlank = UBound(arr) + 1
End Function
