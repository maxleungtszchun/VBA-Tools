Attribute VB_Name = "userDefinedFunction"
Option Explicit

' https://gist.github.com/maxleungtszchun/05139424c0e9510290cef6f7a04c710f
' distinctcountnoblank() Excel function in the above gist can do the same thing
' it should be faster because it only uses built-in Excel function
' Note COUNTBLANK() counts empty string as blank, which is different from ISBLANK()
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
