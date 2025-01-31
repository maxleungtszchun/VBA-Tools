Attribute VB_Name = "logisticRegression"
Option Explicit

' The following MS website contains the code for calling Solv (and its arguments) from Solver32.dll
' https://answers.microsoft.com/en-us/msoffice/forum/all/calling-solver-dll-through-vba/e453b8d1-14cc-471f-b740-2c0064bb17bb
' same info
' https://github.com/GCuser99/SolverWrapper/blob/main/src/vba/SolvDLL.cls

#If VBA7 Then
Private Declare PtrSafe Function Solv Lib "Solver32.dll" (ByVal obj, ByVal obj, ByVal work_book, ByVal x As Long) As Long
#Else
Private Declare Function Solv Lib "Solver32.dll" (ByVal obj, ByVal obj, ByVal work_book, ByVal x As Long) As Long
#End If

Sub logistic_regression()

    Dim data_ws As Worksheet
    Dim cover_ws As Worksheet
    Dim data_ws_lr As Long
    Dim data_ws_lc As Long
    Dim y_range_string As String
    Dim X_range_string As String
    Dim beta_range_string As String
    Dim log_likelihood_string As String

    ' create a workbook with "data" worksheet and "cover" worksheet
    ' the first column in "data" worksheet is for y variable (with header)
    ' the second column in "data" worksheet is for intercept i.e., a column of 1 (with header)
    ' the third column and so on in "data" worksheet are for other x variables (with header)

    ' Note that the solutions of Solver in Excel can be inaccurate if there are many x variables

    Set data_ws = ThisWorkbook.Worksheets("data")
    Set cover_ws = ThisWorkbook.Worksheets("cover")

    data_ws_lr = data_ws.Cells(cover_ws.Rows.Count, 1).End(xlUp).Row
    data_ws_lc = data_ws.Cells(1, cover_ws.Columns.Count).End(xlToLeft).Column

    y_range_string = "A2:A" & data_ws_lr
    X_range_string = "B2:" & Chr(64 + data_ws_lc) & data_ws_lr

    beta_range_string = "B2:B" & data_ws_lc - 1 + 1

    ' a very small number 1E-10 is added to prevent the case LN(0)
    log_likelihood_string = "=SUM(data!" & y_range_string & "*LN(1/(1+EXP(-MMULT(data!" & X_range_string & "," & beta_range_string & _
        "))))+(1-data!" & y_range_string & ")*LN(1/(1+EXP(MMULT(data!" & X_range_string & "," & beta_range_string & ")))))"

    With cover_ws
        .Columns.Delete
        .Range("B1").Value2 = "Coefficients"
        .Range("A2:A" & data_ws_lc - 1 + 1).Value2 = Application.WorksheetFunction.Transpose(data_ws.Range("B1:" & Chr(64 + data_ws_lc) & "1").Value2)
        .Range(beta_range_string).Value2 = 0
        .Range("D1").Value2 = "Ln Likelihood Value"
        .Range("D2").Formula2 = log_likelihood_string
    End With

    ' I find that the solver parameters/names "solver_xxx" in OpenOffice Calc can also be applied in MS Excel
    ' https://forum.openoffice.org/en/forum/viewtopic.php?t=100959
    ' https://learn.microsoft.com/en-us/office/vba/excel/concepts/functions/solverok-function
    ' By calling the "Solver32.dll" directly, users don't have to install the Solver add-in and reference it manually

    With cover_ws.Names
        .Add Name:="solver_adj", RefersTo:="=" & Range(beta_range_string).Address, Visible:=False
        .Add Name:="solver_typ", RefersToLocal:=1, Visible:=False
        .Add Name:="solver_val", RefersToLocal:=0, Visible:=False
        .Add Name:="solver_opt", RefersTo:="=" & Range("D2").Address, Visible:=False
        .Add Name:="solver_eng", RefersToLocal:=1, Visible:=False
        .Add Name:="solver_scl", RefersToLocal:=1, Visible:=False
        .Add Name:="solver_neg", RefersToLocal:=2, Visible:=False
        .Add Name:="solver_ssz", RefersToLocal:=100000, Visible:=False
        .Add Name:="solver_lin", RefersToLocal:=2, Visible:=False
        .Add Name:="solver_itr", RefersTo:=2147483647#, Visible:=False
    End With

    cover_ws.Activate
    Solver 0

    With cover_ws
        .Cells.NumberFormat = "0.000"
        .Columns.AutoFit
        .Range("A1").Select
    End With

    Windows(ThisWorkbook.Name).DisplayGridlines = False

End Sub

Sub Solver(x As Long)
    Dim dll_loc As String
    
    dll_loc = Application.LibraryPath & Application.PathSeparator & "Solver"
    ChDir (dll_loc)
    ChDrive (dll_loc)
    
    Solv Application, Application, ThisWorkbook, x
End Sub
