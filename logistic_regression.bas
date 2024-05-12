Attribute VB_Name = "logistic_regression"
Option Explicit

' The following MS website contains the code for calling Solv (and its arguments) from Solver32.dll
' https://answers.microsoft.com/en-us/msoffice/forum/all/calling-solver-dll-through-vba/e453b8d1-14cc-471f-b740-2c0064bb17bb
' same info
' https://github.com/GCuser99/SolverWrapper/blob/main/src/vba/SolvDLL.cls'

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

' Create a workbook with "data" worksheet and "cover" worksheet
' the first row of "data" worksheet is header e.g., Cell A1 = "y", Cell B1 = "intercept", Cell C1 = "x1", and so on (at least one x)
' Column A must store data of y, column B must be a vector of 1, column C must store data of x1...

Set data_ws = ThisWorkbook.Worksheets("data")
Set cover_ws = ThisWorkbook.Worksheets("cover")

data_ws_lr = data_ws.Cells(cover_ws.Rows.Count, 1).End(xlUp).Row
data_ws_lc = data_ws.Cells(1, cover_ws.Columns.Count).End(xlToLeft).Column

y_range_string = "A2:A" & data_ws_lr
X_range_string = "B2:" + Chr(65 + data_ws_lc - 1) & data_ws_lr

beta_range_string = "B1:B" & (data_ws_lc - 1)

log_likelihood_string = "=SUM(data!" + y_range_string + "*LN(1/(1+EXP(-MMULT(data!" + X_range_string + "," + beta_range_string + _
    "))))+(1-data!" + y_range_string + ")*LN(1 - 1/(1+EXP(-MMULT(data!" + X_range_string + "," + beta_range_string + ")))))"

cover_ws.Range(beta_range_string).Value2 = 0
cover_ws.Range("D1").Formula2 = log_likelihood_string

' I find that the solver parameters/names "solver_xxx" in OpenOffice Calc can also be applied in MS Excel
' https://forum.openoffice.org/en/forum/viewtopic.php?t=100959
' https://learn.microsoft.com/en-us/office/vba/excel/concepts/functions/solverok-function
' By calling the "Solver32.dll" directly, users don't have to install the Solver add-in and reference it manually

With cover_ws.Names
    .Add Name:="solver_adj", RefersTo:="=" & Range(beta_range_string).Address, Visible:=False
    .Add Name:="solver_typ", RefersToLocal:=1, Visible:=False
    .Add Name:="solver_val", RefersToLocal:=0, Visible:=False
    .Add Name:="solver_opt", RefersTo:="=" & Range("cover!$D$1").Address, Visible:=False
    .Add Name:="solver_eng", RefersToLocal:=1, Visible:=False
End With

Solver 0

End Sub

Sub Solver(x As Long)
    Dim dll_loc As String
    
    dll_loc = Application.LibraryPath & Application.PathSeparator & "Solver"
    ChDir (dll_loc)
    ChDrive (dll_loc)
    
    Solv Application, Application, ThisWorkbook, x
End Sub
