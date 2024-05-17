Attribute VB_Name = "bootstrapLinearRegression"
Option Explicit

Sub run_linear_regression( _
    y_range As Variant, _
    x_range As Variant, _
    Optional intercept As Variant, _
    Optional header As Variant, _
    Optional interval As Variant, _
    Optional output As Variant _
)
    ' the registered function "fnRegress" in "ANALYS32.XLL" and its arguments can be discovered by using Application.RegisteredFunctions
    ' https://stackoverflow.com/questions/1576399/methods-in-excel-addin-xll
    ' https://learn.microsoft.com/en-us/office/vba/api/excel.application.registeredfunctions
    ' by directly calling "fnRegress" in "ANALYS32.XLL", users don't need to install the add-in manually and can use it via VBA

    Dim xll_loc As String, reg_str As String, q As String

    xll_loc = Application.LibraryPath & Application.PathSeparator & "Analysis" & Application.PathSeparator & "ANALYS32.XLL"
    reg_str = "REGISTER.ID(""" & xll_loc & """,""fnRegress"")"
    Application.Run ExecuteExcel4Macro(reg_str), y_range, x_range, intercept, header, interval, output

End Sub

Sub bootstrap_linear_regression()

    ' create a workbook with "data" and "cover" worksheets
    ' the first column in "data" worksheet is for y variable (with header)
    ' the second column and so on in "data" worksheet are for x variables (with header, max. 16 x variables)

    Dim i As Long, j As Long, b_iter As Long, data_ws_lr As Long, data_ws_lc As Long
    Dim data_ws As Worksheet, cover_ws As Worksheet
    Dim total_beta As Double, total_t As Double, avg_beta As Double, sum_square As Double
    Dim beta() As Double, t() As Double

    On Error GoTo Handler

    Set data_ws = ThisWorkbook.Worksheets("data")
    Set cover_ws = ThisWorkbook.Worksheets("cover")

    b_iter = CInt(InputBox("Specify number of bootstrap iteration, suggest > 50"))
    If b_iter < 2 Then Err.Raise Number:=vbObjectError + 10001, Description:="cannot less than 2"

    With data_ws
        data_ws_lr = .Cells(.Rows.Count, 1).End(xlUp).Row
        data_ws_lc = .Cells(1, .Columns.Count).End(xlToLeft).Column
    End With

    ReDim beta(1 To data_ws_lc, 1 To b_iter)
    ReDim t(1 To data_ws_lc, 1 To b_iter)

    With cover_ws
        .Columns.Delete
        .Range("A2").Formula2 = "=RANDARRAY(" & data_ws_lr - 1 & ",1,1," & data_ws_lr - 1 & ",TRUE)"
        .Range("B2").Formula2 = "=INDEX(data!A2:" & Chr(64 + data_ws_lc) & data_ws_lr & ",A2:A" & data_ws_lr & ",SEQUENCE(1," & data_ws_lc & "))"
        .Range("B1:" & Chr(64 + data_ws_lc + 1) & "1").Value2 = data_ws.Range("A1:" & Chr(64 + data_ws_lc) & "1").Value2
    End With

    For j = 1 To b_iter
        cover_ws.Columns(Chr(64 + data_ws_lc + 1 + 2) & ":" & Chr(64 + data_ws_lc + 1 + 2 + 6)).Delete
        cover_ws.Calculate ' the above line should recal the sheet if autocal is on. just to be safe
        run_linear_regression cover_ws.Range("B1:B" & data_ws_lr), cover_ws.Range("C1:" & Chr(64 + data_ws_lc + 1) & data_ws_lr), _
            header:=1, output:=cover_ws.Range(Chr(64 + data_ws_lc + 1 + 2) & "1")

        For i = 1 To data_ws_lc
            beta(i, j) = cover_ws.Range(Chr(64 + data_ws_lc + 1 + 3) & 16 + i).Value2
            t(i, j) = cover_ws.Range(Chr(64 + data_ws_lc + 1 + 5) & 16 + i).Value2
        Next i
    Next j

    cover_ws.Columns.Delete
    run_linear_regression data_ws.Range("A1:A" & data_ws_lr), data_ws.Range("B1:" & Chr(64 + data_ws_lc) & data_ws_lr), header:=1, output:=cover_ws.Range("A1")

    For i = 1 To data_ws_lc
        total_beta = 0
        For j = 1 To b_iter
            total_beta = total_beta + beta(i, j)
        Next j
        avg_beta = total_beta / b_iter
        cover_ws.Range("I" & i + 16).Value2 = avg_beta

        sum_square = 0
        For j = 1 To b_iter
            sum_square = sum_square + (beta(i, j) - avg_beta) ^ 2
        Next j
        cover_ws.Range("J" & i + 16).Value2 = (sum_square / (b_iter - 1)) ^ 0.5
    Next i

    For i = 1 To data_ws_lc
        total_t = 0
        For j = 1 To b_iter
            total_t = total_t + t(i, j)
        Next j
        cover_ws.Range("K" & i + 16).Value2 = total_t / b_iter
    Next i

    With cover_ws
        .Range("I16").Value2 = "Bootstrapped Coeff"
        .Range("J16").Value2 = "Bootstrapped SE"
        .Range("K16").Value2 = "Bootstrapped t Stat"
        .Cells.NumberFormat = "0.000"
        .Range("B8").NumberFormat = "0"
        .Range("B12:B14").NumberFormat = "0"

        .Range("B16:B" & 16 + data_ws_lc).Copy
        .Range("I16").CurrentRegion.Select
        Selection.PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False
        .Columns.AutoFit
        .Range("A1").Select
    End With

    Windows(ThisWorkbook.Name).DisplayGridlines = False
Exit Sub

Handler:
    MsgBox Err.Description & ", " & Err.Number
End Sub
