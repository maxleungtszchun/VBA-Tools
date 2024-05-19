Attribute VB_Name = "bootstrapLinearRegressionFast"
Option Explicit

Sub bootstrap_linear_regression_fast()

    ' this version uses Application.WorksheetFunction.LinEst(), which returns result in matrix
    ' thus it is faster
    ' moreover, the number of x variables can be larger than 16

    ' create a workbook with "data" and "cover" worksheets
    ' the first column in "data" worksheet is for y variable (with header)
    ' the second column and so on in "data" worksheet are for x variables (with header)

    Dim i As Long, j As Long, b_iter As Long, data_ws_lr As Long, data_ws_lc As Long
    Dim data_ws As Worksheet, cover_ws As Worksheet
    Dim total_beta As Double, total_t As Double, avg_beta As Double, sum_square As Double
    Dim beta() As Double, t() As Double
    Dim boot_result As Variant, boot_y As Variant, boot_x As Variant

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
        cover_ws.Calculate
        boot_y = cover_ws.Range("B2:B" & data_ws_lr)
        boot_x = cover_ws.Range("C2:" & Chr(64 + data_ws_lc + 1) & data_ws_lr)
        boot_result = Application.WorksheetFunction.LinEst(boot_y, boot_x, True, True)

        For i = 0 To data_ws_lc - 1
            beta(i + 1, j) = boot_result(1, data_ws_lc - i)
        Next i
    Next j

    cover_ws.Columns.Delete
    cover_ws.Range("A2").Formula2 = "=IFERROR(SORTBY(TRANSPOSE(LINEST(data!A2:A" & _
        data_ws_lr & ",data!B2:" & Chr(64 + data_ws_lc) & data_ws_lr & ",TRUE,TRUE)),SEQUENCE(" & data_ws_lc & "),-1),"""")"
    cover_ws.Range("A1").Value2 = "Coefficients"
    cover_ws.Range("B1").Value2 = "Standard Error"

    For i = 1 To data_ws_lc
        total_beta = 0
        For j = 1 To b_iter
            total_beta = total_beta + beta(i, j)
        Next j
        avg_beta = total_beta / b_iter
        cover_ws.Range("G" & i + 1).Value2 = avg_beta

        sum_square = 0
        For j = 1 To b_iter
            sum_square = sum_square + (beta(i, j) - avg_beta) ^ 2
        Next j
        cover_ws.Range("H" & i + 1).Value2 = (sum_square / (b_iter - 1)) ^ 0.5
    Next i

    With cover_ws
        .Range("G1").Value2 = "Bootstrapped Coeff"
        .Range("H1").Value2 = "Bootstrapped SE"
        .Cells.NumberFormat = "0.000"
        .Columns.AutoFit
        .Activate
        .Range("A1").Select
    End With

    Windows(ThisWorkbook.Name).DisplayGridlines = False
Exit Sub

Handler:
    MsgBox Err.Description & ", " & Err.Number
End Sub
