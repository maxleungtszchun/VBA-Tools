Attribute VB_Name = "sortMatrix"
Option Explicit

Public Function sortMatrix(matrix As Variant, col_index_1 As Long, order_1 As Integer, col_index_2 As Long, order_2 As Integer) As Variant
    Dim column_1 As Variant
    Dim column_2 As Variant
    
    If Not order_1 = -1 And Not order_1 = 1 Then
        Err.Raise Number:=vbObjectError + 2222, Description:="order_1 must be -1 or 1"
    End If
    
    If Not order_2 = -1 And Not order_2 = 1 Then
        Err.Raise Number:=vbObjectError + 3333, Description:="order_2 must be -1 or 1"
    End If
    
    column_1 = columnInMatrix(matrix, col_index_1)
    column_2 = columnInMatrix(matrix, col_index_2)
    
    sortMatrix = Application.WorksheetFunction.SortBy(matrix, column_1, order_1, column_2, order_2)
End Function

Private Function columnInMatrix(matrix As Variant, col_index As Long) As Variant
    Dim arr As Variant
    Dim i As Long
    
    If col_index <= 0 Then
        Err.Raise Number:=vbObjectError + 1111, Description:="col_index must be > 0"
    End If
    
    ReDim arr(1 To UBound(matrix, 1), 1 To 1)
    For i = LBound(matrix, 1) To UBound(matrix, 1)
        arr(i, 1) = matrix(i, col_index)
    Next i
    
    columnInMatrix = arr
End Function
