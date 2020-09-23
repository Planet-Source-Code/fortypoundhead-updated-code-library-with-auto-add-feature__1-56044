Attribute VB_Name = "modQiuckSort"
Option Explicit

Public Sub QSort(Grid As MSFlexGrid, ByVal Column As Integer, ByVal min As Long, _
    ByVal max As Long, ByVal Ascending As Boolean, ByVal NumComp As Boolean)

'Call qsort function to sort
'msflexgrid is a reference to gridcontrol you want to sort
'column is by which column to sort by.
'min and max are the rows to sort.
'ascending is a bolean for ascending or descending sort
'NumComp is a boolean whether it's a numeric search or alphanumeric

    Dim tmp() ' when swap rows keep copy here
    ReDim tmp(Grid.Cols)
    Dim med_value, hi As Long, lo As Long, i As Integer

    If min >= max Then Exit Sub

    med_value = Grid.TextMatrix(min, Column)
    SaveRow Grid, min, tmp

    lo = min
    hi = max

    Do
        Do While Compare(Grid.TextMatrix(hi, Column), med_value, NumComp, Ascending) >= 0
            hi = hi - 1
            If hi <= lo Then Exit Do
        Loop
        If hi <= lo Then
            RestoreRow Grid, lo, tmp
            Exit Do
        End If
        For i = 0 To Grid.Cols - 1
            Grid.TextMatrix(lo, i) = Grid.TextMatrix(hi, i)
        Next i
        lo = lo + 1
        Do While Compare(Grid.TextMatrix(lo, Column), med_value, NumComp, Ascending) < 0
            lo = lo + 1
            If lo >= hi Then Exit Do
        Loop
        If lo >= hi Then
            lo = hi
            RestoreRow Grid, hi, tmp
            Exit Do
        End If
        For i = 0 To Grid.Cols - 1
            Grid.TextMatrix(hi, i) = Grid.TextMatrix(lo, i)
        Next i
    Loop

    QSort Grid, Column, min, lo - 1, Ascending, NumComp
    QSort Grid, Column, lo + 1, max, Ascending, NumComp
End Sub


Private Function Compare(ByVal X, ByVal Y, ByVal NumComp As Boolean, _
    ByVal Ascending As Boolean) As Integer

    Dim b As Integer

    If NumComp Then
        X = CDbl(X)
        Y = CDbl(Y)
    End If
    If X > Y Then b = 1
    If X < Y Then b = -1
    If X = Y Then b = 0
    If Not Ascending Then b = -b
    Compare = b
End Function

Private Sub RestoreRow(Grid As MSFlexGrid, ByVal RowNum As Long, tmpArr())

    Dim i As Long

    For i = 0 To Grid.Cols - 1
        Grid.TextMatrix(RowNum, i) = tmpArr(i)
    Next i
End Sub

Private Sub SaveRow(Grid As MSFlexGrid, ByVal RowNum As Long, tmpArr())

    Dim i As Long

    For i = 0 To Grid.Cols - 1
        tmpArr(i) = Grid.TextMatrix(RowNum, i)
    Next i
End Sub



