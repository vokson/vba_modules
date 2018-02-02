Attribute VB_Name = "M_SP63_13330_2012_Test"
Option Explicit

Sub Test()
Dim EMDD As New C_SP63_13330_2012_Tables
EMDD.Table_Ini

Debug.Print Formula_P411_SP63(20, 1, 30)

End Sub


Sub Table1()
Dim Formulas As New C_SP63_13330_2012_Formulas
Dim EMDD As New C_SP63_13330_2012_Tables
EMDD.Table_Ini
Dim Q As Double, M As Double, N As Double, maxM As Double, maxQ As Double
Dim col As Integer, row As Integer, rowCell As Integer, colCell As Integer, i As Integer
Dim Table_1(21, 8)
Erase Table_1()
Range("C54:K75").ClearContents
N = 0
Q = 0
i = 0
M = Formulas.maxM(Q, M, N)
    Table_1(i, 0) = Q
    Table_1(i, 1) = M
    Table_1(i, 2) = N
    Table_1(i, 4) = Round(Formulas.Formula_B1_SP63(Q, M, N), 3)
    Table_1(i, 5) = Round(Formulas.Formula_412_SP63(Q, M, N), 3)
    Table_1(i, 6) = Round(Formulas.Formula_P48_SP63(Q, M, N), 3)
    Table_1(i, 7) = Round(Formulas.Formula_P47_SP63(Q, M, N), 3)
    Table_1(i, 8) = Round(Formulas.Formula_P411_SP63(Q, M, N), 3)
    Table_1(i, 3) = Round(Application.Max(Table_1(i, 4), Table_1(i, 5), Table_1(i, 6), Table_1(i, 7), Table_1(i, 8)), 3)
i = i + 1
Q = Formulas.maxQ(Q, M, N)
    Table_1(i, 0) = Q
    Table_1(i, 1) = M
    Table_1(i, 2) = N
    Table_1(i, 4) = Round(Formulas.Formula_B1_SP63(Q, M, N), 3)
    Table_1(i, 5) = Round(Formulas.Formula_412_SP63(Q, M, N), 3)
    Table_1(i, 6) = Round(Formulas.Formula_P48_SP63(Q, M, N), 3)
    Table_1(i, 7) = Round(Formulas.Formula_P47_SP63(Q, M, N), 3)
    Table_1(i, 8) = Round(Formulas.Formula_P411_SP63(Q, M, N), 3)
    Table_1(i, 3) = Round(Application.Max(Table_1(i, 4), Table_1(i, 5), Table_1(i, 6), Table_1(i, 7), Table_1(i, 8)), 3)
If Q < Round(Q / 10, 0) * 10 Then
    Q = Round(Q / 10, 0) * 10
    Else
        Q = Round(Q / 10, 0) * 10 + 5
End If
M = Formulas.maxM(Q, 0, 0)
i = i + 1
    Table_1(i, 0) = Q
    Table_1(i, 1) = M
    Table_1(i, 2) = N
    Table_1(i, 4) = Round(Formulas.Formula_B1_SP63(Q, M, N), 3)
    Table_1(i, 5) = Round(Formulas.Formula_412_SP63(Q, M, N), 3)
    Table_1(i, 6) = Round(Formulas.Formula_P48_SP63(Q, M, N), 3)
    Table_1(i, 7) = Round(Formulas.Formula_P47_SP63(Q, M, N), 3)
    Table_1(i, 8) = Round(Formulas.Formula_P411_SP63(Q, M, N), 3)
    Table_1(i, 3) = Round(Application.Max(Table_1(i, 4), Table_1(i, 5), Table_1(i, 6), Table_1(i, 7), Table_1(i, 8)), 3)
maxQ = calcQ()
    Q = Q + 5
    i = 3
    Do While Q < maxQ And i < 22
        M = Formulas.maxM(Q, 0, 0)
            Table_1(i, 0) = Q
            Table_1(i, 1) = M
            Table_1(i, 2) = N
            Table_1(i, 4) = Round(Formulas.Formula_B1_SP63(Q, M, N), 3)
            Table_1(i, 5) = Round(Formulas.Formula_412_SP63(Q, M, N), 3)
            Table_1(i, 6) = Round(Formulas.Formula_P48_SP63(Q, M, N), 3)
            Table_1(i, 7) = Round(Formulas.Formula_P47_SP63(Q, M, N), 3)
            Table_1(i, 8) = Round(Formulas.Formula_P411_SP63(Q, M, N), 3)
            Table_1(i, 3) = Round(Application.Max(Table_1(i, 4), Table_1(i, 5), Table_1(i, 6), Table_1(i, 7), Table_1(i, 8)), 3)
    i = i + 1
    Q = Q + 5
    Loop
Q = maxQ
M = Q * EMDD.ea
    Table_1(i, 0) = Q
    Table_1(i, 1) = M
    Table_1(i, 2) = N
    Table_1(i, 4) = Round(Formulas.Formula_B1_SP63(Q, M, N), 3)
    Table_1(i, 5) = Round(Formulas.Formula_412_SP63(Q, M, N), 3)
    Table_1(i, 6) = Round(Formulas.Formula_P48_SP63(Q, M, N), 3)
    Table_1(i, 7) = Round(Formulas.Formula_P47_SP63(Q, M, N), 3)
    Table_1(i, 8) = Round(Formulas.Formula_P411_SP63(Q, M, N), 3)
    Table_1(i, 3) = Round(Application.Max(Table_1(i, 4), Table_1(i, 5), Table_1(i, 6), Table_1(i, 7), Table_1(i, 8)), 3)
        i = i + 1
        Table_1(i, 0) = Q
        Table_1(i, 1) = 0
        Table_1(i, 2) = N
        Table_1(i, 4) = Round(Formulas.Formula_B1_SP63(Q, M, N), 3)
        Table_1(i, 5) = Round(Formulas.Formula_412_SP63(Q, M, N), 3)
        Table_1(i, 6) = Round(Formulas.Formula_P48_SP63(Q, M, N), 3)
        Table_1(i, 7) = Round(Formulas.Formula_P47_SP63(Q, M, N), 3)
        Table_1(i, 8) = Round(Formulas.Formula_P411_SP63(Q, M, N), 3)
        Table_1(i, 3) = Round(Application.Max(Table_1(i, 4), Table_1(i, 5), Table_1(i, 6), Table_1(i, 7), Table_1(i, 8)), 3)

'// вывод в таблицу
'   rowCell = ActiveCell.row     colCell = ActiveCell.Column
    rowCell = 54
    colCell = 3
For col = 0 To 8
    For row = 0 To 21
        Cells(row + rowCell, col + colCell).Value = Table_1(row, col)
    Next row
Next col
 
End Sub

Sub Table2()
Dim Formulas As New C_SP63_13330_2012_Formulas
Dim EMDD As New C_SP63_13330_2012_Tables
EMDD.Table_Ini
Dim Q As Double, M As Double, N As Double, maxM As Double, maxQ As Double
Dim col As Integer, row As Integer, rowCell As Integer, colCell As Integer, i As Integer
Dim Table_2(21, 8)
Erase Table_2()
Range("C76:K97").ClearContents
M = 0
Q = 0
i = 0
N = Formulas.maxN_End(Q, M, N)
    Table_2(i, 0) = Q
    Table_2(i, 1) = 0
    Table_2(i, 2) = N
    Table_2(i, 4) = Round(Formulas.Formula_B1_SP63(Q, 0, N), 3)
    Table_2(i, 5) = Round(Formulas.Formula_412_SP63(Q, 0, N), 3)
    Table_2(i, 6) = Round(Formulas.Formula_P48_SP63(Q, 0, N), 3)
    Table_2(i, 7) = Round(Formulas.Formula_P47_SP63(Q, 0, N), 3)
    Table_2(i, 8) = Round(Formulas.Formula_P411_SP63(Q, 0, N), 3)
    Table_2(i, 3) = Round(Application.Max(Table_2(i, 4), Table_2(i, 5), Table_2(i, 6), Table_2(i, 7), Table_2(i, 8)), 3)
i = i + 1
    Table_2(i, 0) = Q
    Table_2(i, 1) = M
    Table_2(i, 2) = N
    Table_2(i, 4) = Round(Formulas.Formula_B1_SP63(Q, M, N), 3)
    Table_2(i, 5) = Round(Formulas.Formula_412_SP63(Q, M, N), 3)
    Table_2(i, 6) = Round(Formulas.Formula_P48_SP63(Q, M, N), 3)
    Table_2(i, 7) = Round(Formulas.Formula_P47_SP63(Q, M, N), 3)
    Table_2(i, 8) = Round(Formulas.Formula_P411_SP63(Q, M, N), 3)
    Table_2(i, 3) = Round(Application.Max(Table_2(i, 4), Table_2(i, 5), Table_2(i, 6), Table_2(i, 7), Table_2(i, 8)), 3)
i = i + 1
If N > Round(N / 10, 0) * 10 Then
    N = Round(N / 10, 0) * 10
    Else
        N = Round(N / 10, 0) * 10 - 5
End If
    Do While N >= 0 And i < 22
        M = Formulas.maxM(0, 0, N)
            Table_2(i, 0) = Q
            Table_2(i, 1) = M
            Table_2(i, 2) = N
            Table_2(i, 4) = Round(Formulas.Formula_B1_SP63(Q, M, N), 3)
            Table_2(i, 5) = Round(Formulas.Formula_412_SP63(Q, M, N), 3)
            Table_2(i, 6) = Round(Formulas.Formula_P48_SP63(Q, M, N), 3)
            Table_2(i, 7) = Round(Formulas.Formula_P47_SP63(Q, M, N), 3)
            Table_2(i, 8) = Round(Formulas.Formula_P411_SP63(Q, M, N), 3)
            Table_2(i, 3) = Round(Application.Max(Table_2(i, 4), Table_2(i, 5), Table_2(i, 6), Table_2(i, 7), Table_2(i, 8)), 3)
'Debug.Print i & " - " & N
    i = i + 1
    N = N - 5
    Loop
   
'// вывод в таблицу
'   rowCell = ActiveCell.row     colCell = ActiveCell.Column
    rowCell = 76
    colCell = 3
For col = 0 To 8
    For row = 0 To 21
        Cells(row + rowCell, col + colCell).Value = Table_2(row, col)
    Next row
Next col

End Sub

Sub Table3()
Dim Formulas As New C_SP63_13330_2012_Formulas
Dim EMDD As New C_SP63_13330_2012_Tables
EMDD.Table_Ini
Dim Q As Double, M As Double, N As Double, maxM As Double, maxQ As Double
Dim col As Integer, row As Integer, rowCell As Integer, colCell As Integer, i As Integer
Dim Table_3(21, 8)
Erase Table_3()
Range("C98:K119").ClearContents

M = 0
Q = 0
i = 0
N = Formulas.maxN_End(Q, M, N)
    Table_3(i, 0) = Q
    Table_3(i, 1) = M
    Table_3(i, 2) = N
    Table_3(i, 4) = Round(Formulas.Formula_B1_SP63(Q, 0, N), 3)
    Table_3(i, 5) = Round(Formulas.Formula_412_SP63(Q, 0, N), 3)
    Table_3(i, 6) = Round(Formulas.Formula_P48_SP63(Q, 0, N), 3)
    Table_3(i, 7) = Round(Formulas.Formula_P47_SP63(Q, 0, N), 3)
    Table_3(i, 8) = Round(Formulas.Formula_P411_SP63(Q, 0, N), 3)
    Table_3(i, 3) = Round(Application.Max(Table_3(i, 4), Table_3(i, 5), Table_3(i, 6), Table_3(i, 7), Table_3(i, 8)), 3)
maxQ = calcQ()
i = i + 1
Q = Q + 2.5
N = 0
    Do While Q < maxQ And i < 22
        N = Formulas.maxNM(Q, M, N)
        Table_3(i, 0) = Q
        Table_3(i, 1) = M
        Table_3(i, 2) = N
        Table_3(i, 4) = Round(Formulas.Formula_B1_SP63(Q, M, N), 3)
        Table_3(i, 5) = Round(Formulas.Formula_412_SP63(Q, M, N), 3)
        Table_3(i, 6) = Round(Formulas.Formula_P48_SP63(Q, M, N), 3)
        Table_3(i, 7) = Round(Formulas.Formula_P47_SP63(Q, M, N), 3)
        Table_3(i, 8) = Round(Formulas.Formula_P411_SP63(Q, M, N), 3)
        Table_3(i, 3) = Round(Application.Max(Table_3(i, 4), Table_3(i, 5), Table_3(i, 6), Table_3(i, 7), Table_3(i, 8)), 3)
    i = i + 1
    Q = Q + 2.5
    N = 0
    Loop
    
Q = maxQ
M = Q * EMDD.ea + N * EMDD.ea
            Table_3(i, 0) = Q
            Table_3(i, 1) = M
            Table_3(i, 2) = N
            Table_3(i, 4) = Round(Formulas.Formula_B1_SP63(Q, M, N), 3)
            Table_3(i, 5) = Round(Formulas.Formula_412_SP63(Q, M, N), 3)
            Table_3(i, 6) = Round(Formulas.Formula_P48_SP63(Q, M, N), 3)
            Table_3(i, 7) = Round(Formulas.Formula_P47_SP63(Q, M, N), 3)
            Table_3(i, 8) = Round(Formulas.Formula_P411_SP63(Q, M, N), 3)
            Table_3(i, 3) = Round(Application.Max(Table_3(i, 4), Table_3(i, 5), Table_3(i, 6), Table_3(i, 7), Table_3(i, 8)), 3)
   
'// вывод в таблицу
'   rowCell = ActiveCell.row     colCell = ActiveCell.Column
    rowCell = 98
    colCell = 3
For col = 0 To 8
    For row = 0 To 21
        Cells(row + rowCell, col + colCell).Value = Table_3(row, col)
    Next row
Next col

End Sub

Sub Table4()
Dim Formulas As New C_SP63_13330_2012_Formulas
Dim EMDD As New C_SP63_13330_2012_Tables
EMDD.Table_Ini
Dim Q As Double, M As Double, N As Double, maxM As Double, maxQ As Double
Dim col As Integer, row As Integer, rowCell As Integer, colCell As Integer, i As Integer, Tbl As Integer
Dim Table_4(1 To 5, 21, 8)
Tbl = 1
Erase Table_4()
Range("C120:K229").ClearContents


For Tbl = 1 To 5
'M = 1 + (Tbl - 1) * 0.5
M = Cells(120 + (Tbl - 1), 14)
Q = 0
i = 0
N = Formulas.maxN(Q, M, N)
    Table_4(Tbl, i, 0) = Q
    Table_4(Tbl, i, 1) = M
    Table_4(Tbl, i, 2) = N
    Table_4(Tbl, i, 4) = Round(Formulas.Formula_B1_SP63(Q, 0, N), 3)
    Table_4(Tbl, i, 5) = Round(Formulas.Formula_412_SP63(Q, 0, N), 3)
    Table_4(Tbl, i, 6) = Round(Formulas.Formula_P48_SP63(Q, 0, N), 3)
    Table_4(Tbl, i, 7) = Round(Formulas.Formula_P47_SP63(Q, 0, N), 3)
    Table_4(Tbl, i, 8) = Round(Formulas.Formula_P411_SP63(Q, 0, N), 3)
    Table_4(Tbl, i, 3) = Round(Application.Max(Table_4(Tbl, i, 4), Table_4(Tbl, i, 5), Table_4(Tbl, i, 6), Table_4(Tbl, i, 7), Table_4(Tbl, i, 8)), 3)
maxQ = Formulas.maxQ(0, M, 0)
Q = Formulas.maxQ(Q, M, N)
i = i + 1
    Table_4(Tbl, i, 0) = Q
    Table_4(Tbl, i, 1) = M
    Table_4(Tbl, i, 2) = N
    Table_4(Tbl, i, 4) = Round(Formulas.Formula_B1_SP63(Q, 0, N), 3)
    Table_4(Tbl, i, 5) = Round(Formulas.Formula_412_SP63(Q, 0, N), 3)
    Table_4(Tbl, i, 6) = Round(Formulas.Formula_P48_SP63(Q, 0, N), 3)
    Table_4(Tbl, i, 7) = Round(Formulas.Formula_P47_SP63(Q, 0, N), 3)
    Table_4(Tbl, i, 8) = Round(Formulas.Formula_P411_SP63(Q, 0, N), 3)
    Table_4(Tbl, i, 3) = Round(Application.Max(Table_4(Tbl, i, 4), Table_4(Tbl, i, 5), Table_4(Tbl, i, 6), Table_4(Tbl, i, 7), Table_4(Tbl, i, 8)), 3)
maxQ = Formulas.maxQ(0, M, 0)
If Q < Round(Q / 10, 0) * 10 Then
    Q = Round(Q / 10, 0) * 10
    Else
        Q = Round(Q / 10, 0) * 10 + 5
End If
i = i + 1
'Q = Q + 2.5
N = 0
    Do While Q < maxQ And i < 22
        N = Formulas.maxN(Q, M, N)
        Table_4(Tbl, i, 0) = Q
        Table_4(Tbl, i, 1) = M
        Table_4(Tbl, i, 2) = N
        Table_4(Tbl, i, 4) = Round(Formulas.Formula_B1_SP63(Q, M, N), 3)
        Table_4(Tbl, i, 5) = Round(Formulas.Formula_412_SP63(Q, M, N), 3)
        Table_4(Tbl, i, 6) = Round(Formulas.Formula_P48_SP63(Q, M, N), 3)
        Table_4(Tbl, i, 7) = Round(Formulas.Formula_P47_SP63(Q, M, N), 3)
        Table_4(Tbl, i, 8) = Round(Formulas.Formula_P411_SP63(Q, M, N), 3)
        Table_4(Tbl, i, 3) = Round(Application.Max(Table_4(Tbl, i, 4), Table_4(Tbl, i, 5), Table_4(Tbl, i, 6), Table_4(Tbl, i, 7), Table_4(Tbl, i, 8)), 3)
    i = i + 1
    Q = Q + 2.5
    N = 0
    Loop
    
Q = maxQ
'M = Q * EMDD.ea + N * EMDD.ea
            Table_4(Tbl, i, 0) = Q
            Table_4(Tbl, i, 1) = M
            Table_4(Tbl, i, 2) = N
            Table_4(Tbl, i, 4) = Round(Formulas.Formula_B1_SP63(Q, M, N), 3)
            Table_4(Tbl, i, 5) = Round(Formulas.Formula_412_SP63(Q, M, N), 3)
            Table_4(Tbl, i, 6) = Round(Formulas.Formula_P48_SP63(Q, M, N), 3)
            Table_4(Tbl, i, 7) = Round(Formulas.Formula_P47_SP63(Q, M, N), 3)
            Table_4(Tbl, i, 8) = Round(Formulas.Formula_P411_SP63(Q, M, N), 3)
            Table_4(Tbl, i, 3) = Round(Application.Max(Table_4(Tbl, i, 4), Table_4(Tbl, i, 5), Table_4(Tbl, i, 6), Table_4(Tbl, i, 7), Table_4(Tbl, i, 8)), 3)
   
'// вывод в таблицу
'   rowCell = ActiveCell.row     colCell = ActiveCell.Column

    rowCell = 120 + 22 * (Tbl - 1)
    colCell = 3
    Debug.Print rowCell
For col = 0 To 8
    For row = 0 To 21
        Cells(row + rowCell, col + colCell).Value = Table_4(Tbl, row, col)
    Next row
Next col
'    Tbl = Tbl + 1
Next Tbl
End Sub

Function calcQ()
Dim Formulas As New C_SP63_13330_2012_Formulas
Dim Q As Double, e As Double, N As Double
    calcQ = Formulas.maxQ_End(0, 0, 0)
End Function




