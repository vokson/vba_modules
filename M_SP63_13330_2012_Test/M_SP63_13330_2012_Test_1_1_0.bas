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
Dim Table_1(9, 9)
Erase Table_1()
N = 0
Q = 0
i = 0
M = Formulas.maxM(Q, M, N)
    Table_1(i, 0) = Q
    Table_1(i, 1) = M
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
    Table_1(i, 4) = Round(Formulas.Formula_B1_SP63(Q, M, N), 3)
    Table_1(i, 5) = Round(Formulas.Formula_412_SP63(Q, M, N), 3)
    Table_1(i, 6) = Round(Formulas.Formula_P48_SP63(Q, M, N), 3)
    Table_1(i, 7) = Round(Formulas.Formula_P47_SP63(Q, M, N), 3)
    Table_1(i, 8) = Round(Formulas.Formula_P411_SP63(Q, M, N), 3)
    Table_1(i, 3) = Round(Application.Max(Table_1(i, 4), Table_1(i, 5), Table_1(i, 6), Table_1(i, 7), Table_1(i, 8)), 3)
Q = Round(Q / 10, 0) * 10
M = Formulas.maxM(Q, 0, 0)
i = i + 1
    Table_1(i, 0) = Q
    Table_1(i, 1) = M
    Table_1(i, 4) = Round(Formulas.Formula_B1_SP63(Q, M, N), 3)
    Table_1(i, 5) = Round(Formulas.Formula_412_SP63(Q, M, N), 3)
    Table_1(i, 6) = Round(Formulas.Formula_P48_SP63(Q, M, N), 3)
    Table_1(i, 7) = Round(Formulas.Formula_P47_SP63(Q, M, N), 3)
    Table_1(i, 8) = Round(Formulas.Formula_P411_SP63(Q, M, N), 3)
    Table_1(i, 3) = Round(Application.Max(Table_1(i, 4), Table_1(i, 5), Table_1(i, 6), Table_1(i, 7), Table_1(i, 8)), 3)
maxQ = calcQ()
    Q = Q + 5
    i = 3
    Do While Q < maxQ
        M = Formulas.maxM(Q, 0, 0)
            Table_1(i, 0) = Q
            Table_1(i, 1) = M
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
    Table_1(i, 4) = Round(Formulas.Formula_B1_SP63(Q, M, N), 3)
    Table_1(i, 5) = Round(Formulas.Formula_412_SP63(Q, M, N), 3)
    Table_1(i, 6) = Round(Formulas.Formula_P48_SP63(Q, M, N), 3)
    Table_1(i, 7) = Round(Formulas.Formula_P47_SP63(Q, M, N), 3)
    Table_1(i, 8) = Round(Formulas.Formula_P411_SP63(Q, M, N), 3)
    Table_1(i, 3) = Round(Application.Max(Table_1(i, 4), Table_1(i, 5), Table_1(i, 6), Table_1(i, 7), Table_1(i, 8)), 3)
        i = i + 1
        Table_1(i, 0) = Q
        Table_1(i, 1) = 0
        Table_1(i, 4) = Round(Formulas.Formula_B1_SP63(Q, M, N), 3)
        Table_1(i, 5) = Round(Formulas.Formula_412_SP63(Q, M, N), 3)
        Table_1(i, 6) = Round(Formulas.Formula_P48_SP63(Q, M, N), 3)
        Table_1(i, 7) = Round(Formulas.Formula_P47_SP63(Q, M, N), 3)
        Table_1(i, 8) = Round(Formulas.Formula_P411_SP63(Q, M, N), 3)
        Table_1(i, 3) = Round(Application.Max(Table_1(i, 4), Table_1(i, 5), Table_1(i, 6), Table_1(i, 7), Table_1(i, 8)), 3)

'// вывод в таблицу
    rowCell = ActiveCell.row
    colCell = ActiveCell.Column
For col = 0 To 8
    For row = 0 To 8
        Cells(row + rowCell, col + colCell).Value = Table_1(row, col)
    Next row
Next col
 
End Sub

Function calcQ()
Dim Formulas As New C_SP63_13330_2012_Formulas
Dim Q As Double, e As Double, N As Double
    calcQ = Formulas.maxQ_End(0, 0, 0)
End Function




