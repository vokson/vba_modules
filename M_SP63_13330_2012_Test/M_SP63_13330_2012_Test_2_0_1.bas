Attribute VB_Name = "M_SP63_13330_2012_Test"
Public Table(8, 30, 8) As Double
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
Dim col As Integer, row As Integer, rowCell As Integer, colCell As Integer, i As Integer, Tbl As Integer

Erase Table()
Tbl = 1
Range("C54:K75").ClearContents

i = 0
    Q = 0
    M = Formulas.maxM(Q, M, N)
    N = 0
    BuildTable Tbl, i, Q, M, N
        i = i + 1
        Q = Formulas.maxQ(Q, M, N)
        BuildTable Tbl, i, Q, M, N
            If Q < Round(Q / 10, 0) * 10 Then
                Q = Round(Q / 10, 0) * 10
                Else
                    Q = Round(Q / 10, 0) * 10 + 5
            End If
            M = Formulas.maxM(Q, 0, 0)
            i = i + 1
            BuildTable Tbl, i, Q, M, N
                maxQ = calcQ()
                Q = Q + 5
                i = i + 1
                Do While Q < maxQ And i < 22
                    M = Formulas.maxM(Q, 0, 0)
                    BuildTable Tbl, i, Q, M, N
                        i = i + 1
                        Q = Q + 5
                Loop
                Q = maxQ
                M = Q * EMDD.ea
                BuildTable Tbl, i, Q, M, N
                    i = i + 1
                    BuildTable Tbl, i, Q, 0, N

'// вывод в таблицу
    rowCell = 54
    colCell = 3
For col = 0 To 8
    For row = 0 To 21
        Cells(row + rowCell, col + colCell).Value = Table(Tbl, row, col)
    Next row
Next col
    For row = 0 To 21
        If Table(Tbl, row, 0) = 0 And Table(Tbl, row, 1) = 0 And Table(Tbl, row, 2) = 0 Then
            For col = 3 To 11
                Cells(row + rowCell, col).Value = ""
            Next col
        End If
    Next row
 
End Sub

Sub Table2()
Dim Formulas As New C_SP63_13330_2012_Formulas
Dim EMDD As New C_SP63_13330_2012_Tables
EMDD.Table_Ini
Dim Q As Double, M As Double, N As Double, maxM As Double, maxQ As Double
Dim col As Integer, row As Integer, rowCell As Integer, colCell As Integer, i As Integer, Tbl As Integer

Erase Table()
Tbl = 2
Range("C76:K97").ClearContents

    i = 0
    Q = 0
    M = 0
    N = Formulas.maxN_End(Q, M, N)
    M = 0
    BuildTable Tbl, i, Q, M, N
        i = i + 1
        N = Formulas.maxN_End(Q, M, N)
        BuildTable Tbl, i, Q, M, N
            i = i + 1
            If N > Round(N / 10, 0) * 10 Then
                N = Round(N / 10, 0) * 10
                Else
                    N = Round(N / 10, 0) * 10 - 5
            End If
            Do While N >= 0 And i < 22
                M = Formulas.maxM(0, 0, N)
                BuildTable Tbl, i, Q, M, N
                    i = i + 1
                    N = N - 5
            Loop
   
'// вывод в таблицу
    rowCell = 76
    colCell = 3
For col = 0 To 8
    For row = 0 To 21
        Cells(row + rowCell, col + colCell).Value = Table(Tbl, row, col)
    Next row
Next col
    For row = 0 To 21
        If Table(Tbl, row, 0) = 0 And Table(Tbl, row, 1) = 0 And Table(Tbl, row, 2) = 0 Then
            For col = 3 To 11
                Cells(row + rowCell, col).Value = ""
            Next col
        End If
    Next row

End Sub

Sub Table3()
Dim Formulas As New C_SP63_13330_2012_Formulas
Dim EMDD As New C_SP63_13330_2012_Tables
EMDD.Table_Ini
Dim Q As Double, M As Double, N As Double, maxM As Double, maxQ As Double
Dim col As Integer, row As Integer, rowCell As Integer, colCell As Integer, i As Integer, Tbl As Integer, aCount As Integer, step As Double, Qnext As Double

Erase Table()
Tbl = 3
Range("C98:K127").ClearContents

    aCount = 0
    step = 5
    i = 0
    Q = 0
    M = 0
    N = Formulas.maxN_End(Q, M, N)
    M = 0
    BuildTable Tbl, i, Q, M, N
        maxQ = calcQ()
        i = i + 1
        Q = Q + step
        N = 0
        Do While Q < maxQ And i < 29
            N = Formulas.maxNM(Q, M, N)
                If Round(Formulas.Formula_B1_SP63(Q, M, N), 3) > Round(Formulas.Formula_P47_SP63(Q, M, N), 3) And aCount = 0 Then
                    Qnext = Q
                    Q = Q - step
                        Do While Q < Qnext
                            N = Formulas.maxNM(Q, M, 0)
                            Q = Q + 0.1
                            If Round(Formulas.Formula_B1_SP63(Q, M, N), 3) > Round(Formulas.Formula_P47_SP63(Q, M, N), 3) Then
                                Exit Do
                            End If
                        Loop
                    N = Formulas.maxNM(Q, M, 0)
                    BuildTable Tbl, i, Q, M, N
                    aCount = aCount + 1
                    Q = Qnext
                    i = i + 1
                    step = 2.5
                End If
            N = Formulas.maxNM(Q, M, 0)
            BuildTable Tbl, i, Q, M, N
                i = i + 1
                Q = Q + step
                N = 0
        Loop
            Q = maxQ
            M = Q * EMDD.ea + N * EMDD.ea
            BuildTable Tbl, i, Q, M, N
   
'// вывод в таблицу  
    rowCell = 98
    colCell = 3
For col = 0 To 8
    For row = 0 To 29
        Cells(row + rowCell, col + colCell).Value = Table(Tbl, row, col)
    Next row
Next col
    For row = 0 To 29
        If Table(Tbl, row, 0) = 0 And Table(Tbl, row, 1) = 0 And Table(Tbl, row, 2) = 0 Then
            For col = 3 To 11
                Cells(row + rowCell, col).Value = ""
            Next col
        End If
    Next row

End Sub

Sub Table4()
Dim Formulas As New C_SP63_13330_2012_Formulas
Dim EMDD As New C_SP63_13330_2012_Tables
EMDD.Table_Ini
Dim Q As Double, M As Double, N As Double, maxM As Double, maxQ As Double
Dim col As Integer, row As Integer, rowCell As Integer, colCell As Integer, i As Integer, Tbl As Integer

Erase Table()
Range("C128:K237").ClearContents

For Tbl = 4 To 8
    i = 0
    Q = 0
    M = Cells(128 + (Tbl - 4), 14)
    N = Formulas.maxN(Q, M, N)
    BuildTable Tbl, i, Q, M, N
        maxQ = Formulas.maxQ(0, M, 0)
        Q = Formulas.maxQ(Q, M, N)
        i = i + 1
        BuildTable Tbl, i, Q, M, N
            maxQ = Formulas.maxQ(0, M, 0)
            If Q < Round(Q / 10, 0) * 10 Then
                Q = Round(Q / 10, 0) * 10
                Else
                    Q = Round(Q / 10, 0) * 10 + 5
            End If
            i = i + 1
            N = 0
            Do While Q < maxQ And i < 22
                N = Formulas.maxN(Q, M, N)
                BuildTable Tbl, i, Q, M, N
                    i = i + 1
                    Q = Q + 2.5
                    N = 0
            Loop
                Q = maxQ
                BuildTable Tbl, i, Q, M, N

'// вывод в таблицу
    rowCell = 128 + 22 * (Tbl - 4)
    colCell = 3
    For col = 0 To 8
        For row = 0 To 21
            Cells(row + rowCell, col + colCell).Value = Table(Tbl, row, col)
        Next row
    Next col
    For row = 0 To 21
        If Table(Tbl, row, 0) = 0 And Table(Tbl, row, 1) = 0 And Table(Tbl, row, 2) = 0 Then
            For col = 3 To 11
                Cells(row + rowCell, col).Value = ""
            Next col
        End If
    Next row
Next Tbl
End Sub

Sub BuildTable(Tbl As Integer, i As Integer, Q As Double, M As Double, N As Double)
Dim Formulas As New C_SP63_13330_2012_Formulas
    Table(Tbl, i, 0) = Q
    Table(Tbl, i, 1) = M
    Table(Tbl, i, 2) = N
    Table(Tbl, i, 4) = Round(Formulas.Formula_B1_SP63(Q, M, N), 3)
    Table(Tbl, i, 5) = Round(Formulas.Formula_412_SP63(Q, M, N), 3)
    Table(Tbl, i, 6) = Round(Formulas.Formula_P48_SP63(Q, M, N), 3)
    Table(Tbl, i, 7) = Round(Formulas.Formula_P47_SP63(Q, M, N), 3)
    Table(Tbl, i, 8) = Round(Formulas.Formula_P411_SP63(Q, M, N), 3)
    Table(Tbl, i, 3) = Round(Application.Max(Table(Tbl, i, 4), Table(Tbl, i, 5), Table(Tbl, i, 6), Table(Tbl, i, 7), Table(Tbl, i, 8)), 3)
End Sub

Function calcQ()
Dim Formulas As New C_SP63_13330_2012_Formulas
Dim Q As Double, e As Double, N As Double
    calcQ = Formulas.maxQ_End(0, 0, 0)
End Function
