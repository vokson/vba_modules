Attribute VB_Name = "M_SP63_13330_2012_Test"
Option Explicit

Sub A_Test()
Dim EMDD As New C_SP63_13330_2012_Tables
EMDD.Table_Ini

Debug.Print Formula_P411_SP63(20, 1, 30)

End Sub


'// �������� ��������� ������� ��������� ������� (�� ���������� �, �.�.1 [1]):
Public Function Formula_B1_SP63(Q As Double, M As Double, N As Double) As Double
Application.Volatile

Dim EMDD As New C_SP63_13330_2012_Tables
Dim Formulas As New C_SP63_13330_2012_Formulas
EMDD.Table_Ini
Dim Qanj0 As Double, Nanj0 As Double, gssh As Double, Nanj As Double, Qanj As Double
       
gssh = 1.65
    Qanj0 = gssh * EMDD.Aanj_Anchor(EMDD.dAnchor) * Sqr(EMDD.Rs(EMDD.GradeAnchor) * EMDD.Rb_gb(EMDD.GradeConcrete)) * 1000 '// ���������� ����, �������������� �������� (������� (�.5) [1])
    Nanj0 = EMDD.Rs(EMDD.GradeAnchor) * EMDD.Aanj_Anchor(EMDD.dAnchor) * 1000 '// ���������� ������������� ����, �������������� ����� ����� ������� (������� (�.6) [1])
    Nanj = Formulas.Nanj(M, N)       '// ���������� ������������� ������ � ����� ���� �������  (������� (�.2) [1])
    Qanj = Formulas.Qanj(Q, M, N)      '// ���������� ������, ������������ �� ���� ��� ������� (������� (�.3) [1])
          
    Formula_B1_SP63 = Qanj / Qanj0 + Nanj / Nanj0
End Function


'// �������� ������ ������ ��� ���������� �������� (�� �.4.12 [2]):
Public Function Formula_412_SP63(Q As Double, M As Double, N As Double) As Double
Application.Volatile

Dim EMDD As New C_SP63_13330_2012_Tables
Dim Formulas As New C_SP63_13330_2012_Formulas
EMDD.Table_Ini
Dim Nanj As Double, Qanj As Double, Nan1 As Double, Qan1 As Double, w As Double, Dl As Double, jc As Double, jb As Double
Dim lan As Double, Nloc As Double, Rbloc As Double, Aloc As Double, Ab As Double, Abmax As Double
Dim minL As Double

    Nanj = Formulas.Nanj(M, N)       '// ���������� ������������� ������ � ����� ���� �������  (������� (�.2) [1])
    Qanj = Formulas.Qanj(Q, M, N)      '// ���������� ������, ������������ �� ���� ��� ������� (������� (�.3) [1])
    Nan1 = Nanj / EMDD.n_Anchor '// ������������� � ���������� ������ � ����� �������� �������
    Qan1 = Qanj / EMDD.n_Anchor '// ������������� � ���������� ������ � ����� �������� �������
    w = 0.7
    Dl = 11
    If Nan1 <= 0 Then
        jc = 0.7
        Else
            jc = 0.3 / (1 + Qan1 / Nan1) + 0.7
    End If
    
    lan = Formulas.lan(w, Dl, jc, Q, M, N) '//��������� ����� ��������� lan ����������� ������� �������������� ������� ��� �������� (�.5.7 [2]):
    If lan < EMDD.dAnchor * 20 And lan < 250 Then
        If EMDD.dAnchor > 250 Then
            lan = EMDD.dAnchor * 20
            Else
                lan = 250
        End If
    End If
        
    If (EMDD.La_Anchor < 15 * EMDD.dAnchor) Then '// ��������� ���� �� �������� ������ Nloc
        Nloc = Nan1 + Qan1 * (15 * EMDD.dAnchor - EMDD.La_Anchor) / lan '// ��� La (����� ������) < 15d (������� ������ (��))
    Else
        Nloc = Nan1 '// ��� La (����� ������) >= 15d (������� ������ (��))(�. 42 [2])
    End If

    EMDD.Aanj_Anchor (EMDD.dAnchor)
    Aloc = (EMDD.Lpl_Anchor / 1000) ^ 2 - EMDD.Aan1 / 10000 '// ������� ������

        If EMDD.lzZ_Anchor < EMDD.lzY_Anchor Then
            minL = EMDD.lzZ_Anchor
            Else
                minL = EMDD.lzY_Anchor
        End If
    
    If 3 * EMDD.Lpl_Anchor > EMDD.Lpl_Anchor + (minL - EMDD.Lpl_Anchor) Then
        Ab = EMDD.Lpl_Anchor + (minL - EMDD.Lpl_Anchor)
        Else
            Ab = 3 * EMDD.Lpl_Anchor
    End If
    Abmax = Ab ^ 2 / 1000000 - EMDD.Aan1 / 10000

    If 0.8 * Sqr(Abmax / Aloc) > 1 Then '// (�.8.82 [1]) �� �� ����� 1,0 � �� ����� 2,5
        If 0.8 * Sqr(Abmax / Aloc) < 2.5 Then
            jb = 0.8 * Sqr(Abmax / Aloc)
            Else
                jb = 2.5
        End If
        Else
            jb = 1
    End If
    Rbloc = EMDD.Rb_gb(EMDD.GradeConcrete) * jb '// (�. 8.81 [1])
Formula_412_SP63 = Nloc / (Rbloc * Aloc * 1000)  '// ����������� �������������; ������� ������ (�� �.41 [2] � ������ ���������� �.8.1.44 [1])

End Function

'// �������� ������ �� ����������� � ���������� �������, ��� N'an > 0  (�� �.4.8 [2]):
'            ��������: ����������� ������ ������������ ����� (e=0) (� ����� ���������)
Public Function Formula_P48_SP63(Q As Double, M As Double, N As Double) As Double
Application.Volatile

Dim EMDD As New C_SP63_13330_2012_Tables
Dim Formulas As New C_SP63_13330_2012_Formulas
EMDD.Table_Ini
Dim Nanj As Double, nan3 As Double
Dim a As Double, b As Double, A1 As Double, Aan As Double
Dim j2 As Double, j3 As Double, e As Double, X As Double, Y As Double, DLpl As Double, nan2 As Double

    Nanj = Formulas.Nanj(M, N)       '// ���������� ������������� ������ � ����� ���� �������  (������� (�.2) [1])
    If (Nanj < 0) Then
        nan2 = N '// ���� Nan,j < 0, �� ��������� N'an=N
    Else
        nan2 = M / (EMDD.Z_Anchor / 1000) - N / EMDD.nan_Anchor '// ���������� ��������� ������ � ����� ���� �������  (������� (�.4) [1])
    End If

If nan2 >= 0 Then
EMDD.Rb_gb (EMDD.GradeConcrete)

'// ���������� ������� �������� ����������� ����������� A1:
'// ��������� ������� �������� ����������� ������������:
'// ������ �� �����������:
    If (EMDD.Cy_Anchor - EMDD.Lpl_Anchor / 2) < EMDD.La_Anchor Then
        a = (EMDD.Y_Anchor + 2 * EMDD.Cy_Anchor) / 1000
        Else
            a = (EMDD.Y_Anchor + EMDD.Lpl_Anchor + 2 * EMDD.La_Anchor) / 1000
    End If

'// ������ �� ���������:
    If (EMDD.Cz_Anchor - EMDD.Lpl_Anchor / 2) < EMDD.La_Anchor Then
        If EMDD.La_Anchor < (EMDD.Z_Anchor + EMDD.Cz_Anchor) Then
            b = (EMDD.Cz_Anchor + EMDD.Lpl_Anchor / 2 + EMDD.La_Anchor) / 1000
            Else
                b = (2 * EMDD.Cz_Anchor + EMDD.Z_Anchor) / 1000
        End If
    Else
            b = (EMDD.Lpl_Anchor + 2 * EMDD.La_Anchor) / 1000
    End If
    
    X = Application.Min(EMDD.nan_Anchor - 1, Int(EMDD.La_Anchor / EMDD.lzZ_Anchor))
    Y = EMDD.La_Anchor - X * EMDD.lzZ_Anchor
    If Y > (EMDD.lzZ_Anchor - EMDD.Lpl_Anchor) And EMDD.La_Anchor <= EMDD.Z_Anchor Then
        DLpl = Y - EMDD.lzZ_Anchor + EMDD.Lpl_Anchor
        Else
            DLpl = 0
    End If
'// ������� �������� ������� � ���� �����������:
Aan = (EMDD.Lpl_Anchor ^ 2 * (X + 1) + DLpl * EMDD.Lpl_Anchor) * EMDD.n_Anchor / 1000000 '//������� �������� ������� � ���� �����������

A1 = a * b - Aan

j2 = 0.5 '//  ��� �������� ������
j3 = 1 '//��� ������� � ���������� ���� ������
e = 0 '// �������������� �� �����������

    nan3 = j2 * j3 * A1 * EMDD.Rbt_gb * 1000 / (1 + 3.5 * e / 1000 / a)
    Formula_P48_SP63 = Nanj / nan3

Else
    Formula_P48_SP63 = 0
End If
End Function

'// �������� ������ �� ����������� � ���������� �������, ��� N'an < 0  (�� �.4.7,� [2]):
'             ��������: ����������� ������ ������������ ����� (e2=0) (� ����� ���������)
Public Function Formula_P47_SP63(Q As Double, M As Double, N As Double) As Double
Application.Volatile

Dim EMDD As New C_SP63_13330_2012_Tables
Dim Formulas As New C_SP63_13330_2012_Formulas
EMDD.Table_Ini
Dim Nanj As Double, nan2 As Double, nan4 As Double
Dim e0 As Double, A1 As Double, A2 As Double, e1 As Double, e2 As Double, j2 As Double, j3 As Double
Dim Aan As Double, a As Double, DLpl As Double, X As Double, Y As Double, L1 As Double

    Nanj = Formulas.Nanj(M, N)       '// ���������� ������������� ������ � ����� ���� �������  (������� (�.2) [1])
    If (Nanj < 0) Then
        nan2 = N '// ���� Nan,j < 0, �� ��������� N'an=N
    Else
        nan2 = M / (EMDD.Z_Anchor / 1000) - N / EMDD.nan_Anchor '// ���������� ��������� ������ � ����� ���� �������  (������� (�.4) [1])
    End If

If nan2 < 0 Then
'// ���������� ������� �������� ����������� ����������� A:
    If N = 0 Then
        e0 = 0
    Else
        If M / N * 1000 <= EMDD.Z_Anchor / 2 Then
            e0 = M / N * 1000
            Else
                e0 = EMDD.Z_Anchor / 2
        End If
    End If
'// ��������� ������� �������� ����������� ������������:
'// ������ �� ���������:
    If (EMDD.Cz_Anchor - EMDD.Lpl_Anchor / 2) < EMDD.La_Anchor Then
        If EMDD.La_Anchor < (EMDD.Cz_Anchor - EMDD.Lpl_Anchor / 2 + 2 * e0) Then
            A1 = (EMDD.Cz_Anchor + EMDD.Z_Anchor + EMDD.Lpl_Anchor / 2 - 2 * e0 + EMDD.La_Anchor) / 1000
            Else
                A1 = (2 * EMDD.Cz_Anchor + EMDD.Z_Anchor) / 1000
        End If
    Else
            A1 = (EMDD.Z_Anchor + EMDD.Lpl_Anchor - 2 * e0 + 2 * EMDD.La_Anchor) / 1000
    End If
'// ������ �� �����������:
    If (EMDD.Cy_Anchor - EMDD.Lpl_Anchor / 2) < EMDD.La_Anchor Then
        A2 = (EMDD.Y_Anchor + 2 * EMDD.Cy_Anchor) / 1000
        Else
            A2 = (EMDD.Y_Anchor + EMDD.Lpl_Anchor + 2 * EMDD.La_Anchor) / 1000
    End If

'// ��������������� ������������ ������� A:
'// �� ���������:
    If EMDD.La_Anchor < (EMDD.Cz_Anchor - EMDD.Lpl_Anchor / 2) Then
        e1 = A1 * 1000 / 2 - (EMDD.La_Anchor + EMDD.Lpl_Anchor / 2 + EMDD.Z_Anchor / 2 - e0)
    Else
        e1 = A1 * 1000 / 2 - (EMDD.Cz_Anchor + EMDD.Z_Anchor / 2 - e0)
    End If
'// �� �����������:
    e2 = 0

    L1 = EMDD.Z_Anchor + EMDD.La_Anchor - 2 * e0
    X = Application.Min(EMDD.nan_Anchor - 1, Int(L1 / EMDD.lzZ_Anchor))
    Y = L1 - X * EMDD.lzZ_Anchor
    If Y > EMDD.lzZ_Anchor - EMDD.Lpl_Anchor And L1 <= EMDD.Z_Anchor Then
        DLpl = Y - EMDD.lzZ_Anchor + EMDD.Lpl_Anchor
        Else
            DLpl = 0
    End If
'// ������� �������� ������� � ���� �����������:
    Aan = (EMDD.Lpl_Anchor ^ 2 * (X + 1) + DLpl * EMDD.Lpl_Anchor) * EMDD.n_Anchor / 1000000
    a = A2 * A1 - Aan
'// ������� ����������� ������ ��� ������ N'an < 0 (�� �.32 [2]):
    j2 = 0.5 '//  ��� �������� ������

'//��� ������� � ���������� ���� ������:
    j3 = 1
    EMDD.Rb_gb (EMDD.GradeConcrete)
    nan4 = j2 * j3 * a * EMDD.Rbt_gb * 1000 / (1 + 3.5 * e1 / 1000 / A1 + 3.5 * e2 / 1000 / A2)
    
    Formula_P47_SP63 = N / nan4
            
Else
    Formula_P47_SP63 = 0
End If
End Function


'// �������� ������ �� ����������� �������� � ���� �������� (�� �.4.11 [2] � �.3.108 [3]):
'             ��������: ����������� ������ ������������ ���� (e=0) (� ����� ���������)
Public Function Formula_P411_SP63(Q As Double, M As Double, N As Double) As Double
Application.Volatile

Dim EMDD As New C_SP63_13330_2012_Tables
Dim Formulas As New C_SP63_13330_2012_Formulas
EMDD.Table_Ini
Dim h As Double, b As Double, Aout As Double, A1 As Double, a As Double, A2 As Double, dn As Double
Dim j2 As Double, e As Double, Q5 As Double

'// ���������� �������� ����������� ����������� (��. ���. 14):
'// ���������� �� �������� ���������� ���� ������� �� ���� �������� � ����������� ���������� ����:
    
    If EMDD.Y_Anchor + EMDD.Cz_Anchor > EMDD.Ha_Anchor Then '// �� ����� Ha
        h = EMDD.Ha_Anchor
    Else
        h = EMDD.Y_Anchor + EMDD.Cz_Anchor
    End If

'// ������ ��������:
    If EMDD.Cy_Anchor < h Then '// �� ����� Y+2*h
        b = EMDD.Y_Anchor + 2 * EMDD.Cy_Anchor
    Else
        b = EMDD.Y_Anchor + 2 * h
    End If
    
    A1 = (b ^ 2 - EMDD.Y_Anchor ^ 2) / 4 / 1000000
    a = (b - EMDD.Y_Anchor) / 2
        If a < h Then
            A2 = b * (h - a) / 1000000
            Else
                A2 = 0
        End If
'// ������� ����������� ������ �������� (�� �.39 [2])
    Aout = A1 + A2
'// dn - ������������ �������� �. 3.108 [3], � ������ �������� �� ��������� ������ ����� ���������� ���� Q ���������� ���� N
'//                                         ����������� �� ����� 0,2
'//                                         ��� N<0 ����������� ������ 1,0
    EMDD.Rb_gb (EMDD.GradeConcrete)
    If N > 0 Then
        dn = Application.Max(0.2, 1 - 0.3 * N / (Aout * EMDD.Rbt_gb * 1000))
        Else
            dn = 1
    End If

    j2 = 0.5 '//  ��� �������� ������
    e = 0 '//  �������������� �� �����������
    
    Q5 = j2 * EMDD.Rbt_gb * 1000 * (b / 1000) * (h / 1000) / (1 + 3.5 * e / b) * dn
    
    Formula_P411_SP63 = Q / Q5
            
End Function


