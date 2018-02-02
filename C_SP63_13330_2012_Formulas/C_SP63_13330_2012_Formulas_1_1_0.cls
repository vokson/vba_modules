VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_SP63_13330_2012_Formulas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Nanj(M As Double, N As Double) As Double
Dim EMDD As New C_SP63_13330_2012_Tables
EMDD.Table_Ini
   
    Nanj = M / (EMDD.Z_Anchor / 1000) + N / EMDD.nan_Anchor '// Наибольшее растягивающее усилие в одном ряду анкеров  (формула (Б.2) [1])
        If (Nanj < 0) Then
            Nanj = 0 '// Если Nan,j < 0, то принимаем Nan,j=0
        End If
        
End Function

Public Function Qanj(Q As Double, M As Double, N As Double) As Double
Dim EMDD As New C_SP63_13330_2012_Tables
Dim nanjC As Double, nan2 As Double
EMDD.Table_Ini
       
    nanjC = Nanj(M, N) '// Наибольшее растягивающее усилие в одном ряду анкеров  (формула (Б.2) [1])
        If (nanjC < 0) Then
            nan2 = N '// Если Nan,j < 0, то принимаем N'an=N
        Else
            nan2 = M / (EMDD.Z_Anchor / 1000) - N / EMDD.nan_Anchor '// Наибольшее сжимающее усилие в одном ряду анкеров  (формула (Б.4) [1])
                If (nan2 < 0) Then
                    nan2 = 0 '// Если N'an < 0, то принимаем N'an=0
                End If
        End If
    Qanj = (Q - 0.3 * nan2) / EMDD.nan_Anchor '// Сдвигающее усилие, приходящееся на один ряд анкеров (формула (Б.3) [1])
        If (Qanj < 0) Then
            Qanj = 0 '// Если Qan,j < 0, то принимаем Qan,j=0
        End If
End Function

Public Function lan(w As Double, Dl As Double, jc As Double, Q As Double, M As Double, N As Double) As Double

Dim EMDD As New C_SP63_13330_2012_Tables
EMDD.Table_Ini

    lan = jc * (w * Formula_B1_SP63(Q, M, N) * EMDD.Rs(EMDD.GradeAnchor) / (EMDD.Rb_gb(EMDD.GradeConcrete) / EMDD.Gb(EMDD.kConcreteGB, EMDD.kConcreteN, EMDD.kConcreteLay)) + Dl) * EMDD.dAnchor

End Function

Public Function maxM(Q As Double, M As Double, N As Double)
'// максимальное М:
    Do While Max_Formulas(Q, M, N) < 1
        M = Round(M + 0.1, 1)
    Loop
    M = M - 0.1

        Do While Max_Formulas(Q, M, N) < 1
            M = Round(M + 0.01, 2)
        Loop
        M = M - 0.01

            Do While Max_Formulas(Q, M, N) < 1
                M = Round(M + 0.001, 3)
            Loop
    maxM = M - 0.001
End Function

Public Function maxQ(Q As Double, M As Double, N As Double)
'// максимальное Q:
    Do While Max_Formulas(Q, M, N) < 1
        Q = Q + 10
    Loop
    Q = Q - 10

        Do While Max_Formulas(Q, M, N) < 1
            Q = Q + 1
        Loop
        Q = Q - 1
            Do While Max_Formulas(Q, M, N) < 1
                Q = Q + 0.1
            Loop
    maxQ = Q - 0.1
End Function

Public Function maxQ_End(Q As Double, M As Double, N As Double)
'// максимальное Q:
    Do While Max_Formulas(Q, M, N) < 1
        Q = Q + 10
        M = Q * 0.01
    Loop
    Q = Q - 10

        Do While Max_Formulas(Q, M, N) < 1
            Q = Q + 1
            M = Q * 0.01
        Loop
        Q = Q - 1
            Do While Max_Formulas(Q, M, N) < 1
                Q = Q + 0.1
                M = Q * 0.01
        Loop
    maxQ_End = Q - 0.1
End Function

'// Проверка прочности анкеров закладной деталли (по Приложению Б, п.Б.1 [1]):
Public Function Formula_B1_SP63(Q As Double, M As Double, N As Double) As Double
Application.Volatile

Dim EMDD As New C_SP63_13330_2012_Tables
EMDD.Table_Ini
Dim Qanj0 As Double, Nanj0 As Double, gssh As Double, NanjF As Double, QanjF As Double
       
gssh = 1.65
    Qanj0 = gssh * EMDD.Aanj_Anchor(EMDD.dAnchor) * Sqr(EMDD.Rs(EMDD.GradeAnchor) * EMDD.Rb_gb(EMDD.GradeConcrete)) * 1000 '// Сдвигающая сила, воспринимаемая анкерами (формула (Б.5) [1])
    Nanj0 = EMDD.Rs(EMDD.GradeAnchor) * EMDD.Aanj_Anchor(EMDD.dAnchor) * 1000 '// Предельная растягивающая сила, воспринимаемая одним рядом анкеров (формула (Б.6) [1])
    NanjF = Nanj(M, N)       '// Наибольшее растягивающее усилие в одном ряду анкеров  (формула (Б.2) [1])
    QanjF = Qanj(Q, M, N)      '// Сдвигающее усилие, приходящееся на один ряд анкеров (формула (Б.3) [1])
          
    Formula_B1_SP63 = QanjF / Qanj0 + NanjF / Nanj0
End Function


'// Проверка смятия бетона под пластинами усиления (по п.4.12 [2]):
Public Function Formula_412_SP63(Q As Double, M As Double, N As Double) As Double
Application.Volatile

Dim EMDD As New C_SP63_13330_2012_Tables
EMDD.Table_Ini
Dim NanjF As Double, QanjF As Double, Nan1 As Double, Qan1 As Double, w As Double, Dl As Double, jc As Double, jb As Double
Dim lanF As Double, Nloc As Double, Rbloc As Double, Aloc As Double, Ab As Double, Abmax As Double
Dim minL As Double

    NanjF = Nanj(M, N)       '// Наибольшее растягивающее усилие в одном ряду анкеров  (формула (Б.2) [1])
    QanjF = Qanj(Q, M, N)      '// Сдвигающее усилие, приходящееся на один ряд анкеров (формула (Б.3) [1])
    Nan1 = NanjF / EMDD.n_Anchor '// Растягивающие и сдвигающие усилия в одном анкерном стержне
    Qan1 = QanjF / EMDD.n_Anchor '// Растягивающие и сдвигающие усилия в одном анкерном стержне
    w = 0.7
    Dl = 11
    If Nan1 <= 0 Then
        jc = 0.7
        Else
            jc = 0.3 / (1 + Qan1 / Nan1) + 0.7
    End If
    
    lanF = lan(w, Dl, jc, Q, M, N) '//Требуемая длина анкеровки lan арматурного стержня периодического профиля без усиления (п.5.7 [2]):
    If lanF < EMDD.dAnchor * 20 And lanF < 250 Then
        If EMDD.dAnchor > 250 Then
            lanF = EMDD.dAnchor * 20
            Else
                lanF = 250
        End If
    End If
        
    If (EMDD.La_Anchor < 15 * EMDD.dAnchor) Then '// Локальная сила на усиление анкера Nloc
        Nloc = Nan1 + Qan1 * (15 * EMDD.dAnchor - EMDD.La_Anchor) / lanF '// Для La (длина анкера) < 15d (Диаметр анкера (мм))
    Else
        Nloc = Nan1 '// Для La (длина анкера) >= 15d (Диаметр анкера (мм))(ф. 42 [2])
    End If

    EMDD.Aanj_Anchor (EMDD.dAnchor)
    Aloc = (EMDD.Lpl_Anchor / 1000) ^ 2 - EMDD.Aan1 / 10000 '// Площадь смятия

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

    If 0.8 * Sqr(Abmax / Aloc) > 1 Then '// (ф.8.82 [1]) но не менее 1,0 и не более 2,5
        If 0.8 * Sqr(Abmax / Aloc) < 2.5 Then
            jb = 0.8 * Sqr(Abmax / Aloc)
            Else
                jb = 2.5
        End If
        Else
            jb = 1
    End If
    Rbloc = EMDD.Rb_gb(EMDD.GradeConcrete) * jb '// (ф. 8.81 [1])
Formula_412_SP63 = Nloc / (Rbloc * Aloc * 1000)  '// Коэффициент использования; Условие смятия (по ф.41 [2] с учетом требований п.8.1.44 [1])

End Function

'// Проверка бетона на выкалывание с усилениями анкеров, при N'an > 0  (по п.4.8 [2]):
'            ВНИМАНИЕ: Учитывается только симметричный отрыв (e=0) (в запас прочности)
Public Function Formula_P48_SP63(Q As Double, M As Double, N As Double) As Double
Application.Volatile

Dim EMDD As New C_SP63_13330_2012_Tables
'Dim Formulas As New C_SP63_13330_2012_Formulas
EMDD.Table_Ini
Dim NanjF As Double, nan3 As Double
Dim a As Double, b As Double, A1 As Double, Aan As Double
Dim j2 As Double, j3 As Double, e As Double, X As Double, Y As Double, DLpl As Double, nan2 As Double

    NanjF = Nanj(M, N)       '// Наибольшее растягивающее усилие в одном ряду анкеров  (формула (Б.2) [1])
    If (NanjF < 0) Then
        nan2 = N '// Если Nan,j < 0, то принимаем N'an=N
    Else
        nan2 = M / (EMDD.Z_Anchor / 1000) - N / EMDD.nan_Anchor '// Наибольшее сжимающее усилие в одном ряду анкеров  (формула (Б.4) [1])
    End If

If nan2 >= 0 Then
EMDD.Rb_gb (EMDD.GradeConcrete)

'// Определяем площадь проекции поверхности выкалывания A1:
'// Принимаем размеры проекции поверхности выдавливания:
'// размер по горизонтали:
    If (EMDD.Cy_Anchor - EMDD.Lpl_Anchor / 2) < EMDD.La_Anchor Then
        a = (EMDD.Y_Anchor + 2 * EMDD.Cy_Anchor) / 1000
        Else
            a = (EMDD.Y_Anchor + EMDD.Lpl_Anchor + 2 * EMDD.La_Anchor) / 1000
    End If

'// размер по вертикали:
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
'// Площадь анкерных пластин в зоне выкалывания:
Aan = (EMDD.Lpl_Anchor ^ 2 * (X + 1) + DLpl * EMDD.Lpl_Anchor) * EMDD.n_Anchor / 1000000 '//Площадь анкерных пластин в зоне выкалывания

A1 = a * b - Aan

j2 = 0.5 '//  для тяжелого бетона
j3 = 1 '//Для заделки в растянутой зоне бетона
e = 0 '// эксцентриситет по горизонтали

    nan3 = j2 * j3 * A1 * EMDD.Rbt_gb * 1000 / (1 + 3.5 * e / 1000 / a)
    Formula_P48_SP63 = NanjF / nan3

Else
    Formula_P48_SP63 = 0
End If
End Function

'// Проверка бетона на выкалывание с усилениями анкеров, при N'an < 0  (по п.4.7,а [2]):
'             ВНИМАНИЕ: Учитывается только симметричный отрыв (e2=0) (в запас прочности)
Public Function Formula_P47_SP63(Q As Double, M As Double, N As Double) As Double
Application.Volatile

Dim EMDD As New C_SP63_13330_2012_Tables
'Dim Formulas As New C_SP63_13330_2012_Formulas
EMDD.Table_Ini
Dim NanjF As Double, nan2 As Double, nan4 As Double
Dim e0 As Double, A1 As Double, A2 As Double, e1 As Double, e2 As Double, j2 As Double, j3 As Double
Dim Aan As Double, a As Double, DLpl As Double, X As Double, Y As Double, L1 As Double

    NanjF = Nanj(M, N)       '// Наибольшее растягивающее усилие в одном ряду анкеров  (формула (Б.2) [1])
    If (NanjF < 0) Then
        nan2 = N '// Если Nan,j < 0, то принимаем N'an=N
    Else
        nan2 = M / (EMDD.Z_Anchor / 1000) - N / EMDD.nan_Anchor '// Наибольшее сжимающее усилие в одном ряду анкеров  (формула (Б.4) [1])
    End If

If nan2 < 0 Then
'// Определяем площадь проекции поверхности выкалывания A:
    If N = 0 Then
        e0 = 0
    Else
        If M / N * 1000 <= EMDD.Z_Anchor / 2 Then
            e0 = M / N * 1000
            Else
                e0 = EMDD.Z_Anchor / 2
        End If
    End If
'// Принимаем размеры проекции поверхности выдавливания:
'// размер по вертикали:
    If (EMDD.Cz_Anchor - EMDD.Lpl_Anchor / 2) < EMDD.La_Anchor Then
        If EMDD.La_Anchor < (EMDD.Cz_Anchor - EMDD.Lpl_Anchor / 2 + 2 * e0) Then
            A1 = (EMDD.Cz_Anchor + EMDD.Z_Anchor + EMDD.Lpl_Anchor / 2 - 2 * e0 + EMDD.La_Anchor) / 1000
            Else
                A1 = (2 * EMDD.Cz_Anchor + EMDD.Z_Anchor) / 1000
        End If
    Else
            A1 = (EMDD.Z_Anchor + EMDD.Lpl_Anchor - 2 * e0 + 2 * EMDD.La_Anchor) / 1000
    End If
'// размер по горизонтали:
    If (EMDD.Cy_Anchor - EMDD.Lpl_Anchor / 2) < EMDD.La_Anchor Then
        A2 = (EMDD.Y_Anchor + 2 * EMDD.Cy_Anchor) / 1000
        Else
            A2 = (EMDD.Y_Anchor + EMDD.Lpl_Anchor + 2 * EMDD.La_Anchor) / 1000
    End If

'// Эксцентриситеты относительно площади A:
'// по вертикали:
    If EMDD.La_Anchor < (EMDD.Cz_Anchor - EMDD.Lpl_Anchor / 2) Then
        e1 = A1 * 1000 / 2 - (EMDD.La_Anchor + EMDD.Lpl_Anchor / 2 + EMDD.Z_Anchor / 2 - e0)
    Else
        e1 = A1 * 1000 / 2 - (EMDD.Cz_Anchor + EMDD.Z_Anchor / 2 - e0)
    End If
'// по горизонтали:
    e2 = 0

    L1 = EMDD.Z_Anchor + EMDD.La_Anchor - 2 * e0
    X = Application.Min(EMDD.nan_Anchor - 1, Int(L1 / EMDD.lzZ_Anchor))
    Y = L1 - X * EMDD.lzZ_Anchor
    If Y > EMDD.lzZ_Anchor - EMDD.Lpl_Anchor And L1 <= EMDD.Z_Anchor Then
        DLpl = Y - EMDD.lzZ_Anchor + EMDD.Lpl_Anchor
        Else
            DLpl = 0
    End If
'// Площадь анкерных пластин в зоне выкалывания:
    Aan = (EMDD.Lpl_Anchor ^ 2 * (X + 1) + DLpl * EMDD.Lpl_Anchor) * EMDD.n_Anchor / 1000000
    a = A2 * A1 - Aan
'// Условие выкалывания бетона для случая N'an < 0 (по ф.32 [2]):
    j2 = 0.5 '//  для тяжелого бетона

'//Для заделки в растянутой зоне бетона:
    j3 = 1
    EMDD.Rb_gb (EMDD.GradeConcrete)
    nan4 = j2 * j3 * a * EMDD.Rbt_gb * 1000 / (1 + 3.5 * e1 / 1000 / A1 + 3.5 * e2 / 1000 / A2)
    
    Formula_P47_SP63 = N / nan4
            
Else
    Formula_P47_SP63 = 0
End If
End Function


'// Проверка бетона на откалывание анкерами у края элемента (по п.4.11 [2] и п.3.108 [3]):
'             ВНИМАНИЕ: Учитывается только симметричный срез (e=0) (в запас прочности)
Public Function Formula_P411_SP63(Q As Double, M As Double, N As Double) As Double
Application.Volatile

Dim EMDD As New C_SP63_13330_2012_Tables
EMDD.Table_Ini
Dim h As Double, b As Double, Aout As Double, A1 As Double, a As Double, A2 As Double, dn As Double
Dim j2 As Double, e As Double, Q5 As Double

'// Определяем габариты поверхности откалывания (см. рис. 14):
'// Расстояние от наиболее удаленного ряда анкеров до края элемента в направлении сдвигающей силы:
    
    If EMDD.Y_Anchor + EMDD.Cz_Anchor > EMDD.Ha_Anchor Then '// Не более Ha
        h = EMDD.Ha_Anchor
    Else
        h = EMDD.Y_Anchor + EMDD.Cz_Anchor
    End If

'// Ширина элемента:
    If EMDD.Cy_Anchor < h Then '// не более Y+2*h
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
'// Условие откалывания бетона анкерами (по ф.39 [2])
    Aout = A1 + A2
'// dn - определяется согласно п. 3.108 [3], в случае действия на закладную деталь кроме сдвигающей силы Q отрывающей силы N
'//                                         принимается не менее 0,2
'//                                         при N<0 принимается равным 1,0
    EMDD.Rb_gb (EMDD.GradeConcrete)
    If N > 0 Then
        dn = Application.Max(0.2, 1 - 0.3 * N / (Aout * EMDD.Rbt_gb * 1000))
        Else
            dn = 1
    End If

    j2 = 0.5 '//  для тяжелого бетона
    e = 0 '//  эксцентриситет по горизонтали
    
    Q5 = j2 * EMDD.Rbt_gb * 1000 * (b / 1000) * (h / 1000) / (1 + 3.5 * e / b) * dn
    
    Formula_P411_SP63 = Q / Q5
            
End Function

Public Function Max_Formulas(Q As Double, M As Double, N As Double) As Double
Max_Formulas = Application.Max(Formula_B1_SP63(Q, M, N), Formula_412_SP63(Q, M, N), Formula_P48_SP63(Q, M, N), Formula_P47_SP63(Q, M, N), Formula_P411_SP63(Q, M, N))
End Function



