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
