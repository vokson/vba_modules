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

' СП 63.13330.2012 (изм.2) - Формула 8.55
' Rb - [MPa] = N/mm2
' b - [mm]
' h0 - [mm]
' FIn - [-]
' RESULT - [N]
Function f8_55(Rb As Double, b As Double, h0 As Double, FIn As Double) As Double
    f8_55 = 0.3 * Rb * b * h0 * FIn
End Function

' Rbt - [N/mm2]
' q_sw - [N/mm]
' b - [mm]
' h0 - [mm]
' FIn - [-]
' RESULT - [N]

' Use C_Math [1.2.0]

Public Function f8_56( _
    Rbt As Double, _
    q_sw As Double, _
    b As Double, _
    h0 As Double, _
    FIn As Double _
) As Double

    Dim math As New C_Math

    Dim Q As Double
    Dim Qb As Double
    Dim Qsw As Double
    Dim C As Double
    Dim isFirst As Boolean

    isFirst = True

    For C = h0 To 2 * h0 Step 0.05 * h0

        If q_sw < 0.25 * Rbt * b Then
            Qb = 4 * 1.5 * h0 ^ 2 * q_sw / C * FIn
        Else
            Qb = getFormula8_57(Rbt, b, h0, C, FIn)
        End If

        Qb = math.max(Qb, 0.5 * Rbt * b * h0)
        Qb = math.min(Qb, 2.5 * Rbt * b * h0)
      
        Qsw = getFormula8_58(q_sw, C)
        
        If isFirst = True Then
            Q = Qb + Qsw
            isFirst = False
        Else
            Q = math.min(Q, Qb + Qsw)
        End If

    Next C

    f8_56 = Q

    Set math = Nothing

End Function

' СП 63.13330.2012 (изм.2) - Формула 8.57
' Rbt- [MPa] = N/mm2
' b - [mm]
' h0 - [mm]
' C - [mm]
' FIn - [-]
' RESULT - [N]
Function f8_57(Rbt As Double, b As Double, h0 As Double, C As Double, FIn As Double) As Double
    f8_57 = 1.5 * Rbt * b * h0 ^ 2 / C * FIn
End Function

' СП 63.13330.2012 (изм.2) - Формула 8.58
' qsw- [N/mm]
' C - [mm]
' RESULT - [N]
Function f8_58(qsw As Double, C As Double) As Double
    f8_58 = 0.75 * qsw * C
End Function

' СП 63.13330.2012 (изм.2) - Формула 8.59
' Rsw - [N/mm2]
' Asw - [mm2]
' sw - [mm]
' RESULT - [N/mm]
Function f8_59(Rsw As Double, Asw As Double, sw As Double) As Double
    f8_59 = Rsw * Asw / sw
End Function