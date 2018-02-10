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
Function f8_55(Rb As Double, b As Double, h0 As Double, FIn As Double) As Double
    f8_55 = 0.3 * Rb * b * h0 * FIn
End Function

' СП 63.13330.2012 (изм.2) - Формула 8.57
Function f8_57(Rbt As Double, b As Double, h0 As Double, C As Double, FIn As Double) As Double
    f8_57 = 1.5 * Rbt * b * h0 ^ 2 / C * FIn
End Function

' СП 63.13330.2012 (изм.2) - Формула 8.58
Function f8_58(qsw As Double, C As Double) As Double
    f8_58 = 0.75 * qsw * C
End Function

' СП 63.13330.2012 (изм.2) - Формула 8.59
Function f8_59(Rsw As Double, Asw As Double, sw As Double) As Double
    f8_59 = Rsw * Asw / sw
End Function