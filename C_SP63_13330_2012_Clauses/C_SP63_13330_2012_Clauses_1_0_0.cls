VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_SP63_13330_2012_Clauses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' СП 63.13330.2012 (изм.2) - Пункт 8.1.34
' stress - (-) сжатие, (+) растяжение
Function c8_1_34(ByVal stress As Double, ByVal Rb As Double,ByVal  Rbt As Double) As Double
    
  If (stress < 0) Then

    stress = Abs(stress)
    If 0 <= stress And stress <= 0.25 * Rb Then c8_1_34 = 1 + stress/Rb
    If 0.25 * Rb < stress And stress <= 0.75 * Rb Then c8_1_34 = 1.25
    If 0.75 * Rb < stress And stress <= Rb Then c8_1_34 = 5 * (1 - stress / Rb)
    If stress > Rb Then c8_1_34 = 0#

  ElseIf (stress > 0) Then

    stress = Abs(stress)
    If 0 <= stress And stress <= Rbt Then c8_1_34 = 1 - stress / 2 / Rbt
    If stress > Rbt Then c8_1_34 = 0#

  Else

    c8_1_34 = 1#

  End If

End Function