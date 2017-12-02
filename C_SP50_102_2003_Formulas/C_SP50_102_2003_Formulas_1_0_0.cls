VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_SP50_102_2003_Formulas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private math As C_Math

Private Sub Class_Initialize()
    Set math = New C_Math
End Sub

Private Sub Class_Terminate()
    Set math = Nothing
End Sub

' Таблица Д.2
' pileCase As Integer
'    0 - опирание сваи на дисперсный грунт
'    1 - опирание сваи на скальный грунт
'    2 - заделка сваи в скальный грунт
' length As Double
' parameter As String - "A0", "B0", "C0" (англ.)
Public Function E2(pileCase As Integer, length As Double, parameter As String) As Double
    
End Function