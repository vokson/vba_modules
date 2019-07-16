VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_SP24_13330_2011_Formulas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const VERSION = "1.2.0"

' СП 24.13330.2011 (изм.1) - Формула 7.26
Function f7_12(Y1 As Double, Y1_ As Double, h As Double, d As Double, _
         a1 As Double, a2 As Double, a3 As Double, a4 As Double) As Double
    f7_12 = 0.75 * a4 * (a1 * Y1_ * d + a2 * a3 * Y1 * h)
End Function

' СП 24.13330.2011 (изм.1) - Формула 7.26
Function f7_26(B1 As Double, qs As Double) As Double
    f7_26 = B1 * qs
End Function

' СП 24.13330.2011 (изм.1) - Формула 7.28
Function f7_28(sum_Bi_Fsi_Hi As Double, h As Double) As Double
    f7_28 = sum_Bi_Fsi_Hi / h
End Function

' СП 24.13330.2011 (изм.1) - Формула 7.32
Function f7_32(force As Double, betta As Double, G1 As Double, length As Double) As Double
    f7_32 = betta * Abs(force) / G1 / length
End Function

' СП 24.13330.2011 (изм.1) - Формула 7.33
Function f7_33(G1 As Double, G2 As Double, nu1 As Double, nu2 As Double, _
                            EA As Double, d As Double, length As Double) As Double
   
   Dim knu As Double, knu1 As Double, lambda1 As Double, ksi As Double
   Dim alpha_dash As Double, betta_dash As Double, betta As Double
   
   knu = f7_35((nu1 + nu2) / 2)
   knu1 = f7_35(nu1)
   
   ksi = EA / G1 / length ^ 2
   lambda1 = f7_34(ksi)
   
   alpha_dash = 0.17 * Log(knu1 * length / d)
   betta_dash = 0.17 * Log(knu * G1 * length / G2 / d)
   
   f7_33 = betta_dash / lambda1 + (1 - (betta_dash / alpha_dash)) / ksi
   
'   Debug.Print "knu = " & knu
'   Debug.Print "knu1 = " & knu1
'   Debug.Print "ksi = " & ksi
'   Debug.Print "lambda1 = " & lambda1
'   Debug.Print "alpha_dash = " & alpha_dash
'   Debug.Print "betta_dash = " & betta_dash
   
End Function

Function f7_34(ksi As Double) As Double
   f7_34 = 2.12 * ksi ^ 0.75 / (1 + 2.12 * ksi ^ 0.75)
End Function

Function f7_35(nu As Double) As Double
   f7_35 = 2.82 - 3.78 * nu + 2.18 * nu ^ 2
End Function

' Формула В.3
' K [кН/м4]
' bp [м]
' E [кПа]
' I [м4]
Public Function fC_3(K As Double, bp As Double, Yc as Double, E As Double, I As Double) As Double ' [1/м]

    If K = 0 Or bp = 0 Or Yc = 0 Then
      MsgBox "SP24.13330.2011 Formulas В.3 - Yc, K and bp shall not be zero"
      Exit Function
    End If

    fC_3 = (K * bp / Yc / E / I) ^ (1 / 5)
End Function