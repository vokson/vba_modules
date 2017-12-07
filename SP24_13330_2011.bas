Attribute VB_Name = "SP24_13330_2011"
Option Explicit

Function getClause7_4_3(e As Double, nu As Double) As Double
    getClause7_4_3 = e / 2 / (1 + nu)
End Function

Function getFormula7_32(force As Double, G1 As Double, G2 As Double, nu1 As Double, nu2 As Double, _
                            EA As Double, d As Double, length As Double) As Double
    
    Dim betta As Double
    
    betta = getFormula7_33(G1, G2, nu1, nu2, EA, d, length)
    Debug.Print "betta = " & betta
    
    getFormula7_32 = betta * Abs(force) / G1 / length
   
End Function

Function getFormula7_33(G1 As Double, G2 As Double, nu1 As Double, nu2 As Double, _
                            EA As Double, d As Double, length As Double) As Double
   
   Dim knu As Double, knu1 As Double, lambda1 As Double, ksi As Double
   Dim alpha_dash As Double, betta_dash As Double, betta As Double
   
   knu = getFormula7_35((nu1 + nu2) / 2)
   knu1 = getFormula7_35(nu1)
   
   ksi = EA / G1 / length ^ 2
   lambda1 = getFormula7_34(ksi)
   
   alpha_dash = 0.17 * Log(knu1 * length / d)
   betta_dash = 0.17 * Log(knu * G1 * length / G2 / d)
   
   getFormula7_33 = betta_dash / lambda1 + (1 - (betta_dash / alpha_dash)) / ksi
   
'   Debug.Print "knu = " & knu
'   Debug.Print "knu1 = " & knu1
'   Debug.Print "ksi = " & ksi
'   Debug.Print "lambda1 = " & lambda1
'   Debug.Print "alpha_dash = " & alpha_dash
'   Debug.Print "betta_dash = " & betta_dash
   
End Function

Function getFormula7_34(ksi As Double) As Double
   getFormula7_34 = 2.12 * ksi ^ 0.75 / (1 + 2.12 * ksi ^ 0.75)
End Function

Function getFormula7_35(nu As Double) As Double
   getFormula7_35 = 2.82 - 3.78 * nu + 2.18 * nu ^ 2
End Function



Function getFormula7_26(B1 As Double, qs As Double) As Double
    getFormula7_26 = B1 * qs
End Function

Function getFormula7_28(sum_Bi_Fsi_Hi As Double, h As Double) As Double
    getFormula7_28 = sum_Bi_Fsi_Hi / h
End Function


