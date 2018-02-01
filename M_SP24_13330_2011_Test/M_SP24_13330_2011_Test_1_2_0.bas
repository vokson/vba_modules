Attribute VB_Name = "M_SP24_13330_2011_Test"
Option Explicit

Public Function getTable7_2_Note_2(originalSoilElevation As Double, finalSoilElevation As Double) As Double
    Dim sp As New C_SP24_13330_2011
    getTable7_2_Note_2 = sp.Tables.t7_2_Note_2(originalSoilElevation, finalSoilElevation)
    Set sp = Nothing
End Function


Public Function getTable7_2_Note_4( _
    typeOfSoil As String, _
    subtypeOfSoil As String, _
    densityOfSoil As String, _
    Optional isDensityAsPerCPTWithoutHoles As Boolean _
) As Double

    Dim sp As New C_SP24_13330_2011
    Dim soil As New C_Soil

    soil.ClassOfSoil = "ƒ»—œ≈–—Õ€…"
    soil.TypeBySize = typeOfSoil
    soil.SubtypeBySize = subtypeOfSoil
    soil.TypeByDensity = densityOfSoil

    getTable7_2_Note_4 = sp.Tables.t7_2_Note_4(soil, isDensityAsPerCPTWithoutHoles)

    Set sp = Nothing
    Set soil = Nothing

End Function

Function getTable7_2_forSand(depth As Double, subtypeOfSoil As String) As Double

    Dim sp As New C_SP24_13330_2011
    Dim soil As New C_Soil

    soil.ClassOfSoil = "ƒ»—œ≈–—Õ€…"
    soil.TypeBySize = "œ≈—Œ "
    soil.SubtypeBySize = subtypeOfSoil

    getTable7_2_forSand = sp.Tables.t7_2_forSand(depth, soil.SubtypeBySize)

    Set sp = Nothing
    Set soil = Nothing

End Function

Function getTable7_2_forClay(depth As Double, IL As Double) As Double

    Dim sp As New C_SP24_13330_2011
    Dim soil As New C_Soil

    soil.ClassOfSoil = "ƒ»—œ≈–—Õ€…"
    soil.TypeBySize = "√À»Õ¿"
    soil.LiquidityIndex = IL

    getTable7_2_forClay = sp.Tables.t7_2_forClay(depth, soil.LiquidityIndex)

    Set sp = Nothing
    Set soil = Nothing

End Function

Function getTable7_2(depth As Double, typeOfSoil As String, subtypeOfSoil As String, _
    density As String, Optional IL As Double = 0, Optional IP As Double = 0, Optional e As Double = 0, _
    Optional isDensityAsPerCPTWithoutHoles As Boolean = False) As Double

    Dim sp As New C_SP24_13330_2011
    Dim soil As New C_Soil

    soil.ClassOfSoil = "ƒ»—œ≈–—Õ€…"
    soil.TypeBySize = typeOfSoil
    soil.SubtypeBySize = subtypeOfSoil
    soil.TypeByDensity = density
    soil.PlasticityIndex = IP
    soil.LiquidityIndex = IL
    soil.VoidRatio = e

    getTable7_2 = sp.Tables.t7_2(depth, soil, isDensityAsPerCPTWithoutHoles)

    Set sp = Nothing
    Set soil = Nothing

End Function

Function getTable7_3_forClay(depth As Double, IL As Double) As Double

    Dim sp As New C_SP24_13330_2011
    Dim soil As New C_Soil

    soil.ClassOfSoil = "ƒ»—œ≈–—Õ€…"
    soil.TypeBySize = "√À»Õ¿"
    soil.LiquidityIndex = IL

    getTable7_3_forClay = sp.Tables.t7_3_forClay(depth, soil.LiquidityIndex)

    Set sp = Nothing
    Set soil = Nothing

End Function

Function getTable7_3_forSand(depth As Double, subtypeOfSoil As String) As Double

    Dim sp As New C_SP24_13330_2011
    getTable7_3_forSand = sp.Tables.t7_3_forSand(depth, subtypeOfSoil)
    Set sp = Nothing
End Function

Function getTable7_3_Note_3(typeOfSoil As String, densityOfSoil As String) As Double
    Dim sp As New C_SP24_13330_2011
    getTable7_3_Note_3= sp.Tables.t7_3_Note_3(typeOfSoil, densityOfSoil)
    Set sp = Nothing
End Function

Function getTable7_3_Note_4(typeOfSoil As String, e As Double) As Double
    Dim sp As New C_SP24_13330_2011
    getTable7_3_Note_4= sp.Tables.t7_3_Note_4(typeOfSoil, e)
    Set sp = Nothing
End Function

Function getTable7_3(depth As Double, typeOfSoil As String, subtypeOfSoil As String, _
    density As String, Optional IL As Double = 0, Optional e As Double = 0) As Double

    Dim sp As New C_SP24_13330_2011
    Dim soil As New C_Soil

    soil.ClassOfSoil = "ƒ»—œ≈–—Õ€…"
    soil.TypeBySize = typeOfSoil
    soil.SubtypeBySize = subtypeOfSoil
    soil.TypeByDensity = density
    soil.LiquidityIndex = IL
    soil.VoidRatio = e

    getTable7_3 = sp.Tables.t7_3(depth, soil)

    Set sp = Nothing
    Set soil = Nothing

End Function

Function getTable7_7(what As String, Fi As Double, Optional d As Double, Optional h As Double) As Double

    Dim sp As New C_SP24_13330_2011

    Select Case what
        Case "A1": getTable7_7 = sp.Tables.t7_7_a1(Fi)
        Case "A2": getTable7_7 = sp.Tables.t7_7_a2(Fi)
        Case "A3": getTable7_7 = sp.Tables.t7_7_a3(h, d, Fi)
        Case "A4": getTable7_7 = sp.Tables.t7_7_a4(d, Fi)
    End Select

    Set sp = Nothing

End Function

Function getTable7_8(h As Double, IL As Double) As Double

    Dim sp As New C_SP24_13330_2011
    getTable7_8 = sp.Tables.t7_8(h, IL)
    Set sp = Nothing

End Function

Function getTable7_16_B1(qs As Double, isDrivenPile As Boolean, isCompressed As Boolean, _
    isScrewPileInSandyWaterSaturatedSoil As Boolean) As Double

    Dim sp As New C_SP24_13330_2011
    getTable7_16_B1 = sp.Tables.t7_16_B1 (qs, isDrivenPile, isCompressed, isScrewPileInSandyWaterSaturatedSoil)
    Set sp = Nothing
End Function

Function getTable7_16_B2i(fs As Double, typeOfZond As Integer, isSand As Boolean) As Double
    Dim sp As New C_SP24_13330_2011
    getTable7_16_B2i = sp.Tables.t7_16_B2i (fs, typeOfZond, isSand)
    Set sp = Nothing
End Function  

Function getFormula7_12(Y1 As Double, Y1_ As Double, h As Double, d As Double, _
         a1 As Double, a2 As Double, a3 As Double, a4 As Double) As Double
    Dim sp As New C_SP24_13330_2011
    getFormula7_12 = sp.Formulas.f7_12 (Y1, Y1_, h, d, a1, a2, a3, a4)
    Set sp = Nothing
End Function

Function getFormula7_26(B1 As Double, qs As Double) As Double
    Dim sp As New C_SP24_13330_2011
    getFormula7_26 = sp.Formulas.f7_26 (B1, qs)
    Set sp = Nothing
End Function

Function getFormula7_28(sum_Bi_Fsi_Hi As Double, h As Double) As Double
    Dim sp As New C_SP24_13330_2011
    getFormula7_28 = sp.Formulas.f7_28(sum_Bi_Fsi_Hi, h)
    Set sp = Nothing
End Function

Function getFormula7_32(force As Double, betta As Double, G1 As Double, length As Double) As Double

    Dim sp As New C_SP24_13330_2011
    getFormula7_32 = sp.Formulas.f7_32(force, betta, G1, length)
    Set sp = Nothing

End Function

Function getFormula7_33(G1 As Double, G2 As Double, nu1 As Double, nu2 As Double, _
                            EA As Double, d As Double, length As Double) As Double
   
    Dim sp As New C_SP24_13330_2011
    getFormula7_33 = sp.Formulas.f7_33(G1, G2, nu1, nu2, EA, d, length)
    Set sp = Nothing
   
End Function

Function getFormula7_34(ksi As Double) As Double
    Dim sp As New C_SP24_13330_2011
    getFormula7_34 = sp.Formulas.f7_34(ksi)
    Set sp = Nothing
End Function

Function getFormula7_35(nu As Double) As Double
    Dim sp As New C_SP24_13330_2011
    getFormula7_35 = sp.Formulas.f7_35(nu)
    Set sp = Nothing
End Function

Function getClause7_4_3(e As Double, nu As Double) As Double
    Dim sp As New C_SP24_13330_2011
    getClause7_4_3 = sp.Clauses.c7_4_3(e, nu)
    Set sp = Nothing
End Function

Function getTableG_1( _
    depth As Double, _
    typeBySize As String, _
    subtypeBySize As String, _
    IL As Double, _
    D As Double, _
    Sr As Double _
) As Double

    Dim sp As New C_SP24_13330_2011
    Dim soil As New C_Soil

    soil.ClassOfSoil = "ƒ»—œ≈–—Õ€…"
    soil.TypeBySize = typeBySize
    soil.SubtypeBySize = subtypeBySize
    soil.LiquidityIndex = IL
    soil.GranulationFactor = D
    soil.DegreeOfSaturation = Sr

    getTableG_1 = sp.Tables.tG_1(depth, soil)

    Set sp = Nothing
    Set soil = Nothing

End Function