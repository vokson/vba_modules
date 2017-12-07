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