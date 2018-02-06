Attribute VB_Name = "M_SP22_13330_2011_Test"
Option Explicit

Public Function getTables5_10(typeOfSoil As String, IL As Double) As Double

    Dim sp As New C_SP22_13330_2011
    Dim soil As New C_Soil

    soil.ClassOfSoil = "ƒ»—œ≈–—Õ€…"
    soil.TypeBySize = typeOfSoil
    soil.LiquidityIndex = IL

    getTables5_10 = sp.Tables.t5_10(soil)

    Set sp = Nothing
    Set soil = Nothing
End Function

Public Function getTables5_4_Yc1(typeOfSoil As String, subtypeOfSoil As String, typeDensity As String, Saturation As String, IL As Double) As Double

    Dim sp As New C_SP22_13330_2011
    Dim soil As New C_Soil

    soil.ClassOfSoil = "ƒ»—œ≈–—Õ€…"
    soil.TypeBySize = typeOfSoil
    soil.SubtypeBySize = subtypeOfSoil
    soil.TypeByDensity = typeDensity
    soil.TypeByDegreeOfSaturation = Saturation
    soil.LiquidityIndex = IL

    getTables5_4_Yc1 = sp.Tables.t5_4_Yc1(soil)

    Set sp = Nothing
    Set soil = Nothing
End Function

Public Function getTables5_4_Yc2(typeOfSoil As String, subtypeOfSoil As String, typeDensity As String, Saturation As String, IL As Double, L_H As Double, ModelFlexible As Boolean) As Double

    Dim sp As New C_SP22_13330_2011
    Dim soil As New C_Soil

    soil.ClassOfSoil = "ƒ»—œ≈–—Õ€…"
    soil.TypeBySize = typeOfSoil
    soil.SubtypeBySize = subtypeOfSoil
    soil.TypeByDensity = typeDensity
    soil.TypeByDegreeOfSaturation = Saturation
    soil.LiquidityIndex = IL
    

    getTables5_4_Yc2 = sp.Tables.t5_4_Yc2(soil, L_H, ModelFlexible)

    Set sp = Nothing
    Set soil = Nothing
End Function

