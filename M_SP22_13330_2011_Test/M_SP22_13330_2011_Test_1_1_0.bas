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

Public Function getTables5_5_My(InternalFrictionAngle_2 As Double)
    Dim sp As New C_SP22_13330_2011

    getTables5_5_My = sp.Tables.t5_5("My", InternalFrictionAngle_2)
    
    Set sp = Nothing
End Function

Public Function getTables5_5_Mq(InternalFrictionAngle_2 As Double)
    Dim sp As New C_SP22_13330_2011
    
    getTables5_5_Mq = sp.Tables.t5_5("Mq", InternalFrictionAngle_2)
    
    Set sp = Nothing
End Function

Public Function getTables5_5_Mc(InternalFrictionAngle_2 As Double)
    Dim sp As New C_SP22_13330_2011
    
    getTables5_5_Mc = sp.Tables.t5_5("Mc", InternalFrictionAngle_2)

    Set sp = Nothing
End Function

Public Function getTables5_8(ksi As Double, L_B As Double, typeOfFoundation As String)
    Dim sp As New C_SP22_13330_2011

    getTables5_8 = sp.Tables.t5_8(ksi, L_B, typeOfFoundation)

    Set sp = Nothing
End Function

Public Function getTables5_9(L_B As Double, typeOfFoundation As String)
    Dim sp As New C_SP22_13330_2011

    getTables5_9 = sp.Tables.t5_9(L_B, typeOfFoundation)

    Set sp = Nothing
End Function

Public Function getTables5_12y(InternalFrictionAngle_1 as Double, delta as Double)
    dim sp as new C_SP22_13330_2011

    getTables5_12y = sp.Tables.t5_12("Ny", InternalFrictionAngle_1, delta)

    set sp = Nothing
End Function

Public Function getTables5_12q(InternalFrictionAngle_1 as Double, delta as Double)
    dim sp as new C_SP22_13330_2011

    getTables5_12q = sp.Tables.t5_12("Nq", InternalFrictionAngle_1, delta)

    set sp = Nothing
End Function
Public Function getTables5_12c(InternalFrictionAngle_1 as Double, delta as Double)
    dim sp as new C_SP22_13330_2011

    getTables5_12c = sp.Tables.t5_12("Nc", InternalFrictionAngle_1, delta)

    set sp = Nothing
End Function