Attribute VB_Name = "M_SP22_13330_2011_Test"
Option Explicit

Public Function getTables5_10( typeOfSoil As String, IL As Double) As Double

    Dim sp As New C_SP22_13330_2011
    Dim soil As New C_Soil

    soil.ClassOfSoil = "ƒ»—œ≈–—Õ€…"
    soil.TypeBySize = typeOfSoil
    soil.LiquidityIndex = IL

    getTables5_10 = sp.Tables.t5_10(soil)

    Set sp = Nothing
    Set soil = Nothing

    ' Hello Alex
End Function