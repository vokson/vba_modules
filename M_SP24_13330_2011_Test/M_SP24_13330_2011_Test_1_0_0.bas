Attribute VB_Name = "M_SP24_13330_2011_Test"
Option Explicit

Public Function getTables7_2_Note_2(originalSoilElevation As Double, finalSoilElevation As Double) As Double
    Dim sp As New C_SP24_13330_2011
    getTables7_2_Note_2 = sp.Tables.t7_2_Note_2(originalSoilElevation, finalSoilElevation)
    Set sp = Nothing
End Function


Public Function getTables7_2_Note_4( _
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

    getTables7_2_Note_4 = sp.Tables.t7_2_Note_4(soil, isDensityAsPerCPTWithoutHoles)

    Set sp = Nothing
    Set soil = Nothing

End Function

Function getTables7_2_forSand(depth As Double, subtypeOfSoil As String) As Double

    Dim sp As New C_SP24_13330_2011
    Dim soil As New C_Soil

    soil.ClassOfSoil = "ƒ»—œ≈–—Õ€…"
    soil.TypeBySize = "œ≈—Œ "
    soil.SubtypeBySize = subtypeOfSoil

    getTables7_2_forSand = sp.Tables.t7_2_forSand(depth, soil)

    Set sp = Nothing
    Set soil = Nothing

End Function