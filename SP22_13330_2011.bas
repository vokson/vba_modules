Attribute VB_Name = "SP22_13330_2011"
Option Explicit

Function getTable5_10(typeOfSoil As String, Optional IL As Double = 0) As Double

    Select Case typeOfSoil
        Case SOIL_TYPE_MACROFRAGMENTAL:
            getTable5_10 = 0.27
        Case SOIL_TYPE_SAND, SOIL_TYPE_CLAY_SANDY:
            getTable5_10 = (0.3 + 0.35) / 2
        Case SOIL_TYPE_SAND, SOIL_TYPE_CLAY_LOAM:
            getTable5_10 = (0.37 + 0.35) / 2
        Case SOIL_TYPE_CLAY:
            If IL <= 0 Then getTable5_10 = (0.2 + 0.3) / 2
            If IL > 0 And IL <= 0.25 Then getTable5_10 = 0.3 + (0.38 - 0.3) * IL / 0.25
            If IL > 0.25 And IL <= 1# Then getTable5_10 = 0.38 + (0.45 - 0.38) * (IL - 0.25) / 0.75
            If IL > 1# Then getTable5_10 = 0.45
    End Select
    
End Function


