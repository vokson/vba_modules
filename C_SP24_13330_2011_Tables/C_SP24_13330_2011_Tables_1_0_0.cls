VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_SP24_13330_2011_Tables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private math As C_Math

Private Const SIZE_TYPE_SAND = "ÏÅÑÎÊ"

Private Const SIZE_SAND_SUBTYPE_GRAVEL = "ÃÐÀÂÅËÈÑÒÛÉ"
Private Const SIZE_SAND_SUBTYPE_COARSE = "ÊÐÓÏÍÛÉ"
Private Const SIZE_SAND_SUBTYPE_MIDDLE = "ÑÐÅÄÍÅÉ ÊÐÓÏÍÎÑÒÈ"
Private Const SIZE_SAND_SUBTYPE_SMALL = "ÌÅËÊÈÉ"
Private Const SIZE_SAND_SUBTYPE_FINE = "ÏÛËÅÂÀÒÛÉ"

Private Const DENSITY_TYPE_SAND_HIGH = "ÏËÎÒÍÛÉ"
Private Const DENSITY_TYPE_SAND_MIDDLE = "ÑÐÅÄÍÅÉ ÏËÎÒÍÎÑÒÈ"
Private Const DENSITY_TYPE_SAND_LOW = "ÐÛÕËÛÉ"

Private Sub Class_Initialize()
    Set math = New C_Math
End Sub

Private Sub Class_Terminate()
    Set math = Nothing
End Sub

' ÑÏ 24.13330.2011 (èçì.1) - Òàáëèöà 7.2 Ïðèìå÷àíèå 2
Public Function t7_2_Note_2(originalSoilElevation As Double, finalSoilElevation As Double) As Double
    Dim result As Double
    result = originalSoilElevation
    
    If Abs(originalSoilElevation - finalSoilElevation) > 3 Then
        If finalSoilElevation > originalSoilElevation Then
            result = finalSoilElevation - 3
        Else
            result = finalSoilElevation + 3
        End If
    End If

    t7_2_Note_2 = result
End Function

' ÑÏ 24.13330.2011 (èçì.1) - Òàáëèöà 7.2 Ïðèìå÷àíèå 4
' soil - C_Soil v.2.0.0
Public Function t7_2_Note_4( _
    soil As C_Soil, _
    Optional isDensityAsPerCPTWithoutHoles As Boolean = False _
) As Double

    Dim result As Double
    result = 1#

    If soil.TypeBySize = SIZE_TYPE_SAND And soil.TypeByDensity = DENSITY_TYPE_SAND_HIGH Then
    
        If isDensityAsPerCPTWithoutHoles = True Then
            Select Case soil.SubtypeBySize
                Case SIZE_SAND_SUBTYPE_COARSE, SIZE_SAND_SUBTYPE_MIDDLE:
                    result = 2.0
                Case SIZE_SAND_SUBTYPE_SMALL, SIZE_SAND_SUBTYPE_FINE:
                    result = 2.3
            End Select
        Else
            Select Case soil.SubtypeBySize
                Case SIZE_SAND_SUBTYPE_COARSE, SIZE_SAND_SUBTYPE_MIDDLE:
                    result = 1.6
                Case SIZE_SAND_SUBTYPE_SMALL, SIZE_SAND_SUBTYPE_FINE:
                    result = 1.75
            End Select
            
            If result > 20000 Then result = 20000
        End If
        
    End If

    t7_2_Note_4 = result
    
End Function

' ÑÏ 24.13330.2011 (èçì.1) - Òàáëèöà 7.2 äëÿ ïåñêîâ
' soil - C_Soil v.2.0.0
Function t7_2_forSand(depth As Double, soil As C_Soil) As Double
    
    If soil.TypeBySize <> SIZE_TYPE_SAND Then _
        MsgBox "SP24.13330.2011.Tables.t7_2_forSand is only for sand", vbOK, "ERROR"

    Dim values As New Dictionary
    values.Item("depth") = Array(3, 4, 5, 7, 10, 15, 20, 25, 30, 35, 40)
    values.Item(SIZE_SAND_SUBTYPE_GRAVEL) = Array(7500, 8300, 8800, 9700, 10500, 11700, 12600, 13400, 14200, 15000, 15800)
    values.Item(SIZE_SAND_SUBTYPE_COARSE) = Array(6600, 6800, 7000, 7300, 7700, 8200, 8500, 9000, 9500, 10000, 10500)
    values.Item(SIZE_SAND_SUBTYPE_MIDDLE) = Array(3100, 3200, 3400, 3700, 4000, 4400, 4800, 5200, 5600, 6000, 6400)
    values.Item(SIZE_SAND_SUBTYPE_SMALL) = Array(2000, 2100, 2200, 2400, 2600, 2900, 3200, 3500, 3800, 4100, 4400)
    values.Item(SIZE_SAND_SUBTYPE_FINE) = Array(1100, 1250, 1300, 1400, 1500, 1650, 1800, 1950, 2100, 2250, 2400)

    t7_2_forSand = math.interpolateOneDimensionalArray(depth, values.Item("depth"), values.Item(soil.SubtypeBySize))
End Function


