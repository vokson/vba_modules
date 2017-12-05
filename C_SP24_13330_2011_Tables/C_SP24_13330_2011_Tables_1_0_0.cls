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

Private Const SIZE_TYPE_MACROFRAGMENTAL = "ÊÐÓÏÍÎÎÁËÎÌÎ×ÍÛÉ"
Private Const SIZE_TYPE_SAND = "ÏÅÑÎÊ"
Private Const SIZE_TYPE_CLAY_SANDY = "ÑÓÏÅÑÜ"
Private Const SIZE_TYPE_CLAY_LOAM = "ÑÓÃËÈÍÎÊ"
Private Const SIZE_TYPE_CLAY = "ÃËÈÍÀ"

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
    isDensityAsPerCPTWithoutHoles As Boolean _
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
Function t7_2_forSand(depth As Double, SubtypeBySize AS String) As Double
    
    ' If soil.TypeBySize <> SIZE_TYPE_SAND Then _
        ' MsgBox "SP24.13330.2011.Tables.t7_2_forSand is only for sand", vbOK, "ERROR"

    Dim values As New Dictionary
    values.Item("depth") = Array(3, 4, 5, 7, 10, 15, 20, 25, 30, 35, 40)
    values.Item(SIZE_SAND_SUBTYPE_GRAVEL) = Array(7500, 8300, 8800, 9700, 10500, 11700, 12600, 13400, 14200, 15000, 15800)
    values.Item(SIZE_SAND_SUBTYPE_COARSE) = Array(6600, 6800, 7000, 7300, 7700, 8200, 8500, 9000, 9500, 10000, 10500)
    values.Item(SIZE_SAND_SUBTYPE_MIDDLE) = Array(3100, 3200, 3400, 3700, 4000, 4400, 4800, 5200, 5600, 6000, 6400)
    values.Item(SIZE_SAND_SUBTYPE_SMALL) = Array(2000, 2100, 2200, 2400, 2600, 2900, 3200, 3500, 3800, 4100, 4400)
    values.Item(SIZE_SAND_SUBTYPE_FINE) = Array(1100, 1250, 1300, 1400, 1500, 1650, 1800, 1950, 2100, 2250, 2400)

    t7_2_forSand = math.interpolateOneDimensionalArray(depth, values.Item("depth"), values.Item(SubtypeBySize))
End Function

' ÑÏ 24.13330.2011 (èçì.1) - Òàáëèöà 7.2 äëÿ ãëèíèñòûõ ãðóíòîâ
Function t7_2_forClay(depth As Double, IL As Double) As Double
    
    Dim i As Integer, minIL As Double, maxIL As Double, min As Double, max As Double
    Dim k As Integer, minDepthKey As Integer, maxDepthKey As Integer, minDepth As Double, maxDepth As Double
    
    Dim values As NEw Dictionary
    
    Const ARRAY_LENGTH = 11
    values.Item("depth") = Array(3, 4, 5, 7, 10, 15, 20, 25, 30, 35, 40)
    values.Item("IL") = Array(0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6)
    values.Item(0) = Array(7500, 8300, 8800, 9700, 10500, 11700, 12600, 13400, 14200, 15000, 15800)
    values.Item(1) = Array(4000, 5100, 6200, 6900, 7300, 7500, 8500, 9000, 9500, 10000, 10500)
    values.Item(2) = Array(3000, 3800, 4000, 4300, 5000, 5600, 6200, 6800, 7400, 8000, 8600)
    values.Item(3) = Array(2000, 2500, 2800, 3300, 3500, 4000, 4500, 5200, 5600, 6000, 6400)
    values.Item(4) = Array(1200, 1600, 2000, 2200, 2400, 2900, 3200, 3500, 3800, 4100, 4400)
    values.Item(5) = Array(1100, 1250, 1300, 1400, 1500, 1650, 1800, 1950, 2100, 2250, 2400)
    values.Item(6) = Array(600, 700, 800, 850, 900, 1000, 1100, 1200, 1300, 1400, 1500)
    
    If (IL < 0) Then IL = 0
        
    minIL = 6: maxIL = 6
    For i = 1 To 6
'            Debug.Print "i = " & i
        If (values.Item("IL")(i) >= IL) Then
            minIL = i - 1
            maxIL = i
            Exit For
        End If
    Next i
    
'        Debug.Print "minIL = " & minIL
'        Debug.Print "maxIL = " & maxIL
    
    
    For k = 1 To ARRAY_LENGTH
        If (values.Item("depth")(k) >= depth) Then
            minDepth = values.Item("depth")(k - 1)
            maxDepth = values.Item("depth")(k)
            minDepthKey = k - 1
            maxDepthKey = k
            Exit For
        End If
    Next k
    
'        Debug.Print "minDepthKey = " & minDepthKey
'        Debug.Print "maxDepthKey = " & maxDepthKey
'
'        Debug.Print "minDepth = " & minDepth
'        Debug.Print "maxDepth = " & maxDepth
'
    min = values.Item(minIL)(minDepthKey) + (depth - minDepth) / (maxDepth - minDepth) _
            * (values.Item(minIL)(maxDepthKey) - values.Item(minIL)(minDepthKey))
    max = values.Item(maxIL)(minDepthKey) + (depth - minDepth) / (maxDepth - minDepth) _
            * (values.Item(maxIL)(maxDepthKey) - values.Item(maxIL)(minDepthKey))
            
'        Debug.Print "min = " & min
'        Debug.Print "max = " & max
        
    If (minIL <> maxIL) Then
        t7_2_forClay = min + (IL - values.Item("IL")(minIL)) / _
            (values.Item("IL")(maxIL) - values.Item("IL")(minIL)) * (max - min)
    Else
        t7_2_forClay = min
    End If
    
End Function

' ÑÏ 24.13330.2011 (èçì.1) - Òàáëèöà 7.2
' soil - C_Soil v.2.0.0
Function t7_2(depth As Double, soil As C_Soil, Optional isDensityAsPerCPTWithoutHoles As Boolean = False) As Double
    
    
    If soil.TypeBySize = SIZE_TYPE_SAND And soil.TypeByDensity <> DENSITY_TYPE_SAND_LOW Then
        t7_2 = t7_2_forSand(depth, soil.SubtypeBySize)
    End If
    
    If soil.TypeBySize = SIZE_TYPE_CLAY Or soil.TypeBySize = SIZE_TYPE_CLAY_LOAM Then
        t7_2 = t7_2_forClay(depth, soil.LiquidityIndex)
    End If
    
     If soil.TypeBySize = SIZE_TYPE_CLAY_SANDY Then
        If soil.PlasticityIndex <= 4 And soil.VoidRatio < 0.8 Then
            t7_2 = t7_2_forSand(depth, SIZE_SAND_SUBTYPE_FINE)
        Else
            t7_2 = t7_2_forClay(depth, soil.LiquidityIndex)
        End If
    End If
    
    t7_2 = t7_2 * t7_2_Note_4(soil, isDensityAsPerCPTWithoutHoles)
                    
End Function


