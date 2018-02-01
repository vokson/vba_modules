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

Private Const SIZE_TYPE_MACROFRAGMENTAL = "КРУПНООБЛОМОЧНЫЙ"
Private Const SIZE_TYPE_SAND = "ПЕСОК"
Private Const SIZE_TYPE_CLAY_SANDY = "СУПЕСЬ"
Private Const SIZE_TYPE_CLAY_LOAM = "СУГЛИНОК"
Private Const SIZE_TYPE_CLAY = "ГЛИНА"

Private Const SIZE_SAND_SUBTYPE_GRAVEL = "ГРАВЕЛИСТЫЙ"
Private Const SIZE_SAND_SUBTYPE_COARSE = "КРУПНЫЙ"
Private Const SIZE_SAND_SUBTYPE_MIDDLE = "СРЕДНЕЙ КРУПНОСТИ"
Private Const SIZE_SAND_SUBTYPE_SMALL = "МЕЛКИЙ"
Private Const SIZE_SAND_SUBTYPE_FINE = "ПЫЛЕВАТЫЙ"

Private Const DENSITY_TYPE_SAND_HIGH = "ПЛОТНЫЙ"
Private Const DENSITY_TYPE_SAND_MIDDLE = "СРЕДНЕЙ ПЛОТНОСТИ"
Private Const DENSITY_TYPE_SAND_LOW = "РЫХЛЫЙ"

Private Sub Class_Initialize()
    Set math = New C_Math
End Sub

Private Sub Class_Terminate()
    Set math = Nothing
End Sub

' СП 24.13330.2011 (изм.1) - Таблица 7.2 Примечание 2
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

' СП 24.13330.2011 (изм.1) - Таблица 7.2 Примечание 4
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

' СП 24.13330.2011 (изм.1) - Таблица 7.2 для песков
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

' СП 24.13330.2011 (изм.1) - Таблица 7.2 для глинистых грунтов
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

' СП 24.13330.2011 (изм.1) - Таблица 7.2
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

' СП 24.13330.2011 (изм.1) - Таблица 7.3 для глинистых грунтов
Function t7_3_forClay(depth As Double, IL As Double) As Double
    
    Dim i As Integer, minIL As Double, maxIL As Double, min As Double, max As Double
    Dim k As Integer, minDepthKey As Integer, maxDepthKey As Integer, minDepth As Double, maxDepth As Double
    
    Dim values As New Dictionary
    
    Const ARRAY_LENGTH = 14
    values.Item("depth") = Array(1, 2, 3, 4, 5, 6, 8, 10, 15, 20, 25, 30, 35, 40)
    values.Item("IL") = Array(0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1#)
    values.Item(0) = Array(35, 42, 48, 53, 56, 58, 62, 65, 72, 79, 86, 93, 100, 107)
    values.Item(1) = Array(23, 30, 35, 38, 40, 42, 44, 46, 51, 56, 61, 66, 70, 74)
    values.Item(2) = Array(15, 21, 25, 27, 29, 31, 33, 34, 38, 41, 44, 47, 50, 53)
    values.Item(3) = Array(12, 17, 20, 22, 24, 25, 26, 27, 28, 30, 32, 34, 36, 38)
    values.Item(4) = Array(8, 12, 14, 16, 17, 18, 19, 19, 20, 20, 20, 21, 22, 23)
    values.Item(5) = Array(4, 7, 8, 9, 10, 10, 10, 10, 11, 12, 12, 12, 13, 14)
    values.Item(6) = Array(4, 5, 7, 8, 8, 8, 8, 8, 8, 8, 8, 9, 9, 9)
    values.Item(7) = Array(3, 4, 6, 7, 7, 7, 7, 7, 7, 7, 7, 8, 8, 8)
    values.Item(8) = Array(2, 4, 5, 5, 6, 6, 6, 6, 6, 6, 6, 7, 7, 7)
    
    If (depth < 1#) Then depth = 1#
    If (IL < 0.2) Then IL = 0.2
        
    minIL = 8: maxIL = 8
    For i = 1 To 8
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
        t7_3_forClay = min + (IL - values.Item("IL")(minIL)) / _
            (values.Item("IL")(maxIL) - values.Item("IL")(minIL)) * (max - min)
    Else
        t7_3_forClay = min
    End If
    
End Function

' СП 24.13330.2011 (изм.1) - Таблица 7.3 для песков
Function t7_3_forSand(depth As Double, SubtypeBySize As String) As Double
    
    t7_3_forSand = 0

    Select Case SubtypeBySize
        Case SIZE_SAND_SUBTYPE_COARSE, SIZE_SAND_SUBTYPE_MIDDLE:
            t7_3_forSand = t7_3_forClay(depth, 0.2)
            
        Case SIZE_SAND_SUBTYPE_SMALL:
            t7_3_forSand = t7_3_forClay(depth, 0.3)
            
        Case SIZE_SAND_SUBTYPE_FINE:
            t7_3_forSand = t7_3_forClay(depth, 0.4)
    End Select
                
End Function

' СП 24.13330.2011 (изм.1) - Таблица 7.3 Прим. 3
Function t7_3_Note_3(TypeBySize As String, TypeByDensity As String) As Double
    t7_3_Note_3 = 1#
    If TypeBySize = SIZE_TYPE_SAND And TypeByDensity = DENSITY_TYPE_SAND_HIGH Then t7_3_Note_3 = 1.3
End Function

' СП 24.13330.2011 (изм.1) - Таблица 7.3 Прим. 4
Function t7_3_Note_4(TypeBySize As String, e As Double) As Double
    t7_3_Note_4 = 1#
    
    If (TypeBySize = SIZE_TYPE_CLAY And e < 0.6) Or _
        ((TypeBySize = SIZE_TYPE_CLAY_SANDY Or TypeBySize = SIZE_TYPE_CLAY_LOAM) And e < 0.5) _
    Then t7_3_Note_4 = 1.15
End Function

' СП 24.13330.2011 (изм.1) - Таблица 7.3
' soil - C_Soil v.2.0.0
Function t7_3(depth As Double, soil As C_Soil) As Double
    
    If soil.TypeBySize = SIZE_TYPE_SAND And soil.TypeByDensity <> DENSITY_TYPE_SAND_LOW Then
        t7_3 = t7_3_forSand(depth, soil.SubtypeBySize)
    End If
    
    If soil.TypeBySize = SIZE_TYPE_CLAY Or _
       soil.TypeBySize = SIZE_TYPE_CLAY_LOAM Or _
       soil.TypeBySize = SIZE_TYPE_CLAY_SANDY _
    Then
        t7_3 = t7_3_forClay(depth, soil.LiquidityIndex)
    End If
    
    t7_3 = t7_3 * t7_3_Note_3(soil.TypeBySize, soil.TypeByDensity)
    t7_3 = t7_3 * t7_3_Note_4(soil.TypeBySize, soil.VoidRatio)
End Function

' СП 24.13330.2011 (изм.1) - Таблица 7.7 - Альфа 1
Function t7_7_a1(Fi As Double) As Double
    
    Dim values As New Dictionary
    
    values.Item("key") = Array(23#, 25#, 27#, 29#, 31#, 33#, 35#, 37#, 39#)
    values.Item("value") = Array(9.5, 12.6, 17.3, 24.4, 34.6, 48.6, 71.3, 108#, 163#)

    t7_7_a1 = math.interpolateOneDimensionalArray(Fi, values.Item("key"), values.Item("value"))
    
End Function

' СП 24.13330.2011 (изм.1) - Таблица 7.7 - Альфа 2
Function t7_7_a2(Fi As Double) As Double
    
    Dim values As New Dictionary
    
    values.Item("key") = Array(23#, 25#, 27#, 29#, 31#, 33#, 35#, 37#, 39#)
    values.Item("value") = Array(18.6, 24.8, 32.8, 45.5, 64#, 87.6, 127#, 185#, 260#)

    t7_7_a2 = math.interpolateOneDimensionalArray(Fi, values.Item("key"), values.Item("value"))
    
End Function

' СП 24.13330.2011 (изм.1) - Таблица 7.7 - Альфа 3
Function t7_7_a3(h As Double, d As Double, Fi As Double) As Double

    If h <= 0 Or d <= 0 Then
        MsgBox "SP 24.13330.2011 Table 7.7 - h and d must be > 0", vbOk, "ERROR"
        Exit Function
    End If
    
    Dim values As New Dictionary
    
    values.Item("keyArray1") = Array(4#, 5#, 7.5, 10#, 12.5, 15#, 17.5, 20#, 22.5, 25#)
    values.Item("keyArray2") = Array(23#, 25#, 27#, 29#, 31#, 33#, 35#, 37#, 39#)
    values.Item("valueArray") = Array( _
        Array(0.78, 0.79, 0.80, 0.82, 0.84, 0.85, 0.85, 0.85, 0.87), _
        Array(0.75, 0.76, 0.77, 0.79, 0.81, 0.82, 0.83, 0.84, 0.85), _
        Array(0.68, 0.70, 0.71, 0.74, 0.76, 0.78, 0.80, 0.82, 0.84), _
        Array(0.62, 0.65, 0.67, 0.70, 0.73, 0.75, 0.77, 0.79, 0.81), _
        Array(0.58, 0.61, 0.63, 0.67, 0.70, 0.73, 0.75, 0.78, 0.80), _
        Array(0.55, 0.58, 0.61, 0.65, 0.68, 0.71, 0.73, 0.76, 0.79), _
        Array(0.51, 0.55, 0.58, 0.62, 0.66, 0.69, 0.72, 0.75, 0.78), _
        Array(0.49, 0.53, 0.57, 0.61, 0.65, 0.68, 0.72, 0.75, 0.78), _
        Array(0.46, 0.51, 0.55, 0.60, 0.64, 0.67, 0.71, 0.74, 0.77), _
        Array(0.44, 0.49, 0.54, 0.59, 0.63, 0.67, 0.70, 0.74, 0.77) _
    )

    t7_7_a3 = math.interpolateTwoDimensionalArray( _
        h/d, Fi, values.Item("keyArray1"), values.Item("keyArray2"), values.Item("valueArray") _
    )
    
End Function

' СП 24.13330.2011 (изм.1) - Таблица 7.8
' Для h>20м при IL=0.5 и 0.6 взяты значения, как для IL=0.4
Function t7_8(h As Double, IL As Double) As Double

    If h <= 0 Then
        MsgBox "SP 24.13330.2011 Table 7.8 - h must be > 0", vbOk, "ERROR"
        Exit Function
    End If

    If h > 20# And IL > 0.4 Then
        MsgBox "SP 24.13330.2011 Table 7.8 - h and IL are out of table's values", vbOk, "ERROR"
        Exit Function
    End If
    
    Dim values As New Dictionary
    
    values.Item("keyArray1") = Array(3#, 5#, 7#, 10#, 12#, 15#, 18#, 20#, 30#, 40#)
    values.Item("keyArray2") = Array(0#, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6)
    values.Item("valueArray") = Array( _
        Array(850, 750, 650, 500, 400, 300, 250), _
        Array(1000, 850, 750, 650, 500, 400, 350), _
        Array(1150, 1000, 850, 750, 600, 500, 450), _
        Array(1350, 1200, 1050, 950, 800, 700, 600), _
        Array(1550, 1400, 1250, 1100, 950, 800, 700), _
        Array(1800, 1650, 1500, 1300, 1100, 1000, 800), _
        Array(2100, 1900, 1700, 1500, 1300, 1150, 950), _
        Array(2300, 2100, 1900, 1650, 1450, 1250, 1050), _
        Array(3300, 3000, 2600, 2300, 2000, 1250, 1050), _
        Array(4500, 4000, 3500, 3000, 2500, 1250, 1050) _
    )

    t7_8 = math.interpolateTwoDimensionalArray( _
        h, IL, values.Item("keyArray1"), values.Item("keyArray2"), values.Item("valueArray") _
    )
    
End Function

' СП 24.13330.2011 (изм.1) - Таблица 7.7 - Альфа 4
Function t7_7_a4(d As Double, Fi As Double) As Double

    If d <= 0 Then
        MsgBox "SP 24.13330.2011 Table 7.7 - d must be > 0", vbOk, "ERROR"
        Exit Function
    End If
    
    Dim values As New Dictionary
    
    values.Item("keyArray1") = Array(0.8, 4#)
    values.Item("keyArray2") = Array(23#, 25#, 27#, 29#, 31#, 33#, 35#, 37#, 39#)
    values.Item("valueArray") = Array( _
        Array(0.34, 0.31, 0.29, 0.27, 0.26, 0.25, 0.24, 0.23, 0.22), _
        Array(0.25, 0.24, 0.23, 0.22, 0.21, 0.20, 0.19, 0.18, 0.17) _
    )

    t7_7_a4 = math.interpolateTwoDimensionalArray( _
        d, Fi, values.Item("keyArray1"), values.Item("keyArray2"), values.Item("valueArray") _
    )
    
End Function

' СП 24.13330.2011 (изм.1) - Таблица 7.16 (коэффициент В1)
' qs - [кПа]
Function t7_16_B1(qs As Double, isDrivenPile As Boolean, Optional isCompressed As Boolean = True, _
    Optional isScrewPileInSandyWaterSaturatedSoil As Boolean = False) As Double
    
    If qs < 0 Then
        t7_16_B1 = 0
        Debug.Print "t7_16_B1: qs = " & qs & " is lower than zero."
        Exit Function
    End If
    
    Dim values As New Dictionary
    
    values.Item("qs") = Array(1000, 2500, 5000, 7500, 10000, 15000, 20000, 30000)
    values.Item("DRIVEN_PILE") = Array(0.9, 0.8, 0.65, 0.55, 0.45, 0.35, 0.3, 0.2)
    values.Item("SCREW_PILE_COMPRESSION") = Array(0.5, 0.45, 0.32, 0.26, 0.23, 0.23, 0.23, 0.23)
    values.Item("SCREW_PILE_TENSION") = Array(0.4, 0.38, 0.27, 0.22, 0.19, 0.19, 0.19, 0.19)
    
    Dim typeOfPile As String
    If isDrivenPile = True Then
        typeOfPile = "DRIVEN_PILE"
    Else
        If isCompressed = True Then
            typeOfPile = "SCREW_PILE_COMPRESSION"
        Else
            typeOfPile = "SCREW_PILE_TENSION"
        End If
    End If
    
    t7_16_B1 = math.interpolateOneDimensionalArray(qs, values.Item("qs"), values.Item(typeOfPile))
    
    If isScrewPileInSandyWaterSaturatedSoil = True And isDrivenPile = False Then _
        t7_16_B1 = t7_16_B1 / 2
    
End Function

' СП 24.13330.2011 (изм.1) - Таблица 7.16 (коэффициент В2i)
' fs - [КПа]
' typeOfZond = 1, 2 or 3
Function t7_16_B2i(fs As Double, typeOfZond As Integer, isSand As Boolean) As Double
    
    If fs < 0 Then
        t7_16_B2i = 0
        Debug.Print "t7_16_B2i: fs = " & fs & " is lower than zero."
        Exit Function
    End If
    
    Dim values As New Dictionary
    
    values.Item("fs") = Array(20, 40, 60, 80, 100, 120)
    values.Item("TYPE_I_SAND") = Array(2.4, 1.65, 1.2, 1#, 0.85, 0.75)
    values.Item("TYPE_I_CLAY") = Array(1.5, 1#, 0.75, 0.6, 0.5, 0.4)
    values.Item("TYPE_II_SAND") = Array(0.75, 0.6, 0.55, 0.5, 0.45, 0.4)
    values.Item("TYPE_II_CLAY") = Array(1#, 0.75, 0.6, 0.45, 0.4, 0.3)
    
    Dim typeOfZondAndSoil As String
    
    Select Case typeOfZond
        Case 1: typeOfZondAndSoil = "I"
        Case 2, 3: typeOfZondAndSoil = "II"
    End Select
    
    Select Case isSand
        Case True: typeOfZondAndSoil = "TYPE_" & typeOfZondAndSoil & "_SAND"
        Case False: typeOfZondAndSoil = "TYPE_" & typeOfZondAndSoil & "_CLAY"
    End Select
    
    t7_16_B2i = math.interpolateOneDimensionalArray(fs, values.Item("fs"), values.Item(typeOfZondAndSoil))
End Function

' СП 24.13330.2011 (изм.1) - Таблица Ж.1 (касательная сила морозного пучения)
' soil - C_Soil v.2.1.0
Function tG_1(depth As Double, soil As C_Soil) As Double

    Dim values As New Dictionary
    values.Item("depth") = Array(1.5, 2.5, 3.0)
    values.Item("TYPE_1") = Array(110#, 90#, 70#)
    values.Item("TYPE_2") = Array(90#, 70#, 55#)
    values.Item("TYPE_3") = Array(70#, 55#, 40#)

    tG_1 = 0#
    
    If soil.TypeBySize = SIZE_TYPE_SAND And _
        (soil.SubtypeBySize = SIZE_SAND_SUBTYPE_SMALL Or soil.SubtypeBySize = SIZE_SAND_SUBTYPE_FINE) Then

        If (soil.GranulationFactor > 5# And soil.DegreeOfSaturation > 0.95) Then
            tG_1 = math.interpolateOneDimensionalArray(depth, values.Item("depth"), values.Item("TYPE_1"))

        ElseIf (soil.GranulationFactor > 1# And _
            soil.DegreeOfSaturation > 0.8 And soil.DegreeOfSaturation <= 0.95) Then
                tG_1 = math.interpolateOneDimensionalArray(depth, values.Item("depth"), values.Item("TYPE_2"))

        ElseIf (soil.GranulationFactor > 1# And _
            soil.DegreeOfSaturation > 0.6 And soil.DegreeOfSaturation <= 0.8) Then
                tG_1 = math.interpolateOneDimensionalArray(depth, values.Item("depth"), values.Item("TYPE_3"))
        End If

    End If

    If soil.TypeBySize = SIZE_TYPE_CLAY Or soil.TypeBySize = SIZE_TYPE_CLAY_SANDY Or _
         soil.TypeBySize = SIZE_TYPE_CLAY_LOAM Then

        If soil.LiquidityIndex > 0.5 Then 
             tG_1 = math.interpolateOneDimensionalArray(depth, values.Item("depth"), values.Item("TYPE_1"))
        End If

        If soil.LiquidityIndex > 0.25 And  soil.LiquidityIndex <= 0.5  Then 
             tG_1 = math.interpolateOneDimensionalArray(depth, values.Item("depth"), values.Item("TYPE_2"))
        End If

        If soil.LiquidityIndex <= 0.25 Then 
             tG_1 = math.interpolateOneDimensionalArray(depth, values.Item("depth"), values.Item("TYPE_3"))
        End If

    End If

End Function
