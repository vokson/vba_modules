Attribute VB_Name = "SP24_13330_2011"
Option Explicit

Function getTable7_2Note2(originalSoilElevation As Double, finalSoilElevation As Double) As Double

    getTable7_2Note2 = originalSoilElevation
    
    If Abs(originalSoilElevation - finalSoilElevation) > 3 Then
        If finalSoilElevation > originalSoilElevation Then
            getTable7_2Note2 = finalSoilElevation - 3
        Else
            getTable7_2Note2 = finalSoilElevation + 3
        End If
    End If
    
End Function

' R - â ÊÏà
Function getTable7_2Note4(R As Double, typeOfSoil As String, subtypeOfSoil As String, densityOfSoil As String, _
            Optional isDensityAsPerCPTWithoutHoles As Boolean = False) As Double

    getTable7_2Note4 = R

    If typeOfSoil = SOIL_TYPE_SAND And densityOfSoil = SAND_DENSITY_HIGH Then
    
        If isDensityAsPerCPTWithoutHoles = True Then
            Select Case subtypeOfSoil
                Case SAND_SUBTYPE_COARSE, SAND_SUBTYPE_MIDDLE: getTable7_2Note4 = 2 * R
                Case SAND_SUBTYPE_SMALL, SAND_SUBTYPE_FINE: getTable7_2Note4 = 2.3 * R
            End Select
        Else
            Select Case subtypeOfSoil
                Case SAND_SUBTYPE_COARSE, SAND_SUBTYPE_MIDDLE: getTable7_2Note4 = 1.6 * R
                Case SAND_SUBTYPE_SMALL, SAND_SUBTYPE_FINE: getTable7_2Note4 = 1.75 * R
            End Select
            
            If getTable7_2Note4 > 20000 Then getTable7_2Note4 = 20000
        End If
        
    End If
    
End Function

Function getTable7_2(depth As Double, typeOfSoil As String, subtypeOfSoil As String, _
    density As String, Optional IL As Double = 0, Optional IP As Double = 0, Optional e As Double = 0, _
    Optional isDensityAsPerCPTWithoutHoles As Boolean = False) As Double
    
    
    If typeOfSoil = SOIL_TYPE_SAND And density <> SAND_DENSITY_LOW Then
        getTable7_2 = getTable7_2ForSand(depth, subtypeOfSoil)
    End If
    
    If typeOfSoil = SOIL_TYPE_CLAY Or typeOfSoil = SOIL_TYPE_CLAY_LOAM Then
        getTable7_2 = getTable7_2ForClay(depth, IL)
    End If
    
     If typeOfSoil = SOIL_TYPE_CLAY_SANDY Then
        If IP <= 4 And e < 0.8 Then
            getTable7_2 = getTable7_2ForSand(depth, SAND_SUBTYPE_FINE)
        Else
            getTable7_2 = getTable7_2ForClay(depth, IL)
        End If
    End If
    
    getTable7_2 = getTable7_2Note4(getTable7_2, typeOfSoil, subtypeOfSoil, _
                    density, isDensityAsPerCPTWithoutHoles)
                    
End Function

Function getTable7_2ForClay(depth As Double, IL As Double) As Double
    
    Dim i As Integer, minIL As Double, maxIL As Double, min As Double, max As Double
    Dim k As Integer, minDepthKey As Integer, maxDepthKey As Integer, minDepth As Double, maxDepth As Double
    
    Dim values As Scripting.Dictionary
    Set values = New Scripting.Dictionary
    
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
        getTable7_2ForClay = min + (IL - values.Item("IL")(minIL)) / _
            (values.Item("IL")(maxIL) - values.Item("IL")(minIL)) * (max - min)
    Else
        getTable7_2ForClay = min
    End If
    
End Function

Function getTable7_2ForSand(depth As Double, sandType As String) As Double
    
    Dim i As Integer, minIL As Double, maxIL As Double, min As Double, max As Double
    Dim k As Integer, minDepthKey As Integer, maxDepthKey As Integer, minDepth As Double, maxDepth As Double
    
    Dim values As Scripting.Dictionary
    Set values = New Scripting.Dictionary
    
    Const ARRAY_LENGTH = 11
    values.Item("depth") = Array(3, 4, 5, 7, 10, 15, 20, 25, 30, 35, 40)
    values.Item(SAND_SUBTYPE_GRAVEL) = Array(7500, 8300, 8800, 9700, 10500, 11700, 12600, 13400, 14200, 15000, 15800)
    values.Item(SAND_SUBTYPE_COARSE) = Array(6600, 6800, 7000, 7300, 7700, 8200, 8500, 9000, 9500, 10000, 10500)
    values.Item(SAND_SUBTYPE_MIDDLE) = Array(3100, 3200, 3400, 3700, 4000, 4400, 4800, 5200, 5600, 6000, 6400)
    values.Item(SAND_SUBTYPE_SMALL) = Array(2000, 2100, 2200, 2400, 2600, 2900, 3200, 3500, 3800, 4100, 4400)
    values.Item(SAND_SUBTYPE_FINE) = Array(1100, 1250, 1300, 1400, 1500, 1650, 1800, 1950, 2100, 2250, 2400)
    
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
        getTable7_2ForSand = values.Item(sandType)(minDepthKey) + (depth - minDepth) / (maxDepth - minDepth) _
                * (values.Item(sandType)(maxDepthKey) - values.Item(sandType)(minDepthKey))
                
End Function

Function getTable7_4ForClay(depth As Double, IL As Double) As Double
    
    Dim i As Integer, minIL As Double, maxIL As Double, min As Double, max As Double
    Dim k As Integer, minDepthKey As Integer, maxDepthKey As Integer, minDepth As Double, maxDepth As Double
    
    Dim values As Scripting.Dictionary
    Set values = New Scripting.Dictionary
    
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
        getTable7_4ForClay = min + (IL - values.Item("IL")(minIL)) / _
            (values.Item("IL")(maxIL) - values.Item("IL")(minIL)) * (max - min)
    Else
        getTable7_4ForClay = min
    End If
    
End Function

Function getTable7_4ForSand(depth As Double, sandType As String) As Double
    
    Select Case sandType
        Case SAND_SUBTYPE_COARSE, SAND_SUBTYPE_MIDDLE:
            getTable7_4ForSand = getTable7_4ForClay(depth, 0.2)
            
        Case SAND_SUBTYPE_SMALL:
            getTable7_4ForSand = getTable7_4ForClay(depth, 0.3)
            
        Case SAND_SUBTYPE_FINE:
            getTable7_4ForSand = getTable7_4ForClay(depth, 0.4)
    End Select
                
End Function

Function getTable7_4(depth As Double, typeOfSoil As String, subtypeOfSoil As String, _
    density As String, Optional IL As Double = 0, Optional e As Double = 0) As Double
    
    
    If typeOfSoil = SOIL_TYPE_SAND And density <> SAND_DENSITY_LOW Then
        getTable7_4 = getTable7_4ForSand(depth, subtypeOfSoil)
    End If
    
    If typeOfSoil = SOIL_TYPE_CLAY Or typeOfSoil = SOIL_TYPE_CLAY_LOAM Or typeOfSoil = SOIL_TYPE_CLAY_SANDY Then
        getTable7_4 = getTable7_4ForClay(depth, IL)
    End If
    
    getTable7_4 = getTable7_4 * getTable7_4Note3(typeOfSoil, density)
    getTable7_4 = getTable7_4 * getTable7_4Note4(typeOfSoil, e)
                    
End Function

Function getTable7_4Note3(typeOfSoil As String, densityOfSoil As String) As Double

    getTable7_4Note3 = 1#
    If typeOfSoil = SOIL_TYPE_SAND And densityOfSoil = SAND_DENSITY_HIGH Then getTable7_4Note3 = 1.3
    
End Function

Function getTable7_4Note4(typeOfSoil As String, e As Double) As Double

    getTable7_4Note4 = 1#
    
    If (typeOfSoil = SOIL_TYPE_CLAY And e < 0.6) Or _
        ((typeOfSoil = SOIL_TYPE_CLAY_SANDY Or typeOfSoil = SOIL_TYPE_CLAY_LOAM) And e < 0.5) _
    Then getTable7_4Note4 = 1.15
    
End Function

Function getClause7_4_3(e As Double, nu As Double) As Double
    getClause7_4_3 = e / 2 / (1 + nu)
End Function

Function getFormula7_32(force As Double, G1 As Double, G2 As Double, nu1 As Double, nu2 As Double, _
                            EA As Double, d As Double, length As Double) As Double
    
    Dim betta As Double
    
    betta = getFormula7_33(G1, G2, nu1, nu2, EA, d, length)
    Debug.Print "betta = " & betta
    
    getFormula7_32 = betta * Abs(force) / G1 / length
   
End Function

Function getFormula7_33(G1 As Double, G2 As Double, nu1 As Double, nu2 As Double, _
                            EA As Double, d As Double, length As Double) As Double
   
   Dim knu As Double, knu1 As Double, lambda1 As Double, ksi As Double
   Dim alpha_dash As Double, betta_dash As Double, betta As Double
   
   knu = getFormula7_35((nu1 + nu2) / 2)
   knu1 = getFormula7_35(nu1)
   
   ksi = EA / G1 / length ^ 2
   lambda1 = getFormula7_34(ksi)
   
   alpha_dash = 0.17 * Log(knu1 * length / d)
   betta_dash = 0.17 * Log(knu * G1 * length / G2 / d)
   
   getFormula7_33 = betta_dash / lambda1 + (1 - (betta_dash / alpha_dash)) / ksi
   
'   Debug.Print "knu = " & knu
'   Debug.Print "knu1 = " & knu1
'   Debug.Print "ksi = " & ksi
'   Debug.Print "lambda1 = " & lambda1
'   Debug.Print "alpha_dash = " & alpha_dash
'   Debug.Print "betta_dash = " & betta_dash
   
End Function

Function getFormula7_34(ksi As Double) As Double
   getFormula7_34 = 2.12 * ksi ^ 0.75 / (1 + 2.12 * ksi ^ 0.75)
End Function

Function getFormula7_35(nu As Double) As Double
   getFormula7_35 = 2.82 - 3.78 * nu + 2.18 * nu ^ 2
End Function

' qs â ÊÏà
Function getTable7_16_B1(qs As Double, isDrivenPile As Boolean, Optional isCompressed As Boolean = True, _
    Optional isScrewPileInSandyWaterSaturatedSoil As Boolean = False) As Double
    
    If qs < 0 Then
        getTable7_16_B1 = 0
        Debug.Print "getTable7_16_B1: qs = " & qs & " is lower than zero."
        Exit Function
    End If
    
    Dim values As Scripting.Dictionary
    Set values = New Scripting.Dictionary
    
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
    
    getTable7_16_B1 = interpolateOneDimensionalArray(qs, values.Item("qs"), values.Item(typeOfPile))
    
    If isScrewPileInSandyWaterSaturatedSoil = True And isDrivenPile = False Then _
        getTable7_16_B1 = getTable7_16_B1 / 2
    
End Function

' fs â ÊÏà
' typeOfZond = 1, 2 or 3
Function getTable7_16_B2i(fs As Double, typeOfZond As Integer, isSand As Boolean) As Double
    
    If fs < 0 Then
        getTable7_16_B2i = 0
        Debug.Print "getTable7_16_B2i: fs = " & fs & " is lower than zero."
        Exit Function
    End If
    
    Dim values As Scripting.Dictionary
    Set values = New Scripting.Dictionary
    
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
    
    
    getTable7_16_B2i = interpolateOneDimensionalArray(fs, values.Item("fs"), values.Item(typeOfZondAndSoil))
    
End Function

Function getFormula7_26(B1 As Double, qs As Double) As Double
    getFormula7_26 = B1 * qs
End Function

Function getFormula7_28(sum_Bi_Fsi_Hi As Double, h As Double) As Double
    getFormula7_28 = sum_Bi_Fsi_Hi / h
End Function


