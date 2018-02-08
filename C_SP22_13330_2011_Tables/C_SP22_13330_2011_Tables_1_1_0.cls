VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_SP22_13330_2011_Tables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private math As C_Math

Private Const SIZE_SAND_SUBTYPE_GRAVEL = "ГРАВЕЛИСТЫЙ"
Private Const SIZE_SAND_SUBTYPE_COARSE = "КРУПНЫЙ"
Private Const SIZE_SAND_SUBTYPE_MIDDLE = "СРЕДНЕЙ КРУПНОСТИ"
Private Const SIZE_SAND_SUBTYPE_SMALL = "МЕЛКИЙ"
Private Const SIZE_SAND_SUBTYPE_FINE = "ПЫЛЕВАТЫЙ"

Private Const DENSITY_TYPE_SAND_HIGH = "ПЛОТНЫЙ"
Private Const DENSITY_TYPE_SAND_MIDDLE = "СРЕДНЕЙ ПЛОТНОСТИ"
Private Const DENSITY_TYPE_SAND_LOW = "РЫХЛЫЙ"

Private Const SATURATION_TYPE_1 = "МАЛОЙ СТЕПЕНИ ВОДОНАСЫЩЕНИЯ"
Private Const SATURATION_TYPE_2 = "СРЕДНЕЙ СТЕПЕНИ ВОДОНАСЫЩЕНИЯ"
Private Const SATURATION_TYPE_3 = "ВОДОНАСЫЩЕННЫЙ"

Private Sub Class_Initialize()
    Set math = New C_Math
End Sub

Private Sub Class_Terminate()
    Set math = Nothing
End Sub


' Таблица 5.10
' C_Soil v.2.0.0
Function t5_10(soil As C_Soil) As Double

    Dim result As Double

    Select Case soil.TypeBySize
        Case "КРУПНООБЛОМОЧНЫЙ":
            result = 0.27

        Case "ПЕСОК", "СУПЕСЬ":
            result = (0.3 + 0.35) / 2

        Case "СУГЛИНОК":
            result = (0.37 + 0.35) / 2

        Case "ГЛИНА":
            If soil.LiquidityIndex <= 0 Then result = (0.2 + 0.3) / 2

            If soil.LiquidityIndex > 0 And soil.LiquidityIndex <= 0.25 Then result = (0.3 + 0.38) / 2

            If soil.LiquidityIndex > 0.25 And soil.LiquidityIndex <= 1 Then result = (0.38 + 0.45) / 2

            If soil.LiquidityIndex > 1 Then result = 0.45
    End Select

    t5_10 = result

End Function


' Таблица 5.4
' C_Soil v.2.1.0
Function t5_4_Yc1(soil As C_Soil) As Double

    Dim result As Double

    Select Case soil.TypeBySize
        Case "КРУПНООБЛОМОЧНЫЙ":
            result = 1.4

        Case "ПЕСОК":
            result = 1.4

            If soil.SubtypeBySize = SIZE_SAND_SUBTYPE_SMALL Then result = 1.3   ' Пески мелкие
            
            If soil.SubtypeBySize = SIZE_SAND_SUBTYPE_FINE Then
                result = 1.1                                                    ' Пески пылеватые насыщенные водой
                If soil.TypeByDegreeOfSaturation = SATURATION_TYPE_1 Or soil.TypeByDegreeOfSaturation = SATURATION_TYPE_2 Then result = 1.25 ' Пески пылеватые маловлажные и влажные
            End If

            If soil.TypeByDensity = DENSITY_TYPE_SAND_LOW Then result = 1 ' Пески рыхлые
        
        Case "ГЛИНА", "СУПЕСЬ", "СУГЛИНОК":
            result = 1.1 ' При IL > 0.5

            If soil.LiquidityIndex <= 0.25 Then result = 1.25 ' При IL <= 0.25

            If soil.LiquidityIndex > 0.25 And soil.LiquidityIndex <= 0.5 Then result = 1.2 ' При 0.25 < IL <= 0.5
            
    End Select

    t5_4_Yc1 = result

End Function

' Таблица 5.4
' C_Soil v.2.1.0
Function t5_4_Yc2(soil As C_Soil, L_H As Double, ModelFlexible As Boolean) As Double

    Dim result As Double

    If ModelFlexible Then
        t5_4_Yc2 = 1
        Exit Function
    End If

    Select Case soil.TypeBySize
        Case "КРУПНООБЛОМОЧНЫЙ":
            result = 1.2 + (1.4 - 1.2) * (4 - L_H) / 2.5
            If L_H >= 4 Then result = 1.2
            If L_H <= 1.5 Then result = 1.4

        Case "ПЕСОК":
            result = 1.2 + (1.4 - 1.2) * (4 - L_H) / 2.5
            If L_H >= 4 Then result = 1.2
            If L_H <= 1.5 Then result = 1.4

            If soil.SubtypeBySize = SIZE_SAND_SUBTYPE_SMALL Then ' Пески мелкие
                result = 1.1 + (1.3 - 1.1) * (4 - L_H) / 2.5
                If L_H >= 4 Then result = 1.1
                If L_H <= 1.5 Then result = 1.3
            End If
            
            If soil.SubtypeBySize = SIZE_SAND_SUBTYPE_FINE Then ' Пески пылеватые
              result = 1 + (1.2 - 1) * (4 - L_H) / 2.5
                If L_H >= 4 Then result = 1
                If L_H <= 1.5 Then result = 1.2
            End If

            If soil.TypeByDensity = DENSITY_TYPE_SAND_LOW Then result = 1 ' Пески рыхлые
        
        Case "ГЛИНА", "СУПЕСЬ", "СУГЛИНОК":
            result = 1 ' При IL > 0.5

            If soil.LiquidityIndex <= 0.5 Then ' При IL <= 0.5
                result = 1 + (1.1 - 1) * (4 - L_H) / 2.5
                If L_H >= 4 Then result = 1
                If L_H <= 1.5 Then result = 1.1
            End If
           
    End Select
    
    t5_4_Yc2 = result

End Function

' Таблица 5.5
Function t5_5(k As String, InternalFrictionAngle_2 As Double) As Double
Dim kFi2 As Variant, resA As Double, resB As Double

    Select Case k
        Case "My":
            kFi2 = Array(0, 0.01, 0.03, 0.04, 0.06, 0.08, 0.1, 0.12, 0.14, 0.16, 0.18, 0.21, 0.23, 0.26, 0.29, 0.32, 0.36, 0.39, 0.43, 0.47, 0.51, 0.56, 0.61, _
                0.66, 0.72, 0.78, 0.84, 0.91, 0.98, 1.06, 1.15, 1.24, 1.34, 1.44, 1.55, 1.68, 1.81, 1.95, 2.11, 2.28, 2.46, 2.66, 2.88, 3.12, 3.38, 3.66, 3.66)
        Case "Mq":
            kFi2 = Array(1, 1.06, 1.12, 1.18, 1.25, 1.32, 1.39, 1.47, 1.55, 1.64, 1.73, 1.83, 1.94, 2.05, 2.17, 2.3, 2.43, 2.57, 2.73, 2.89, 3.06, 3.24, 3.44, _
                3.65, 3.87, 4.11, 4.37, 4.64, 4.93, 5.25, 5.59, 5.95, 6.34, 6.76, 7.22, 7.71, 8.24, 8.81, 9.44, 10.11, 10.85, 11.64, 12.51, 13.46, 14.5, 15.64, 15.64)
        Case "Mc":
            kFi2 = Array(3.14, 3.23, 3.32, 3.41, 3.51, 3.61, 3.71, 3.82, 3.93, 4.05, 4.17, 4.29, 4.42, 4.55, 4.69, 4.84, 4.99, 5.15, 5.31, 5.48, 5.66, 5.84, 6.04, _
                6.24, 6.45, 6.67, 6.9, 7.14, 7.4, 7.67, 7.95, 8.24, 8.55, 8.88, 9.22, 9.58, 9.97, 10.37, 10.8, 11.25, 11.73, 12.24, 12.79, 13.37, 13.98, 14.64, 14.64)
        End Select

    resA = CDbl(kFi2(Int(InternalFrictionAngle_2)))
    resB = CDbl(kFi2(Int(InternalFrictionAngle_2) + 1))
        t5_5 = resA + (resB - resA) * (InternalFrictionAngle_2 - Int(InternalFrictionAngle_2)) / 1

End Function

' Таблица 5.8
Function t5_8(ksi As Double, L_B As Double, typeOfFoundation As String)
Dim kCircle As Variant, kLine As Variant, resA As Double, resB As Double, i As Integer

    Select Case typeOfFoundation
        Case "CIRCLE":
            kCircle = Array(1, 0.949, 0.756, 0.547, 0.39, 0.285, 0.214, 0.165, 0.13, 0.106, 0.087, 0.073, 0.062, 0.053, 0.046, 0.04, 0.036, 0.031, 0.028, _
                        0.024, 0.022, 0.021, 0.019, 0.017, 0.016, 0.015, 0.014, 0.013, 0.012, 0.011, 0.01, 0.01)
            
            resA = CDbl(kCircle(Int(ksi / 0.4)))
            resB = CDbl(kCircle(Int(ksi / 0.4) + 1))
            
            t5_8 = resA + (resB - resA) * (ksi - (Int(ksi / 0.4)) * 0.4) / 0.4
        
        Case "LINE":
            kLine = Array(1, 0.977, 0.881, 0.755, 0.642, 0.55, 0.477, 0.42, 0.374, 0.337, 0.306, 0.28, 0.258, 0.239, 0.223, 0.208, 0.196, 0.185, 0.175, 0.166, _
                        0.158, 0.15, 0.143, 0.137, 0.132, 0.126, 0.122, 0.117, 0.113, 0.109, 0.106, 0.106)
            
            resA = CDbl(kLine(Int(ksi / 0.4)))
            resB = CDbl(kLine(Int(ksi / 0.4) + 1))
            
            t5_8 = resA + (resB - resA) * (ksi - (Int(ksi / 0.4)) * 0.4) / 0.4
        
        Case "RECTANGLE":
            Dim values As New Dictionary
    
            values.Item("keyArray1") = Array(0, 0.4, 0.8, 1.2, 1.6, 2, 2.4, 2.8, 3.2, 3.6, 4, 4.4, 4.8, 5.2, 5.6, 6, 6.4, 6.8, 7.2, 7.6, 8, 8.4, 8.8, 9.2, 9.6, _
                                            10, 10.4, 10.8, 11.2, 11.6, 12)
            values.Item("keyArray2") = Array(1, 1.4, 1.8, 2.4, 3.2, 5)
            values.Item("valueArray") = Array( _
                Array(1, 1, 1, 1, 1, 1), Array(0.96, 0.972, 0.975, 0.976, 0.977, 0.977), _
                Array(0.8, 0.848, 0.866, 0.876, 0.879, 0.881), Array(0.606, 0.682, 0.717, 0.739, 0.749, 0.754), _
                Array(0.449, 0.532, 0.578, 0.612, 0.629, 0.639), Array(0.336, 0.414, 0.463, 0.505, 0.53, 0.545), _
                Array(0.257, 0.325, 0.374, 0.419, 0.449, 0.47), Array(0.201, 0.26, 0.304, 0.349, 0.383, 0.41), _
                Array(0.16, 0.21, 0.251, 0.294, 0.329, 0.36), Array(0.131, 0.173, 0.209, 0.25, 0.285, 0.319), _
                Array(0.108, 0.145, 0.176, 0.214, 0.248, 0.285), Array(0.091, 0.123, 0.15, 0.185, 0.218, 0.255), _
                Array(0.077, 0.105, 0.13, 0.161, 0.192, 0.23), Array(0.067, 0.091, 0.113, 0.141, 0.17, 0.208), _
                Array(0.058, 0.079, 0.099, 0.124, 0.152, 0.189), _
                Array(0.051, 0.07, 0.087, 0.11, 0.136, 0.173), _
                Array(0.045, 0.062, 0.077, 0.099, 0.122, 0.158), _
                Array(0.04, 0.055, 0.069, 0.088, 0.11, 0.145), _
                Array(0.036, 0.049, 0.062, 0.08, 0.1, 0.133), _
                Array(0.032, 0.044, 0.056, 0.072, 0.091, 0.123), _
                Array(0.029, 0.04, 0.051, 0.066, 0.084, 0.113), _
                Array(0.026, 0.037, 0.046, 0.06, 0.077, 0.105), _
                Array(0.024, 0.033, 0.042, 0.055, 0.071, 0.098), _
                Array(0.022, 0.031, 0.039, 0.051, 0.065, 0.091), _
                Array(0.02, 0.028, 0.036, 0.047, 0.06, 0.085), _
                Array(0.019, 0.026, 0.033, 0.043, 0.056, 0.079), _
                Array(0.017, 0.024, 0.031, 0.04, 0.052, 0.074), _
                Array(0.016, 0.022, 0.029, 0.037, 0.049, 0.069), _
                Array(0.015, 0.021, 0.027, 0.035, 0.045, 0.065), _
                Array(0.014, 0.02, 0.025, 0.033, 0.042, 0.061), _
                Array(0.013, 0.018, 0.023, 0.031, 0.04, 0.058))

            t5_8 = math.interpolateTwoDimensionalArray( _
                    ksi, L_B, values.Item("keyArray1"), values.Item("keyArray2"), values.Item("valueArray"))

    End Select
End Function

' Таблица 5.9
Function t5_9(L_B As Double, typeOfFoundation As String) As Double

    Dim values As New Dictionary
    values.Item("KeyArray1") = Array(1, 1.2, 1.5, 2, 3, 5, 10)

    Select Case typeOfFoundation
        Case "CIRCLE":
            t5_9 = 0.75
            Exit Function
            
        Case "RECTANGLE_LONG":
            values.Item("valueArray") = Array(0.5, 0.57, 0.68, 0.82, 1.17, 1.42, 2)
            
        Case "RECTANGLE_SHORT":
            values.Item("valueArray") = Array(0.5, 0.43, 0.36, 0.28, 0.2, 0.12, 0.07)
    End Select

    t5_9 = math.interpolateOneDimensionalArray(L_B, values.Item("KeyArray1"), values.Item("valueArray"))
    
End Function

'Таблица 5.12
Function t5_12(N As String, InternalFrictionAngle_1 As Double, delta As Double) As Double

    Dim result As Double, deltaBorder As Double, NBorder As Double, FiBorder As Double, FiNext As Double, NOfFiNext As Double
    Dim values As New Dictionary
    
    values.Item("KeyArray1") = Array(0, 5, 10, 15, 20, 25, 30, 35, 40, 45) ' Угол внутреннего трения грунта Fi_1
    values.Item("KeyArray2") = Array(0, 5, 10, 15, 20, 25, 30, 35, 40, 45) ' Угол наклона к вертикали равнодействующей внешней нагрузки delta
    values.Item("KeyDelta") = Array(0, 4.9, 9.8, 14.5, 18.9, 22.9, 26.5, 29.8, 32.7, 35.2) ' предельный угол наклона нагрузки

    Select Case N
        Case "Ny":
            values.Item("KeyNy") = Array(0, 0.05, 0.12, 0.21, 0.36, 0.58, 0.45, 1.6, 2.79, 5.22) ' значения кэффициентов при предельных углах нагрузки

            values.Item("valueArray") = Array( _
                            Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                            Array(0.2, 0.05, 0, 0, 0, 0, 0, 0, 0, 0), _
                            Array(0.6, 0.42, 0.12, 0, 0, 0, 0, 0, 0, 0), _
                            Array(1.35, 1.02, 0.61, 0.21, 0, 0, 0, 0, 0, 0), _
                            Array(2.88, 2.18, 1.47, 0.82, 0.36, 0, 0, 0, 0, 0), _
                            Array(5.87, 4.5, 3.18, 2, 1.05, 0.58, 0, 0, 0, 0), _
                            Array(12.39, 9.43, 6.72, 4.44, 2.63, 1.29, 0.45, 0, 0, 0), _
                            Array(27.5, 20.58, 14.63, 9.79, 6.08, 3.38, 1.6, 0, 0, 0), _
                            Array(66.01, 48.3, 33.84, 22.56, 14.18, 8.26, 4.3, 2.79, 0, 0), _
                            Array(177.61, 126.09, 86.2, 56.5, 32.26, 20.73, 11.26, 5.45, 5.22, 0))

            NBorder = math.interpolateOneDimensionalArray(delta, values.Item("KeyDelta"), values.Item("KeyNy"))

        Case "Nq":
            values.Item("KeyNq") = Array(1, 0.26, 1.16, 2.06, 2.69, 3.6, 4.95, 7.04, 10.46, 16.42) ' значения кэффициентов при предельных углах нагрузки

            values.Item("valueArray") = Array( _
                            Array(1, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                            Array(1.57, 0.26, 0, 0, 0, 0, 0, 0, 0, 0), _
                            Array(2.47, 2.16, 1.16, 0, 0, 0, 0, 0, 0, 0), _
                            Array(3.94, 3.45, 2.84, 2.06, 0, 0, 0, 0, 0, 0), _
                            Array(6.4, 5.56, 4.64, 3.64, 2.69, 0, 0, 0, 0, 0), _
                            Array(10.66, 9.17, 7.65, 6.13, 4.58, 3.6, 0, 0, 0, 0), _
                            Array(18.4, 15.63, 12.94, 10.37, 7.96, 5.67, 4.95, 0, 0, 0), _
                            Array(33.3, 27.86, 22.77, 18.12, 13.94, 10.24, 7.04, 0, 0, 0), _
                            Array(64.19, 52.71, 42.37, 33.26, 25.39, 18.7, 13.11, 10.46, 0, 0), _
                            Array(134.87, 108.24, 85.16, 65.58, 49.26, 35.93, 25.24, 16.82, 16.42, 0))

            NBorder = math.interpolateOneDimensionalArray(delta, values.Item("KeyDelta"), values.Item("KeyNq"))

        Case "Nc":
            values.Item("KeyNc") = Array(5.14, 2.93, 3.38, 3.94, 4.65, 5.58, 6.85, 8.63, 11.27, 15.82) ' значения кэффициентов при предельных углах нагрузки

            values.Item("valueArray") = Array( _
                            Array(5.14, 0, 0, 0, 0, 0, 0, 0, 0, 0), _
                            Array(6.49,2.93, 0, 0, 0, 0, 0, 0, 0, 0), _
                            Array(8.34,6.57,3.38, 0, 0, 0, 0, 0, 0, 0), _
                            Array(10.98,9.13,6.88,3.94, 0, 0, 0, 0, 0, 0), _
                            Array(14.84,12.53,10.02,7.26,4.65, 0, 0, 0, 0, 0), _
                            Array(20.72,17.53,14.26,10.99,7.68,5.58, 0, 0, 0, 0), _
                            Array(30.14,25.34,20.68,16.23,12.05,8.09,6.85, 0, 0, 0), _
                            Array(46.12,38.36,31.09,24.45,18.48,13.19,8.63,0, 0, 0), _
                            Array(75.31,61.63,49.31,38.45,29.07,21.1,14.43,11.27,0, 0), _
                            Array(133.87,107.23,84.16,64.58,48.26,34.93,24.24,15.82,15.82, 0))

            NBorder = math.interpolateOneDimensionalArray(delta, values.Item("KeyDelta"), values.Item("KeyNc"))
       End Select
       
    deltaBorder = math.interpolateOneDimensionalArray(InternalFrictionAngle_1, values.Item("KeyArray1"), values.Item("KeyDelta"))
    FiBorder = math.interpolateOneDimensionalArray(delta, values.Item("KeyDelta"), values.Item("KeyArray1"))
    
    FiNext = (Int(InternalFrictionAngle_1 / 5) + 1) * 5
    
       ' Условие попадания в "срезанный" квадрат:
        If FiBorder < FiNext - 5 Then
            FiBorder = FiNext - 5
            NBorder = math.interpolateTwoDimensionalArray( _
                      FiBorder, delta, values.Item("KeyArray1"), values.Item("KeyArray2"), values.Item("valueArray"))

        End If
        
    NOfFiNext = math.interpolateTwoDimensionalArray( _
                      FiNext, delta, values.Item("KeyArray1"), values.Item("KeyArray2"), values.Item("valueArray"))


    If delta = deltaBorder Then
            result = NBorder
        Else
            If delta > deltaBorder Then
                    result = 0
                Else
                    result = NOfFiNext + (NBorder - NOfFiNext) * (InternalFrictionAngle_1 - FiNext) / (FiBorder - FiNext)
            End If
    End If

    t5_12 = result

End Function
