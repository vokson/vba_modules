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

            If soil.LiquidityIndex > 0 And soil.LiquidityIndex <= 0.25 Then _
               result = 0.3 + (0.38 - 0.3) * soil.LiquidityIndex / 0.25

            If soil.LiquidityIndex > 0.25 And soil.LiquidityIndex <= 1# Then _
               result = 0.38 + (0.45 - 0.38) * (soil.LiquidityIndex - 0.25) / 0.75

            If soil.LiquidityIndex > 1# Then result = 0.45
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

