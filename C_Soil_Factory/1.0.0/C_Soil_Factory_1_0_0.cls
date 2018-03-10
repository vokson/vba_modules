VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Soil_Factory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const STANDART_STEP_FOR_AVERAGE_PARAMETER = 0.1

Private math As C_Math

Private Sub Class_Initialize()
    Set math = New C_Math
End Sub

Private Sub Class_Terminate()
    Set math = Nothing
End Sub


Public Function getAverageParameterBtwDepth( _
    iterator As C_Soil_Factory_Iterator, _
    brh As C_Borehole, _
    soilBase As C_Soil_Database, _
    depth1 As Double, _
    depth2 As Double _
)

    If (depth1 = depth2) Then
        Exit Function
    End If

    If depth1 < 0 Or depth2 < 0 Then
        MsgBox "C_Soil_Factory - √лубина д.б. >= 0"
        Exit Function
    End If
    
    Dim min As Double
    Dim max As Double
    Dim count As Integer
    Dim step As Double
    Dim depth As Double
    Dim nameOfSoil As String
    Dim soil As C_Soil
    Dim additive As Double
    Dim result As Double

    min = math.min(depth1, depth2)
    max = math.max(depth1, depth2)

    count = math.max(1, Fix((max - min) / STANDART_STEP_FOR_AVERAGE_PARAMETER))
    
    step = (max - min) / count
    
    Dim i As Integer
    result = 0#
    For i = 0 To count - 1
    
        depth = min + step * (i + 0.5)
        
        nameOfSoil = brh.getSoilNameAtDepth(depth)
        Set soil = soilBase.getSoil(nameOfSoil)

        ' additive = CDbl(soil.getPropertyByName(nameOfParameter))

        ' if (nameOfParameter = "SOIL_DENSITY" Or _
        '    nameOfParameter = "SOIL_DENSITY_1" Or _
        '    nameOfParameter = "SOIL_DENSITY_2") And brh.isWaterAtDepth(depth) _
        ' Then
        '     additive = additive - soil.waterDensity
        ' End If

        ' result = result + additive

        result = result + iterator.use(soil, depth)
        
    Next i
    
    getAverageParameterBtwDepth = result / count

End Function