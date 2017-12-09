VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_CPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private math As C_Math

Private pName As String
Private pTopElevation As Double

Private pLayerDepth() As Double
Private pLayerFrontResistance() As Double
Private pLayerSideFriction() As Double

Private countOfLayers As Integer

Private Sub Class_Initialize() 
    countOfLayers = 0
    Set math = New C_Math
End Sub 

Private Sub Class_Terminate() 
    Set math = Nothing
End Sub 

' Имя скважины
Property Let nameOfBorehole(word As String)
    pName = word
End Property

Property Get nameOfBorehole() As String
    nameOfBorehole = pName
End Property

' Отметка устья скважины
Property Let topElevation(elevation As Double)
     pTopElevation = elevation
End Property

Property Get topElevation() As Double
    topElevation = pTopElevation
End Property

' Добавление слоя
Public Sub addLayer(depth As Double, FR As Double, SF As Double)

    If depth <=0 Then
        MsgBox "C_CPT - Depth " & depth & " <= 0. Borehole " & nameOfBorehole, vbOK, "ERROR"
        Exit Sub
    End If

    ReDim Preserve pLayerDepth(countOfLayers)
    ReDim Preserve pLayerFrontResistance(countOfLayers)
    ReDim Preserve pLayerSideFriction(countOfLayers)
    countOfLayers = countOfLayers + 1

    pLayerDepth(UBound(pLayerDepth)) = depth
    pLayerFrontResistance(UBound(pLayerFrontResistance)) = FR
    pLayerSideFriction(UBound(pLayerSideFriction)) = SF

End Sub

Public Function DeepCopy() As C_CPT
    Dim other As New C_CPT

    other.nameOfBorehole = Me.nameOfBorehole
    other.topElevation = Me.topElevation

    Dim i As Integer
    For i = LBound(pLayerDepth) To UBound(pLayerDepth)
        Call other.addLayer(pLayerDepth(i), pLayerFrontResistance(i), pLayerSideFriction(i))
    Next i

    Set DeepCopy = other
End Function

Private Function isDepthWithinBorehole(depth As Double) As Boolean

    If depth < pLayerDepth(LBound(pLayerDepth)) Or depth > pLayerDepth(UBound(pLayerDepth)) Then
        MsgBox "C_Borehole - Depth of borehole " & nameOfBorehole & " <= than " & depth, vbOK, "ERROR"
        isDepthWithinBorehole = False
    Else
        isDepthWithinBorehole = True
    End If

End Function

' Лобовое сопротивление на глубине
Public Function getFrontResistanceAtDepth(depth As Double) As Double
    If isDepthWithinBorehole(depth) = False Then Exit Function
    getFrontResistanceAtDepth = math.interpolateOneDimensionalArray(depth, pLayerDepth, pLayerFrontResistance)
End Function

' Боковое сопротивление на глубине
Public Function getSideFrictionAtDepth(depth As Double) As Double
    If isDepthWithinBorehole(depth) = False Then Exit Function
    getSideFrictionAtDepth = math.interpolateOneDimensionalArray(depth, pLayerDepth, pLayerSideFriction)
End Function

' Извлекает часть массива между глубинами
Private Function getArrayBtwDepth(arr As Variant, depth1 As Double, depth2 As Double) As Variant

    If depth1 < pLayerDepth(LBound(pLayerDepth)) Then depth1 = pLayerDepth(LBound(pLayerDepth))
    If depth2 < pLayerDepth(LBound(pLayerDepth)) Then depth2 = pLayerDepth(LBound(pLayerDepth))

    Dim d As Double
    If depth2 < depth1 Then
        d = depth1
        depth1 = depth2
        depth2 = d
    End If

    Dim i As Integer
    Dim result() As Double: ReDim result(0)
    Dim isInside As Boolean, isFirst As Boolean
    isInside = False: isFirst = True
    
    For i = LBound(pLayerDepth) To UBound(pLayerDepth)
    
        If pLayerDepth(i) >= depth1 Then
            isInside = True
        End If
        
        If pLayerDepth(i) >= depth2 Then
            result(UBound(result)) = math.interpolateOneDimensionalArray(depth2, pLayerDepth, arr)
            ReDim Preserve result(UBound(result) + 1)
            Exit For
        End If
        
        If isInside Then
        
            If isFirst = True Then
            
                If (depth1 < pLayerDepth(i)) Then
                    result(UBound(result)) = math.interpolateOneDimensionalArray(depth1, pLayerDepth, arr)
                    ReDim Preserve result(UBound(result) + 1)
                End If
                
                isFirst = False
                
            End If
            
            result(UBound(result)) = arr(i)
            ReDim Preserve result(UBound(result) + 1)
            
        End If
        
    Next i

    ReDim Preserve result(UBound(result) - 1)
    getArrayBtwDepth = result
End Function

' Извлекает часть массива pLayerDepth между глубинами
Public Function getDepthArrayBtwDepth(depth1 As Double, depth2 As Double) As Variant
    getDepthArrayBtwDepth = getArrayBtwDepth(pLayerDepth, depth1, depth2)
End Function

' Извлекает часть массива pLayerSideFriction между глубинами
Public Function getSideFrictionArrayBtwDepth(depth1 As Double, depth2 As Double) As Variant
    getSideFrictionArrayBtwDepth = getArrayBtwDepth(pLayerSideFriction, depth1, depth2)
End Function

' Извлекает часть массива pLayerFrontResistance между глубинами
Public Function getFrontResistanceArrayBtwDepth(depth1 As Double, depth2 As Double) As Variant
    getFrontResistanceArrayBtwDepth = getArrayBtwDepth(pLayerFrontResistance, depth1, depth2)
End Function
