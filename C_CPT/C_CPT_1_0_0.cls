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

Private pName As String
Private pTopElevation As Double

Private pLayerDepth() As Double
Private pLayerFrontResistance() As Double
Private pLayerSideFriction() As Double

Private countOfLayers As Integer

Private Sub Class_Initialize() 
    countOfLayers = 0
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

' Боковое сопротивление на глубине
Public Property Get getFrontResistanceAtDepth(depth As Double) As Double
    
    If depth >=  pLayerDepth(UBound(pLayerDepth)) Then
        MsgBox "C_Borehole - Depth of borehole " & nameOfBorehole & " <= than " & depth, vbOK, "ERROR"
        Exit Sub
    End If

    Dim i As Integer
    For i = LBound(pLayerDepth) To UBound(pLayerDepth)
        If pLayerDepth(i) > depth Then
            getSoilNameAtDepth = pLayerName(i)
            Exit For
        End If
    Next i
    
End Property
