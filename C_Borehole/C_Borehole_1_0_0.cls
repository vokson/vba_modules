VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Borehole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pName As String
Private pTopElevation As Double
Private pWaterDepth As Double
Private pLayerDepth() As Double
Private pLayerName() As String
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

' Уровень подземных вод
Property Let waterDepth(depth As Double)
     pWaterDepth  = depth
End Property

Property Get waterDepth() As Double
    waterDepth = pWaterDepth
End Property

' Добавление слоя
Public Sub addLayer(name As String, bottomOfLayer As Double)

    If bottomOfLayer <=0 Then
        MsgBox "C_Borehole - Bottom of layer must be > 0", vbOK, "ERROR"
        Exit Sub
    End If

    countOfLayers = countOfLayers + 1

    ReDim Preserve pLayerDepth(countOfLayers)
    ReDim Preserve pLayerName(countOfLayers)

    pLayerDepth(UBound(pLayerDepth)) = bottomOfLayer
    pLayerName(UBound(pLayerName)) = name

End Sub

' Получить имя грунта на определенной глубине
Public Property Get getSoilNameAtDepth(depth As Double) As String
    
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
