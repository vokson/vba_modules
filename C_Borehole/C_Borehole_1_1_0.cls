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
Private pTopElevation As Double 'Абсолютная отметка (м)
Private pWaterDepth As Double ' Глубина воды (м)
Private pLayerDepth() As Double ' Глубина низа слоя грунта (м)
Private pLayerName() As String ' Имя слоя грунта (м)
Private countOfLayers As Integer

Private Sub Class_Initialize() 
    countOfLayers = 0
    pWaterDepth = -1000
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
        MsgBox "C_Borehole - Низ слоя " & name & " должен быть > 0. Скважина " & nameOfBorehole, vbOK, "ERROR"
        Exit Sub
    End If

    ReDim Preserve pLayerDepth(countOfLayers)
    ReDim Preserve pLayerName(countOfLayers)
    countOfLayers = countOfLayers + 1

    pLayerDepth(UBound(pLayerDepth)) = bottomOfLayer
    pLayerName(UBound(pLayerName)) = name

End Sub

' Получить имя грунта на определенной глубине
Public Property Get getSoilNameAtDepth(depth As Double) As String
    
    If depth >=  pLayerDepth(UBound(pLayerDepth)) Then
        MsgBox "C_Borehole - Увеличьте глубину скважины. Глубина " _
             & nameOfBorehole & " д.б. > " & depth, vbOK, "ERROR"
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

' Есть ли вода на данной глубине
Public Property Get isWaterAtDepth(depth As Double) As Boolean
    
    If (depth >= pWaterDepth) Then
        isWaterAtDepth = True
    Else
        isWaterAtDepth = False
    End If
    
End Property

Public Function DeepCopy() As C_Borehole
    Dim other As New C_Borehole

    other.nameOfBorehole = Me.nameOfBorehole
    other.topElevation = Me.topElevation
    other.waterDepth = Me.waterDepth

    If (Not Not pLayerDepth) <> 0 Then ' Array has been initialized, so you're good to go.

        Dim i As Integer
        For i = LBound(pLayerDepth) To UBound(pLayerDepth)
            Call other.addLayer(pLayerName(i), pLayerDepth(i))
        Next i

    End if

    Set DeepCopy = other
End Function
