VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Borehole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public name As String
Public topElevation As Double
Public waterDepth As Double
Private pLayerDepth As Variant
Private pLayerName As Variant

Property Let layerDepth(dArray As Variant)
    pLayerDepth = dArray
End Property

Property Let layerName(dArray As Variant)
    pLayerName = dArray
End Property

Public Property Get getSoilNameAtDepth(depth As Double) As String
    Dim i As Integer, length As Integer, sumLayerDepth As Double
'    sumLayerDepth = 0
'    length = pLayerDepth.length
    
    For i = 0 To UBound(pLayerDepth)
        If pLayerDepth(i) > depth Then
            getSoilNameAtDepth = pLayerName(i)
            Exit For
        End If
        
'        sumLayerDepth = sumLayerDepth + pLayerDepth(i)
    Next i
    
    If getSoilNameAtDepth = "" Then
        getSoilNameAtDepth = pLayerName(UBound(pLayerName))
    End If
End Property
