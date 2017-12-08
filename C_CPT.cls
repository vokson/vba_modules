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

Public name As String
Public topElevation As Double
Private pLayerDepth As Variant
Private pLayerSideFriction As Variant
Private pLayerFrontResistance As Variant

Property Let layerDepth(dArray As Variant)
    pLayerDepth = dArray
End Property

Property Let layerSideFriction(dArray As Variant)
    pLayerSideFriction = dArray
End Property

Property Let layerFrontResistance(dArray As Variant)
    pLayerFrontResistance = dArray
End Property

Public Property Get getSideFrictionAtDepth(depth As Double) As String
    
    Dim i As Integer
    For i = LBound(pLayerDepth) To UBound(pLayerDepth)
        If pLayerDepth(i) >= depth Then
            getSideFrictionAtDepth = pLayerSideFriction(i)
            Exit For
        End If
        
    Next i
    
    If getSideFrictionAtDepth = "" Then
        Debug.Print "Can NOT find Side Friction for " & name & " borehole at " & depth & " depth"
    End If
End Property

Public Property Get getFrontResistanceAtDepth(depth As Double) As String
    
    Dim i As Integer
    For i = LBound(pLayerDepth) To UBound(pLayerDepth)
        If pLayerDepth(i) >= depth Then
            getFrontResistanceAtDepth = pLayerFrontResistance(i)
            Exit For
        End If
    Next i
    
    If getFrontResistanceAtDepth = "" Then
        Debug.Print "Can NOT find Front Resistance for " & name & " borehole at " & depth & " depth"
    End If
End Property

Public Property Get getArrayBtwDepth(arr As Variant, depth1 As Double, depth2 As Double) As Variant

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
            result(UBound(result)) = interpolateOneDimensionalArray(depth2, pLayerDepth, arr)
            ReDim Preserve result(UBound(result) + 1)
            Exit For
        End If
        
        If isInside Then
        
            If isFirst = True Then
            
                If (depth1 < pLayerDepth(i)) Then
                    result(UBound(result)) = interpolateOneDimensionalArray(depth1, pLayerDepth, arr)
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
End Property

Public Property Get getDepthArrayBtwDepth(depth1 As Double, depth2 As Double) As Variant
    getDepthArrayBtwDepth = getArrayBtwDepth(pLayerDepth, depth1, depth2)
End Property

Public Property Get getSideFrictionArrayBtwDepth(depth1 As Double, depth2 As Double) As Variant
    getSideFrictionArrayBtwDepth = getArrayBtwDepth(pLayerSideFriction, depth1, depth2)
End Property

Public Property Get getFrontResistanceArrayBtwDepth(depth1 As Double, depth2 As Double) As Variant
    getFrontResistanceArrayBtwDepth = getArrayBtwDepth(pLayerFrontResistance, depth1, depth2)
End Property

