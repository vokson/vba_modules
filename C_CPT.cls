
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

