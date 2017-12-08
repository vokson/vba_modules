Attribute VB_Name = "CALCULATION"
Option Explicit

' depth - глубина от-но устья скважины
Function getR_ByBoreholeName(nameOfBorehole As String, depth As Double, Optional isDensityAsPerCPTWithoutHoles As Boolean = False) As Double
    Dim nameOfSoil As String
    nameOfSoil = getBoreholeSoilNameAtDepth(nameOfBorehole, depth)
    
    getR_ByBoreholeName = getR_BySoilName(nameOfSoil, depth, isDensityAsPerCPTWithoutHoles)
End Function

Function getR_BySoilName(nameOfSoil As String, depth As Double, Optional isDensityAsPerCPTWithoutHoles As Boolean = False) As Double
    Dim newSoil As soil

    Set newSoil = getSoilByName(nameOfSoil)
    
    getR_BySoilName = getTable7_2(depth, newSoil.typeOfSoil, newSoil.subtypeOfSoil, newSoil.density, _
            newSoil.IL, newSoil.IP, newSoil.e, isDensityAsPerCPTWithoutHoles)
End Function

Function getFiAtDepth(nameOfSoil As String, depth As Double, depthOfSeismic As Double, _
                                    depthOfBulk As Double) As Double
    Dim newSoil As soil

    Set newSoil = getSoilByName(nameOfSoil)
    getFiAtDepth = getTable7_4(depth, newSoil.typeOfSoil, newSoil.subtypeOfSoil, newSoil.density, newSoil.IL, newSoil.e)
    
    If (depth <= depthOfBulk) Then getFiAtDepth = -getFiAtDepth
    If (depth <= depthOfSeismic) And (getFiAtDepth > 0) Then getFiAtDepth = 0
    
End Function

' depth1, depth2 - глубина от-но устья скважины
Function getF_ByBoreholeName(nameOfBorehole As String, depth1 As Double, _
            depth2 As Double, depthOfSeismic As Double, depthOfBulk As Double) As Double
            
    Dim STEP_FOR_TABLE_7_4 As Double: STEP_FOR_TABLE_7_4 = 0.1
    getF_ByBoreholeName = 0#
    
    If (depth1 = depth2) Then
        Exit Function
    End If
    
    If (depthOfBulk < 0) Then depthOfBulk = 0
    If (depthOfSeismic < 0) Then depthOfSeismic = 0
    
    Dim top As Double, bottom As Double, depth As Double
    If (depth1 < depth2) Then
        bottom = depth1: top = depth2
    Else
        top = depth1: bottom = depth2
    End If
    
    
    Dim count As Integer, i As Integer, step As Double, nameOfSoil As String
    count = Fix((top - bottom) / STEP_FOR_TABLE_7_4)
    
    If (count <= 0) Then
        getF_ByBoreholeName = 0
        Exit Function
    End If
    
    step = (top - bottom) / count
    
    
    For i = 0 To count - 1
        depth = bottom + step * (i + 0.5)
        nameOfSoil = getBoreholeSoilNameAtDepth(nameOfBorehole, depth)
'
        Debug.Print depth & " - " & nameOfSoil & " = " & step * getFiAtDepth(nameOfSoil, depth, depthOfSeismic, depthOfBulk)
        Debug.Print getF_ByBoreholeName
        
        getF_ByBoreholeName = getF_ByBoreholeName + step * getFiAtDepth(nameOfSoil, depth, depthOfSeismic, depthOfBulk)
    Next i
    
    Debug.Print "FINISH"
    
End Function

' depth1, depth2 - глубина от-но устья скважины
Function getG_or_Nu(isGtoBeCalculated As Boolean, nameOfBorehole As String, depth1 As Double, depth2 As Double) As Double
    Dim STANDARD_STEP As Double: STANDARD_STEP = 0.1:
    
    Dim resultG As Double, resultNu As Double
    resultG = 0: resultNu = 0
    
    If (depth1 = depth2) Then
        Exit Function
    End If
    
    Dim top As Double, bottom As Double, depth As Double
    If (depth1 < depth2) Then
        bottom = depth1: top = depth2
    Else
        top = depth1: bottom = depth2
    End If
    
    
    Dim count As Integer, i As Integer, step As Double, nameOfSoil As String
    count = Fix((top - bottom) / STANDARD_STEP)
    step = (top - bottom) / count
    
    Dim G As Double, nu As Double, soilAtDepth As soil
    
    For i = 0 To count - 1
        depth = bottom + step * (i + 0.5)
        nameOfSoil = getBoreholeSoilNameAtDepth(nameOfBorehole, depth)
        
        Set soilAtDepth = getSoilByName(nameOfSoil)
        
        nu = getTable5_10(soilAtDepth.typeOfSoil, soilAtDepth.IL)
        resultNu = resultNu + nu
        G = getClause7_4_3(soilAtDepth.young_modulus, nu)
        resultG = resultG + G
        
'        Debug.Print nameOfSoil & " : " & depth
'        Debug.Print "E = " & soilAtDepth.young_modulus & ", nu = " & nu & ", G = " & G
    Next i
    
    If isGtoBeCalculated = True Then
        getG_or_Nu = resultG
    Else
        getG_or_Nu = resultNu
    End If
    
    getG_or_Nu = getG_or_Nu / count
    
    
End Function

Function getQs_ByCptName( _
        nameOfCpt As String, _
        d As Double, _
        depthOfPileBottom As Double _
    ) As Double
    
    Dim cpt As C_CPT
    Set cpt = getCptByName(nameOfCpt)
    
    Dim depthArray, frontResistanceArray, depth1 As Double, depth2 As Double
    depth1 = depthOfPileBottom - d
    depth2 = depthOfPileBottom + 4 * d
'    Debug.Print "Depth1 = " & depth1
'    Debug.Print "Depth2 = " & depth2
    
    depthArray = getCptDepthArrayBtwDepth(nameOfCpt, depth1, depth2)
    frontResistanceArray = getCptFrontResistanceArrayBtwDepth(nameOfCpt, depth1, depth2)
    
    Dim i As Integer, totalQs As Double
    totalQs = 0
    For i = LBound(depthArray) To UBound(depthArray) - 1
        totalQs = totalQs + _
            (frontResistanceArray(i) + frontResistanceArray(i + 1)) / 2 * Abs(depthArray(i + 1) - depthArray(i))
    Next i
    
    getQs_ByCptName = totalQs / (depthArray(UBound(depthArray)) - depthArray(LBound(depthArray)))
'    Debug.Print "totalQs = " & totalQs
'    Debug.Print "qs = " & getQs_ByCptName
    
End Function

Function getTotalFs_ByCptName( _
        nameOfCpt As String, _
        typeOfZond As Integer, _
        depth1 As Double, _
        depth2 As Double, _
        depthOfBulk As Double _
    ) As Double
    
    If (depth1 = depth2) Then
        Exit Function
    End If
    
    If (depthOfBulk < 0) Then depthOfBulk = 0
    
    Dim cpt As C_CPT, oSoil As soil, nameOfSoil As String
    
    Set cpt = getCptByName(nameOfCpt)
    
    Dim depthArray, sideFrictionArray
    depthArray = getCptDepthArrayBtwDepth(nameOfCpt, depth1, depth2)
    sideFrictionArray = getCptSideFrictionArrayBtwDepth(nameOfCpt, depth1, depth2)
    
    Dim i As Integer, totalFs As Double, fs As Double, Bi As Double, isSand As Boolean, depth As Double
    totalFs = 0: isSand = False
    
    For i = LBound(depthArray) To UBound(depthArray) - 1
    
        depth = depthArray(i)
        fs = (sideFrictionArray(i) + sideFrictionArray(i + 1)) / 2
    
        nameOfSoil = getBoreholeSoilNameAtDepth(nameOfCpt, depth)
        Set oSoil = getSoilByName(nameOfSoil)
        
        Select Case oSoil.typeOfSoil
        
            Case SOIL_TYPE_SAND, SOIL_TYPE_CLAY, SOIL_TYPE_CLAY_LOAM, SOIL_TYPE_CLAY_SANDY:
                If oSoil.typeOfSoil = SOIL_TYPE_SAND Then isSand = True
                Bi = getTable7_16_B2i(fs * 1000, typeOfZond, isSand)
                
            Case Else: Bi = 0
        End Select
        
        If (depth <= depthOfBulk) Then Bi = -Bi
    
        totalFs = totalFs + fs * Abs(depthArray(i + 1) - depthArray(i)) * Bi
        
    Next i
    
    getTotalFs_ByCptName = totalFs
    
End Function
