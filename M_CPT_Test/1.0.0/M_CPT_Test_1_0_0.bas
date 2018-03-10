Attribute VB_Name = "M_CPT_Test"
Option Explicit

Private math As New C_Math

Public Sub testAll()
    Dim cpt As New C_CPT
    cpt.nameOfBorehole = "Test_CPT"
    cpt.topElevation = 111.11

    Call cpt.addLayer(1.0, 10.0, 100.0)
    Call cpt.addLayer(2.0, 20.0, 200.0)
    Call cpt.addLayer(3.0, 30.0, 300.0)
    Call cpt.addLayer(4.0, 40.0, 400.0)
    Call cpt.addLayer(5.0, 50.0, 500.0)

    Dim copyOfCpt As C_CPT
    Set copyOfCpt = cpt.DeepCopy()

    Call testCPT(cpt)
    Call testCPT(copyOfCpt)

    Call testDepth(cpt)
    Call testDepth(copyOfCpt)

    Call testFR(cpt)
    Call testFR(copyOfCpt)

    Call testSF(cpt)
    Call testSF(copyOfCpt)

    Set cpt = Nothing
    Set copyOfCpt = Nothing

End Sub

Public Sub testCPT(cpt As C_CPT)
    
    Debug.Print "TEST 01 - nameOfBorehole"
    If cpt.nameOfBorehole = "Test_CPT" Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

   
    Debug.Print "TEST 02 - topElevation"
    If cpt.topElevation = 111.11 Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

   
    Debug.Print "TEST 03 - getFrontResistanceAtDepth"
    If cpt.getFrontResistanceAtDepth(1.0) = 10.0 Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "TEST 04 - getFrontResistanceAtDepth"
    If cpt.getFrontResistanceAtDepth(1.5) = 15.0 Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "TEST 05 - getFrontResistanceAtDepth"
    If cpt.getFrontResistanceAtDepth(5.0) = 50.0 Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "TEST 06 - getSideFrictionAtDepth"
    If cpt.getSideFrictionAtDepth(1.0) = 100.0 Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "TEST 07 - getSideFrictionAtDepth"
    If cpt.getSideFrictionAtDepth(1.5) = 150.0 Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "TEST 08 - getSideFrictionAtDepth"
    If cpt.getSideFrictionAtDepth(5.0) = 500.0 Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

End Sub

Public Sub testSF(cpt As C_CPT)
    
    Debug.Print "TEST 01 - getSideFrictionArrayBtwDepth"
    If math.isArraysSame(Array(100#, 200#), cpt.getSideFrictionArrayBtwDepth(0#, 2#)) Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "TEST 02 - getSideFrictionArrayBtwDepth"
    If math.isArraysSame(Array(100#, 200#, 250#), cpt.getSideFrictionArrayBtwDepth(1#, 2.5)) Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "TEST 03 - getSideFrictionArrayBtwDepth"
    If math.isArraysSame(Array(150#, 200#, 300#, 400#, 450#), cpt.getSideFrictionArrayBtwDepth(1.5, 4.5)) Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "TEST 04 - getSideFrictionArrayBtwDepth"
    If math.isArraysSame(Array(450#, 500#), cpt.getSideFrictionArrayBtwDepth(4.5, 6.0)) Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

End Sub

Public Sub testFR(cpt As C_CPT)
    
    Debug.Print "TEST 01 - getFrontResistanceArrayBtwDepth"
    If math.isArraysSame(Array(10#, 20#), cpt.getFrontResistanceArrayBtwDepth(0#, 2#)) Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "TEST 02 - getFrontResistanceArrayBtwDepth"
    If math.isArraysSame(Array(10#, 20#, 25#), cpt.getFrontResistanceArrayBtwDepth(1#, 2.5)) Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "TEST 03 - getFrontResistanceArrayBtwDepth"
    If math.isArraysSame(Array(15#, 20#, 30#, 40#, 45#), cpt.getFrontResistanceArrayBtwDepth(1.5, 4.5)) Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "TEST 04 - getFrontResistanceArrayBtwDepth"
    If math.isArraysSame(Array(45#, 50#), cpt.getFrontResistanceArrayBtwDepth(4.5, 6.0)) Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

End Sub

Public Sub testDepth(cpt As C_CPT)
    
    Debug.Print "TEST 01- getDepthArrayBtwDepth"
    If math.isArraysSame(Array(1#, 2#), cpt.getDepthArrayBtwDepth(0#, 2#)) Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "TEST 02 - getDepthArrayBtwDepth"
    If math.isArraysSame(Array(1#, 2#, 2.5), cpt.getDepthArrayBtwDepth(1#, 2.5)) Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "TEST 03 - getDepthArrayBtwDepth"
    If math.isArraysSame(Array(1.5, 2#, 3#, 4#, 4.5), cpt.getDepthArrayBtwDepth(1.5, 4.5)) Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "TEST 04 - getDepthArrayBtwDepth"
    If math.isArraysSame(Array(4.5, 5#), cpt.getDepthArrayBtwDepth(4.5, 6.0)) Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

End Sub

