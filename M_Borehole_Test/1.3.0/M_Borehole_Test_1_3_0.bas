Attribute VB_Name = "M_Borehole_Test"
Option Explicit

Public Sub test1()
    Dim brh As New C_Borehole
    brh.nameOfBorehole = "Test_BRH"
    brh.topElevation = 111.11
    brh.waterDepth = 4.5

    Call brh.addLayer("Layer 1", 1.5)
    Call brh.addLayer("Layer 2", 3.0)
    Call brh.addLayer("Layer 3", 4.0)
    Call brh.addLayer("Layer 4", 5.0)
    Call brh.addLayer("Layer 5", 10.0)

    Dim copyOfBrh As C_Borehole
    Set copyOfBrh = brh.DeepCopy()

    Call testBorehole(brh)
    Call testBorehole(copyOfBrh)

    Set brh = Nothing
    Set copyOfBrh = Nothing

End Sub

Public Sub testBorehole(brh As C_Borehole)
    
    Debug.Print "TEST 01 - nameOfBorehole"
    If brh.nameOfBorehole = "Test_BRH" Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

   
    Debug.Print "TEST 02 - topElevation"
    If brh.topElevation = 111.11 Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

   
    Debug.Print "TEST 03 - waterDepth"
    If brh.waterDepth = 4.5 Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If
    

    Debug.Print "TEST 04 - getSoilNameAtDepth"
    If brh.getSoilNameAtDepth(1.0) = "Layer 1" Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "TEST 05 - getSoilNameAtDepth"
    If brh.getSoilNameAtDepth(1.5) = "Layer 2" Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "TEST 06 - getSoilNameAtDepth"
    If brh.getSoilNameAtDepth(6.0) = "Layer 5" Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "TEST 07 - getSoilNameAtDepth"
    If brh.getSoilNameAtDepth(10.0) = "Layer 5" Then
        Debug.Print "FAILED"
    Else
        Debug.Print "PASSED"
    End If

    Debug.Print "TEST 08 - isWaterAtDepth"
    If brh.isWaterAtDepth(4.0) = False Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "TEST 09 - isWaterAtDepth"
    If brh.isWaterAtDepth(4.5) = True Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    
End Sub

Public Sub test2_Cut()
    Dim brh As New C_Borehole
    brh.nameOfBorehole = "Test_BRH"
    brh.topElevation = 100.00
    brh.waterDepth = 4.5

    Call brh.addLayer("Layer 1", 1.0)
    Call brh.addLayer("Layer 2", 2.0)
    Call brh.addLayer("Layer 3", 3.0)
    Call brh.addLayer("Layer 4", 4.0)
    Call brh.addLayer("Layer 5", 5.0)

    Dim copyOfBrh As C_Borehole
    Set copyOfBrh = brh.DeepCopyWithOtherTopElevation(98.5, "")

    Dim correct As New C_Borehole
    correct.nameOfBorehole = "Test_BRH"
    correct.topElevation = 98.50
    correct.waterDepth = 3.0

    Call correct.addLayer("Layer 2", 0.5)
    Call correct.addLayer("Layer 3", 1.5)
    Call correct.addLayer("Layer 4", 2.5)
    Call correct.addLayer("Layer 5", 3.5)

    If copyOfBrh.isSameWith(correct)  Then
        Debug.Print "test2_Cut - PASSED"
    Else
        Debug.Print "test2_Cut - FAILED"
    End If

    Set brh = Nothing
    Set copyOfBrh = Nothing

End Sub

Public Sub test2_Fill()
    Dim brh As New C_Borehole
    brh.nameOfBorehole = "Test_BRH"
    brh.topElevation = 100.00
    brh.waterDepth = 4.5

    Call brh.addLayer("Layer 1", 1.0)
    Call brh.addLayer("Layer 2", 2.0)
    Call brh.addLayer("Layer 3", 3.0)
    Call brh.addLayer("Layer 4", 4.0)
    Call brh.addLayer("Layer 5", 5.0)

    Dim copyOfBrh As C_Borehole
    Set copyOfBrh = brh.DeepCopyWithOtherTopElevation(102.5, "Fill Soil")

    Dim correct As New C_Borehole
    correct.nameOfBorehole = "Test_BRH"
    correct.topElevation = 102.50
    correct.waterDepth = 7.0

    Call correct.addLayer("Fill Soil", 2.5)
    Call correct.addLayer("Layer 1", 3.5)
    Call correct.addLayer("Layer 2", 4.5)
    Call correct.addLayer("Layer 3", 5.5)
    Call correct.addLayer("Layer 4", 6.5)
    Call correct.addLayer("Layer 5", 7.5)

    If copyOfBrh.isSameWith(correct) Then
        Debug.Print "test2_Fill - PASSED"
    Else
        Debug.Print "test2_Fill - FAILED"
    End If

    Set brh = Nothing
    Set copyOfBrh = Nothing

End Sub

Public Sub test3_SplitAtDepth()
    Dim brh As New C_Borehole
    brh.nameOfBorehole = "Test_BRH"
    brh.topElevation = 100#
    brh.waterDepth = 4.5

    Call brh.addLayer("Layer 1", 1#)
    Call brh.addLayer("Layer 2", 2#)
    Call brh.addLayer("Layer 3", 3#)
    Call brh.addLayer("Layer 4", 4#)
    Call brh.addLayer("Layer 5", 5#)
    
    ' TEST 01
    
    Dim test_brh As C_Borehole
    Set test_brh = brh.DeepCopy
    
    Call test_brh.splitAtDepth(6#)
    
    If test_brh.isSameWith(brh) Then
        Debug.Print "Test 1 - PASSED"
    Else
        Debug.Print "Test 1 - FAILED"
    End If
    
    Set test_brh = Nothing
    
    ' TEST 02
    
    Set test_brh = brh.DeepCopy
    
    Call test_brh.splitAtDepth(4.005)
    
    If test_brh.isSameWith(brh) Then
        Debug.Print "Test 2 - PASSED"
    Else
        Debug.Print "Test 2 - FAILED"
    End If
    
    Set test_brh = Nothing
    
    ' TEST 03
    
    Set test_brh = brh.DeepCopy
    
    Call test_brh.splitAtDepth(2.5)
    
    Dim correct As New C_Borehole
    correct.nameOfBorehole = "Test_BRH"
    correct.topElevation = 100#
    correct.waterDepth = 4.5

    Call correct.addLayer("Layer 1", 1#)
    Call correct.addLayer("Layer 2", 2#)
    Call correct.addLayer("Layer 3", 2.5)
    Call correct.addLayer("Layer 3", 3#)
    Call correct.addLayer("Layer 4", 4#)
    Call correct.addLayer("Layer 5", 5#)
    
    If test_brh.isSameWith(correct) Then
        Debug.Print "Test 3 - PASSED"
    Else
        Debug.Print "Test 3 - FAILED"
    End If
    
    Set test_brh = Nothing
    Set correct = Nothing

    

    Set brh = Nothing
    

End Sub

Public Sub test3_SplitAtWaterDepth()
    Dim brh As New C_Borehole
    brh.nameOfBorehole = "Test_BRH"
    brh.topElevation = 100#
    brh.waterDepth = 2.5

    Call brh.addLayer("Layer 1", 1#)
    Call brh.addLayer("Layer 2", 2#)
    Call brh.addLayer("Layer 3", 3#)
    Call brh.addLayer("Layer 4", 4#)
    Call brh.addLayer("Layer 5", 5#)
    
    ' TEST 01
    
    Dim test_brh As C_Borehole
    Set test_brh = brh.DeepCopy
    
    Call test_brh.splitAtWaterdepth()
    
    Dim correct As New C_Borehole
    correct.nameOfBorehole = "Test_BRH"
    correct.topElevation = 100#
    correct.waterDepth = 2.5

    Call correct.addLayer("Layer 1", 1#)
    Call correct.addLayer("Layer 2", 2#)
    Call correct.addLayer("Layer 3", 2.5)
    Call correct.addLayer("Layer 3", 3#)
    Call correct.addLayer("Layer 4", 4#)
    Call correct.addLayer("Layer 5", 5#)
    
    If test_brh.isSameWith(correct) Then
        Debug.Print "Test - PASSED"
    Else
        Debug.Print "Test - FAILED"
    End If
    
    Set test_brh = Nothing
    Set correct = Nothing

    Set brh = Nothing
End Sub