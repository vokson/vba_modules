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