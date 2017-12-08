Attribute VB_Name = "M_Borehole_Test"
Option Explicit

Public Sub test()
    Dim brh
    Set brh = New C_Borehole
    
    brh.nameOfBorehole = "Test_BRH"
    Debug.Print "TEST 01 - nameOfBorehole"
    If brh.nameOfBorehole = "Test_BRH" Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    brh.topElevation = 111.11
    Debug.Print "TEST 02 - topElevation"
    If brh.topElevation = 111.11 Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    brh.waterDepth = 222.22
    Debug.Print "TEST 03 - waterDepth"
    If brh.waterDepth = 222.22 Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Call brh.addLayer("Layer 1", 1.5)
    Call brh.addLayer("Layer 2", 3.0)
    Call brh.addLayer("Layer 3", 4.0)
    Call brh.addLayer("Layer 4", 5.0)
    Call brh.addLayer("Layer 5", 10.0)

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

    Set brh = Nothing
End Sub