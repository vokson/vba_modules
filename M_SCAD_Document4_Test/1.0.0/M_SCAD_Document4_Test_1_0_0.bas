Attribute VB_Name = "M_SCAD_Document47_Test"
Option Explicit

Public Sub testAll()
    Call testWriteNodes_1()
End Sub

Public Sub testWriteNodes_1()
    Dim doc4 As New C_SCAD_Document4

    Dim nodes(1 to 3) as Dictionary
    Dim node As New Dictionary

    Set node = New Dictionary
    node.Item("X") = 0
    node.Item("Y") = 0
    node.Item("Z") = 0
    Set nodes(1) = node
    Set node = Nothing

    Set node = New Dictionary
    node.Item("X") = 1
    node.Item("Y") = 2
    node.Item("Z") = 3
    Set nodes(2) = node
    Set node = Nothing

    Set node = New Dictionary
    node.Item("X") = 4.4
    node.Item("Y") = 5.55
    node.Item("Z") = 6.666
    Set nodes(3) = node
    Set node = Nothing

    Dim test1 As String
    test1 = "(4/0 0 0/1 2 3/4.4 5.55 6.666/)"

    Debug.Print "writeNodes: TEST 01"
    If test1 = doc4.writeNodes(nodes) Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If
    
    Set node = Nothing
    Set doc4 = Nothing
End Sub