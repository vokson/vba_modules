Attribute VB_Name = "M_String_Test"
Option Explicit

Public Sub testDeleteDoubleSpaces()
    Dim obj
    Set obj = New C_String
    
    Dim s1 As String
    s1 = "  34 44  55 44 776 gkg km  jthj  tht   tohjt tojh     jthjt "

    Dim s1_correct As String
    s1_correct = " 34 44 55 44 776 gkg km jthj tht tohjt tojh jthjt "
    
    Debug.Print "deleteDoubleSpaces: TEST 01"
    If obj.deleteDoubleSpaces(s1) = s1_correct Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Set obj = Nothing
End Sub

Public Sub tesReplaceNewStringSymbolWith()
    Dim obj
    Set obj = New C_String
    
    Dim s1 As String
    s1 = "ABC" & vbCrLf & "DEF"

    Dim s1_correct As String
    s1_correct = "ABC#DEF"

    Dim s2 As String
    s2 = "ABC" & vbCr & "DEF"

    Dim s2_correct As String
    s2_correct = "ABC#DEF"

    Dim s3 As String
    s3 = "ABC" & vbLf & "DEF"

    Dim s3_correct As String
    s3_correct = "ABC#DEF"

    Dim s4 As String
    s4 = "ABC" & vbCrLf & vbLf & "DEF"

    Dim s4_correct As String
    s4_correct = "ABC##DEF"
    
    Debug.Print "replaceNewStringSymbolWith: TEST 01"
    If obj.replaceNewStringSymbolWith(s1, "#") = s1_correct Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "replaceNewStringSymbolWith: TEST 02"
    If obj.replaceNewStringSymbolWith(s2, "#") = s2_correct Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "replaceNewStringSymbolWith: TEST 03"
    If obj.replaceNewStringSymbolWith(s3, "#") = s3_correct Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "replaceNewStringSymbolWith: TEST 04"
    If obj.replaceNewStringSymbolWith(s4, "#") = s4_correct Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Set obj = Nothing
End Sub