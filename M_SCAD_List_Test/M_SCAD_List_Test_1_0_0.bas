Attribute VB_Name = "M_SCAD_List_Test"
Option Explicit

Public Sub testMakeArrayFromList()
    Dim list As New C_SCAD_List
    Dim math As New C_Math
    
    Dim s1 As String : s1 = "3-5"
    Dim arr1() As Variant: arr1 = Array(3, 4, 5)

    Dim s2 As String : s2 = "3 r 11 2"
    Dim arr2() As Variant: arr2 = Array(3, 5, 7, 9, 11)

    Dim s3 As String : s3 = "2 3 r 11 2 23-25 34"
    Dim arr3() As Variant: arr3 = Array(2, 3, 5, 7, 9, 11, 23, 24, 25, 34)
    
    Debug.Print "makeArrayFromList: TEST 01"
    If math.isArraysSame(list.makeArrayFromList(s1), arr1) Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "makeArrayFromList: TEST 02"
    If math.isArraysSame(list.makeArrayFromList(s2), arr2) Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "makeArrayFromList: TEST 03"
    If math.isArraysSame(list.makeArrayFromList(s3), arr3) Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If
    
    
    Set math = Nothing
    Set list = Nothing
End Sub