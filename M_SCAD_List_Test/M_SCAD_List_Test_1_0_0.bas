Attribute VB_Name = "M_SCAD_List_Test"
Option Explicit

Public Sub testMakeArrayFromList()
    Dim list As New C_SCAD_List
    Dim math As New C_Math
    
    Dim s1 As String : s1 = "3-5"
    Dim arr1() As Variant: arr1 = Array(3, 4, 5)
    
    Debug.Print "makeArrayFromList: TEST 01"
    If math.isArraysSame(list.makeArrayFromList(s1), arr1) Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If
    
    
    Set math = Nothing
    Set list = Nothing
End Sub