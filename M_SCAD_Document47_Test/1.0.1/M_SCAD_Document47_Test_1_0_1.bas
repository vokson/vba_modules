Attribute VB_Name = "M_SCAD_Document47_Test"
Option Explicit

Public Sub testGetGroups()
    Dim doc47 As New C_SCAD_Document47
    Dim math As New C_Math
    Dim dic As Dictionary
    
    Dim s As String : s = "jkgej (47/Name=" & Chr(34) & "Piles" & Chr(34) & " 2  : 6289 7531-7533/" & vbCrLf & "Name=" & _
        Chr(34) & "Axis 1" & Chr(34) & " 2  : 7538 r 7546 8 /Name=" & Chr(34) & _
         "Ax/ is/ 2" & Chr(34) & " 2  : 7618 7729 /)j grr "
    Dim arr1() As Variant: arr1 = Array(6289, 7531, 7532, 7533)
    Dim arr2() As Variant: arr2 = Array(7538, 7546)
    Dim arr3() As Variant: arr3 = Array(7618, 7729)

    Debug.Print "getGroups: TEST 01"
    Set dic = doc47.getGroups(s)
    If math.isArraysSame(dic.Item("Piles"), arr1) And _
       math.isArraysSame(dic.Item("Axis 1"), arr2) And _
       math.isArraysSame(dic.Item("Ax/is/2"), arr3) _
     Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If
    
    Set dic = Nothing
    Set math = Nothing
    Set doc47 = Nothing
End Sub