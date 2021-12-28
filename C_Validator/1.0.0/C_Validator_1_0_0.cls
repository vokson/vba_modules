VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Validator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const VERSION = "1.0.0"

' ѕровер€ет, что число >= minNum
Public Function validateNumberGTE(value As String, Optional minNum As Integer = 0) As Boolean

    On Error GoTo Handler
    Dim id As Integer

    validateNumberGTE = True
    id = CInt(Value)
    If id < minNum Then GoTo Handler
    Exit Function
    
Handler:
    validateNumberGTE = False

End Function

' ѕровер€ет, что число имеет ведущий ноль
Public Function validateHasLeadingZeros(value As String) As Boolean
    validateHasLeadingZeros = False
    If Len(value) <= 1 Then Exit Function
    If Left(value, 1) <> "0" Then Exit Function
    validateHasLeadingZeros = True
End Function

' ѕровер€ет, что значение это число определенной длины
Public Function validateNumberWithLength(value As String, length As Integer) As Boolean
    Dim i As Integer
    Dim pattern As String

    pattern = ""
    For i=1 To length
      pattern = pattern & "#"
    Next i

    validateNumberWithLength = (value Like pattern)
End Function


Public Sub test()
  Call testValidateNumberGTE
  Call testValidateHasLeadingZeros
  Call testValidateNumberWithLength
End Sub

Public Sub testValidateNumberGTE()
    
    Dim tests As New Dictionary
    Dim data As Dictionary

    Dim testName As String
    testName = "validateNumberGTE"

    ' Normal 
    Set data = New Dictionary
    data.item("value") = "1"
    data.item("minNum") = 0
    data.item("answer") = True
    Set tests.item("01") = data

    ' Not number    
    Set data = New Dictionary
    data.item("value") = "A"
    data.item("minNum") = 0
    data.item("answer") = False
    Set tests.item("02") = data

    ' Less than min 
    Set data = New Dictionary
    data.item("value") = "1"
    data.item("minNum") = 2
    data.item("answer") = False
    Set tests.item("03") = data

    Dim idx
    Dim value As Variant
    Dim countOfCorrect As Integer
    countOfCorrect = 0

    Debug.Print "*** " & testName & " ***" 

    For Each idx In tests.Keys
        Set value = tests.item(idx)
        If validateNumberGTE(value.item("value"), value.Item("minNum")) = value.item("answer") Then
            Debug.Print "TEST: " & idx & " - PASSED"
            countOfCorrect = countOfCorrect + 1
        Else
            Debug.Print "TEST: " & idx & " - FAILED"
        End If
    Next
    
    Debug.Print "Success: " & Str(countOfCorrect) & " /" & Str(tests.Count) & " = " & Format(countOfCorrect / tests.Count * 100, "#,##0.00") & "%"
    Debug.Print " "
End Sub

Public Sub testValidateHasLeadingZeros()
    
    Dim tests As New Dictionary
    Dim data As Dictionary

    Dim testName As String
    testName = "validateHasLeadingZeros"

    ' Empty 
    Set data = New Dictionary
    data.item("value") = ""
    data.item("answer") = False
    Set tests.item("01") = data

    ' Zero 
    Set data = New Dictionary
    data.item("value") = "0"
    data.item("answer") = False
    Set tests.item("02") = data

    ' Number without zeros 
    Set data = New Dictionary
    data.item("value") = "01"
    data.item("answer") = True
    Set tests.item("03") = data

    ' Symbol without zeros 
    Set data = New Dictionary
    data.item("value") = "0A"
    data.item("answer") = True
    Set tests.item("04") = data


    Dim idx
    Dim value As Variant
    Dim countOfCorrect As Integer
    countOfCorrect = 0

    Debug.Print "*** " & testName & " ***" 
    
    For Each idx In tests.Keys
        Set value = tests.item(idx)
        If validateHasLeadingZeros(value.item("value")) = value.item("answer") Then
            Debug.Print "TEST: " & idx & " - PASSED"
            countOfCorrect = countOfCorrect + 1
        Else
            Debug.Print "TEST: " & idx & " - FAILED"
        End If
    Next
    
    Debug.Print "Success: " & Str(countOfCorrect) & " /" & Str(tests.Count) & " = " & Format(countOfCorrect / tests.Count * 100, "#,##0.00") & "%"
    Debug.Print " "
End Sub

Public Sub testValidateNumberWithLength()
    
    Dim tests As New Dictionary
    Dim data As Dictionary

    Dim testName As String
    testName = "validateNumberWithLength"

    ' OK 
    Set data = New Dictionary
    data.item("value") = "102"
    data.item("length") = 3
    data.item("answer") = True
    Set tests.item("01") = data

    ' Wrong length 
    Set data = New Dictionary
    data.item("value") = "102"
    data.item("length") = 2
    data.item("answer") = False
    Set tests.item("02") = data

    ' Not number 
    Set data = New Dictionary
    data.item("value") = "A02"
    data.item("length") = 3
    data.item("answer") = False
    Set tests.item("03") = data

    Dim idx
    Dim value As Variant
    Dim countOfCorrect As Integer
    countOfCorrect = 0

    Debug.Print "*** " & testName & " ***" 
    
    For Each idx In tests.Keys
        Set value = tests.item(idx)
        If validateNumberWithLength(value.item("value"), value.item("length")) = value.item("answer") Then
            Debug.Print "TEST: " & idx & " - PASSED"
            countOfCorrect = countOfCorrect + 1
        Else
            Debug.Print "TEST: " & idx & " - FAILED"
        End If
    Next
    
    Debug.Print "Success: " & Str(countOfCorrect) & " /" & Str(tests.Count) & " = " & Format(countOfCorrect / tests.Count * 100, "#,##0.00") & "%"
    Debug.Print " "
End Sub