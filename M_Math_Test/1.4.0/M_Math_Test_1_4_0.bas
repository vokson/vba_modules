Attribute VB_Name = "M_Math_Test"
Option Explicit

CONST VERSION = "1.4.0"

Public Sub testInterpolateOneDimensionalArray()
    Dim math
    Set math = New C_Math
    
    Dim arr1() As Variant: arr1 = Array(10, 20, 30, 40, 50)
    Dim arr2() As Variant: arr2 = Array(500.5, 400.4, 300.3, 200.2, 100.1)
    
    Debug.Print "interpolateOneDimensionalArray: TEST 01"
    If math.interpolateOneDimensionalArray(0, arr1, arr2) = 500.5 Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If
    
    Debug.Print "interpolateOneDimensionalArray: TEST 02"
    If math.interpolateOneDimensionalArray(10, arr1, arr2) = 500.5 Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If
    
    Debug.Print "interpolateOneDimensionalArray: TEST 03"
    If math.interpolateOneDimensionalArray(50, arr1, arr2) = 100.1 Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If
    
    Debug.Print "interpolateOneDimensionalArray: TEST 04"
    If math.interpolateOneDimensionalArray(25, arr1, arr2) = 350.35 Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If
    
    Debug.Print "interpolateOneDimensionalArray: TEST 05"
    If math.interpolateOneDimensionalArray(100, arr1, arr2) = 100.1 Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Set math = Nothing
End Sub

Public Sub testFindIdexOfFirstElementNotLessThan()
    Dim math
    Set math = New C_Math
    
    Dim arr1() As Variant: arr1 = Array(10#, 20#, 30#, 40#, 50#)
    
    Debug.Print "findIdexOfFirstElementNotLessThan: TEST 01"
    If math.findIdexOfFirstElementNotLessThan(0, arr1) = 0 Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If
    
    Debug.Print "findIdexOfFirstElementNotLessThan: TEST 02"
    If math.findIdexOfFirstElementNotLessThan(10, arr1) = 0 Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "findIdexOfFirstElementNotLessThan: TEST 03"
    If math.findIdexOfFirstElementNotLessThan(15, arr1) = 1 Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "findIdexOfFirstElementNotLessThan: TEST 04"
    If math.findIdexOfFirstElementNotLessThan(50, arr1) = 4 Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    ' Debug.Print "findIdexOfFirstElementNotLessThan: TEST 05"
    ' If math.findIdexOfFirstElementNotLessThan(60, arr1) Is Nothing Then
    '     Debug.Print "PASSED"
    ' Else
    '     Debug.Print "FAILED"
    ' End If

    Set math = Nothing
End Sub

Public Sub testFindIdexOfValueInArray()
    Dim math
    Set math = New C_Math
    
    Dim arr1() As Variant: arr1 = Array(10, 20, 30, 40, 50)
    Dim arr2() As Variant: arr2 = Array("10", "20", "30", "40", "50")
    
    Debug.Print "findIdexOfValueInArray: TEST 01"
    If math.findIdexOfValueInArray(10, arr1) = 0 Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If
    
    Debug.Print "findIdexOfValueInArray: TEST 02"
    If math.findIdexOfValueInArray(50, arr1) = 4 Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "findIdexOfValueInArray: TEST 03"
    If math.findIdexOfValueInArray(60, arr1) = -1 Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "findIdexOfValueInArray: TEST 04"
    If math.findIdexOfValueInArray("30", arr2) = 2 Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Set math = Nothing
End Sub

Public Sub testIsArraysSame()
    Dim math
    Set math = New C_Math
    
    Dim arr1() As Variant: arr1 = Array(10, 20, 30, 40, 50)
    Dim arr2() As Variant: arr2 = Array(10#, 20#, 30#, 40#, 50#)
    Dim arr3() As Variant: arr3 = Array(10.1, 20#, 30#, 40#, 50#)
    Dim arr4() As Variant: arr4 = Array("10", 20#, 30#, 40#, 50#)
    Dim arr5() As Variant: arr5 = Array(True, 20#, 30#, 40#, 50#)
    Dim arr6() As Variant: arr6 = Array(True, False)
    Dim arr7() As Variant: arr7 = Array(True, False)
    Dim arr8() As Variant: arr8 = Array(True, True)
    Dim arr9() As Variant: arr9 = Array("10", "20")
    Dim arr10() As Variant: arr10 = Array("10", "20")
    Dim arr11() As Variant: arr11 = Array("1", "20")
    
    Debug.Print "interpolateOneDimensionalArray: TEST 01"
    If math.isArraysSame(arr1, arr2) =True Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "interpolateOneDimensionalArray: TEST 02"
    If math.isArraysSame(arr1, arr3) = False Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "interpolateOneDimensionalArray: TEST 03"
    If math.isArraysSame(arr1, arr4) = True Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "interpolateOneDimensionalArray: TEST 04"
    If math.isArraysSame(arr1, arr5) = False Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "interpolateOneDimensionalArray: TEST 05"
    If math.isArraysSame(arr6, arr7) = True Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "interpolateOneDimensionalArray: TEST 06"
    If math.isArraysSame(arr6, arr8) = False Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "interpolateOneDimensionalArray: TEST 07"
    If math.isArraysSame(arr9, arr10) = True Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "interpolateOneDimensionalArray: TEST 08"
    If math.isArraysSame(arr9, arr11) = False Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If
    
    Set math = Nothing
End Sub

Public Sub testInterpolateTwoDimensionalArray()
    Dim math
    Set math = New C_Math
    
    Dim keyArray1() As Variant: keyArray1 = Array(10, 20, 30, 40, 50)
    Dim keyArray2() As Variant: keyArray2 = Array(1, 2, 3, 4, 5)
    Dim valueArray() As Variant: valueArray = Array( _
        Array(100, 200, 300, 400, 500), _
        Array(600, 700, 800, 900, 1000), _
        Array(1100, 1200, 1300, 1400, 1500), _
        Array(1600, 1700, 1800, 1900, 2000), _
        Array(2100, 2200, 2300, 2400, 2500) _
    )
    
    Debug.Print "interpolateTwoDimensionalArray: TEST 01"
    If math.interpolateTwoDimensionalArray(10#, 1#, keyArray1, keyArray2, valueArray) = 100# Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "interpolateTwoDimensionalArray: TEST 02"
    If math.interpolateTwoDimensionalArray(50#, 5#, keyArray1, keyArray2, valueArray) = 2500# Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "interpolateTwoDimensionalArray: TEST 03"
    If math.interpolateTwoDimensionalArray(50#, 1#, keyArray1, keyArray2, valueArray) = 2100# Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "interpolateTwoDimensionalArray: TEST 04"
    If math.interpolateTwoDimensionalArray(10#, 5#, keyArray1, keyArray2, valueArray) = 500# Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "interpolateTwoDimensionalArray: TEST 05"
    If math.interpolateTwoDimensionalArray(9#, 1#, keyArray1, keyArray2, valueArray) = 100# Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "interpolateTwoDimensionalArray: TEST 06"
    If math.interpolateTwoDimensionalArray(10#, 0#, keyArray1, keyArray2, valueArray) = 100# Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "interpolateTwoDimensionalArray: TEST 07"
    If math.interpolateTwoDimensionalArray(51#, 5#, keyArray1, keyArray2, valueArray) = 2500# Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "interpolateTwoDimensionalArray: TEST 08"
    If math.interpolateTwoDimensionalArray(50#, 6#, keyArray1, keyArray2, valueArray) = 2500# Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "interpolateTwoDimensionalArray: TEST 09"
    If math.interpolateTwoDimensionalArray(15#, 1.5, keyArray1, keyArray2, valueArray) = 400# Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "interpolateTwoDimensionalArray: TEST 10"
    If math.interpolateTwoDimensionalArray(35#, 3#, keyArray1, keyArray2, valueArray) = 1550# Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "interpolateTwoDimensionalArray: TEST 11"
    If math.interpolateTwoDimensionalArray(3.5#, 40#, keyArray1, keyArray2, valueArray) = 500# Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "interpolateTwoDimensionalArray: TEST 12"
    If math.interpolateTwoDimensionalArray(40#, 3.5, keyArray1, keyArray2, valueArray) = 1850# Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If
    

    Set math = Nothing
End Sub

Public Sub testMakeArrayWithStep()
    Dim math
    Set math = New C_Math
    
    Dim array1() As Variant: array1 = Array(1, 2, 3, 4, 5)
    Dim array2() As Variant: array2 = Array(1, 3, 5, 7, 9, 11)
    
    Debug.Print "makeArrayWithStep: TEST 01"
    If math.isArraysSame(math.makeArrayWithStep(1, 5, 1), array1) = True Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "makeArrayWithStep: TEST 02"
    If math.isArraysSame(math.makeArrayWithStep(1, 11, 2), array2) = True Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Set math = Nothing
End Sub

Public Sub testMergeTwoArrays()
    Dim math
    Set math = New C_Math
    
    Dim arr1 As Variant: arr1 = Array(1, 2, 3)
    Dim arr2 As Variant: arr2 = Array(5, 6, 7)
    Dim arr12 As Variant: arr12 = Array(1, 2, 3, 5, 6, 7)

    Dim arr3(1 To 3) As Integer
    arr3(1) = 1 : arr3(2) = 2 : arr3(3) = 3
    Dim arr4(1 To 3) As Integer
    arr4(1) = 5 : arr4(2) = 6 : arr4(3) = 7

    Dim arr34 : arr34 = arr3
    ReDim Preserve arr34(1 To 6)
    arr34(4) = 5 : arr34(5) = 6 : arr34(6) = 7
    
    Debug.Print "makeTestMergeTwoArrays: TEST 01"
    If math.isArraysSame(math.mergeTwoArrays(arr1, arr2), arr12) = True Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "makeTestMergeTwoArrays: TEST 02"
    If math.isArraysSame(math.mergeTwoArrays(arr3, arr4), arr34) = True Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "makeTestMergeTwoArrays: TEST 03"
    If math.isArraysSame(math.mergeTwoArrays(arr1, Array()), arr1) = True Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If

    Debug.Print "makeTestMergeTwoArrays: TEST 04"
    If math.isArraysSame(math.mergeTwoArrays(Array(), arr1), arr1) = True Then
        Debug.Print "PASSED"
    Else
        Debug.Print "FAILED"
    End If


    Set math = Nothing
End Sub

Public Sub testDec2Bin()

    Dim math
    Set math = New C_Math

    Dim count As Integer
    count = 0

    Dim test As New Dictionary

    test.Item(0) = "00000000"
    test.Item(1) = "00000001"
    test.Item(10) = "00001010"

    Dim varKey As variant
    For Each varKey In test.Keys

        if math.dec2bin(CLng(varKey), 8) = test.Item(varKey) Then
            count = count + 1
        Else   
            Debug.Print "Test No." & Str(count + 1) & " - FAILED"

            Debug.Print CStr(varKey)
            Debug.Print test.Item(varKey)
            Debug.Print math.dec2bin(CLng(varKey))
            Exit Sub
        End If
    Next

    Debug.Print Str(count) & " tests PASSED"

    Set math = Nothing

End Sub

Public Sub testBin2Dec()

    Dim math
    Set math = New C_Math

    Dim count As Integer
    count = 0

    

    Dim test As New Dictionary

    test.Item("0") = 0
    test.Item("1") = 1
    test.Item("10101011001101001") = 87657
    test.Item("1010") = 10
    

    Dim varKey As variant
    For Each varKey In test.Keys

        if math.bin2dec(CStr(varKey)) = test.Item(varKey) Then
            count = count + 1
        Else   
            Debug.Print "Test No." & Str(count + 1) & " - FAILED"

            Debug.Print CStr(varKey)
            Debug.Print test.Item(varKey)
            Debug.Print math.bin2dec(CStr(varKey))
            Exit Sub
        End If
    Next

    Debug.Print Str(count) & " tests PASSED"

    Set math = Nothing

End Sub



Public Sub testBytes2Double()

    Dim math
    Set math = New C_Math

    Dim count As Integer
    count = 0

    Dim roundBase As Integer
    roundBase = 5

    Dim test As New Dictionary

    test.Item("17D9CEF753D57440") = 333.333
    test.Item("63affbb7e0365f40") = 124.85746574
    

    Dim varKey As variant
    Dim i As Integer
    Dim b As Byte
    Dim s As String

    For Each varKey In test.Keys

    s = ""
    For i = 1 To 15 STEP 2
        b = "&H" & Mid(Cstr(varKey), i, 2)
        s = s & Chr(b)
    Next i

        if round(math.bytes2double(s, 1#, 0), roundBase) = round(test.Item(varKey), roundBase) Then
            count = count + 1
        Else   
            Debug.Print "Test No." & Str(count + 1) & " - FAILED"

            Debug.Print CStr(varKey)
            Debug.Print test.Item(varKey)
            Debug.Print math.bytes2double(s, 1#, 0)
            Exit Sub
        End If
    Next

    Debug.Print Str(count) & " tests PASSED"

    Set math = Nothing

End Sub
    
Public Sub test()

    Debug.Print "TEST InterpolateOneDimensionalArray"
    Call testInterpolateOneDimensionalArray()
    Debug.Print "----------------------"

    Debug.Print "TEST FindIdexOfFirstElementNotLessThan"
    Call testFindIdexOfFirstElementNotLessThan()
    Debug.Print "----------------------"

    Debug.Print "TEST IsArraysSame"
    Call  testIsArraysSame()
    Debug.Print "----------------------"

    Debug.Print "TEST InterpolateTwoDimensionalArray"
    Call  testInterpolateTwoDimensionalArray()
    Debug.Print "----------------------"

    Debug.Print "TEST MakeArrayWithStep"
    Call  testMakeArrayWithStep()
    Debug.Print "----------------------"

    Debug.Print "TEST MergeTwoArrays"
    Call  testMergeTwoArrays()
    Debug.Print "----------------------"

    Debug.Print "TEST Dec2Bin"
    Call  testDec2Bin()
    Debug.Print "----------------------"

    Debug.Print "TEST Bin2Dec"
    Call  testBin2Dec()
    Debug.Print "----------------------"

    Debug.Print "TEST Bytes2Double"
    Call  testBytes2Double()
    Debug.Print "----------------------"

    Debug.Print "TEST FindIdexOfValueInArray"
    Call testFindIdexOfValueInArray
    Debug.Print "----------------------"

End Sub