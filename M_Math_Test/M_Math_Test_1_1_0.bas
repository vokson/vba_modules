Attribute VB_Name = "M_Math_Test"
Option Explicit

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

Public Sub testIsArraysSame()
    Dim math
    Set math = New C_Math
    
    Dim arr1() As Variant: arr1 = Array(10, 20, 30, 40, 50)
    Dim arr2() As Variant: arr2 = Array(10#, 20#, 30#, 40#, 50#)
    Dim arr3() As Variant: arr3 = Array(10.1, 20#, 30#, 40#, 50#)
    Dim arr4() As Variant: arr4 = Array("10", 20#, 30#, 40#, 50#)
    Dim arr5() As Variant: arr5 = Array(True, 20#, 30#, 40#, 50#)
    
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
    If math.isArraysSame(arr1, arr4) = False Then
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