VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Math"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function pi() As Double
    pi = 3.14159265358979
End Function

' Интерполяция по двум одномерным сортированным массивам одинаковой длины
' valueArray - массив с исходными величинами, сортированный по возрастанию
' value - величина, для которой необходимо найти
' resultArray - массив значений
Function interpolateOneDimensionalArray(value, valueArray, resultArray) As Variant
    Dim k As Long
    
    If IsArray(valueArray) And IsArray(resultArray) Then
    
        interpolateOneDimensionalArray = resultArray(UBound(resultArray))
        
        For k = LBound(valueArray) To UBound(valueArray)
            If (valueArray(k) >= value) Then
            
                If k = LBound(valueArray) Then
                    interpolateOneDimensionalArray = resultArray(LBound(resultArray))
                Else
                    interpolateOneDimensionalArray = resultArray(k - 1) + _
                        (value - valueArray(k - 1)) / (valueArray(k) - valueArray(k - 1)) * (resultArray(k) - resultArray(k - 1))
                End If
                
                Exit For
            End If
        Next k
    
    End If
    
End Function

' Нахождение индекса элемента, значение которого >= value
' valueArray - массив с исходными величинами, сортированный по возрастанию
' value - величина, для которой необходимо найти
Function findIdexOfFirstElementNotLessThan(value, sortedArray) As Integer
    
    If Not IsArray(sortedArray) Then Exit Function
    
    Dim k As Long
    For k = LBound(sortedArray) To UBound(sortedArray)
        If (sortedArray(k) >= value) Then
            findIdexOfFirstElementNotLessThan = k
            Exit Function
        End If
    Next k
    
End Function

' Определяет равны ли два массива друг другу
Public Function isArraysSame(array1, array2) As Boolean
    isArraysSame = True

    Dim i As Integer
    For i = LBound(array1) to UBound(array1)
        If array1(i) <> array2(i) Then
            isArraysSame = False
            Exit Function
        End If
    Next i

End Function

' Двумерная интерполяция
' key1, key2 - ключи в массивах keyArray1, keyArray2
' keyArray1, keyArray2 - одномерные массив ключей, сортированные по возрастанию
' valueArray - двумерный массив величин
' Size(keyArray1, 0) = Size(valueArray, 0)
' Size(keyArray2, 0) = Size(valueArray, 1)
Function interpolateTwoDimensionalArray(key1, key2, keyArray1, keyArray2, valueArray) As Variant
    Dim i As Long
    Dim newArray
    ' Dim lengthOfNewArray As Long : lengthOfNewArray = 0
    
    If IsArray(keyArray1) And IsArray(keyArray2) And IsArray(valueArray) Then

        ReDim newArray(LBound(keyArray1) To UBound(keyArray1))

        For i = LBound(keyArray1) To UBound(keyArray1)
            newArray(i) = interpolateOneDimensionalArray(key2, keyArray2, valueArray(i))
        Next i

        interpolateTwoDimensionalArray = interpolateOneDimensionalArray(key1, keyArray1, newArray)
    
    End If
    
End Function

' Создает массив чисел от min до max с шагом pitch
Function makeArrayWithStep(min As Long, max As Long, pitch As Long) As Long()
    Dim result() As Long
    Dim lengthOfArray As Long
    lengthOfArray = 0

    Dim i As Long
    For i = min To max Step pitch
        Redim Preserve result(lengthOfArray)
        result(UBound(result)) = i
        lengthOfArray = lengthOfArray + 1
    Next i

    makeArrayWithStep = result
End Function

Function mergeTwoArrays(array1, array2) As Variant
    Dim result : result = array1
    Redim Preserve result (LBound(array1) To (UBound(result) + UBound(array2) - LBound(array2) + 1))

    Dim i As Long
    Dim k As Long
    k = LBound(array2)
    For i = UBound(array1) + 1 To Ubound(result)
        result(i) = array2(k)
        k = k + 1
    Next i

    mergeTwoArrays = result
End Function

Function sum(a As Double, b As Double) As Double
    sum = a + b
End Function