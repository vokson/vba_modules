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

CONST VERSION = "1.3.0"

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

Function max(a As Double, b As Double) As Double
    If (a > b) Then
        max = a
    Else
        max = b
    End If
End Function

Function min(a As Double, b As Double) As Double
    If (a < b) Then
        min = a
    Else
        min = b
    End If
End Function

'Decimal To Binary
' =================

' NOTE: You can limit the size of the returned
'              answer by specifying the number of bits
' Work only with POSITIVE numbers
Function dec2bin(ByVal DecimalIn As Variant, _
              Optional NumberOfBits As Variant) As String
    dec2bin = ""
    DecimalIn = Int(CDec(DecimalIn))

    If DecimalIn = 0 Or DecimalIn = -0 Then
        dec2bin = "0"
        GoTO Handler
    End If

    Do While DecimalIn <> 0
        dec2bin = Format$(DecimalIn - 2 * Int(DecimalIn / 2)) & dec2bin
        DecimalIn = Int(DecimalIn / 2)
    Loop

Handler:

    If Not IsMissing(NumberOfBits) Then
       If Len(dec2bin) > NumberOfBits Then
          dec2bin = "Error - Number exceeds specified bit size"
       Else
          dec2bin = Right$(String$(NumberOfBits, "0") & dec2bin, NumberOfBits)
       End If
    End If
End Function
 
'Binary To Decimal
' =================
Function bin2dec(BinaryString As String) As Variant
    Dim X As Integer
    For X = 0 To Len(BinaryString) - 1
        bin2dec = CDec(bin2dec) + Val(Mid(BinaryString, _
                  Len(BinaryString) - X, 1)) * 2 ^ X
    Next
End Function

' Не протестировано !!!
Public Function hex2ascii(st As String) As String
    Dim com As String
    Dim x As Long
    Dim k As String

    For x = 1 To Len(st) - 1 Step 2
        k = Mid(st, x, 2)
        com = com & ChrW(CInt("&H" & k))
    Next

    hex2ascii = com
End Function

' Function decodes 8 bytes string into Double number
Function bytes2double(bytes As String, MNO As Double, STEP As Integer) As Double

    Dim i As Integer

    ' Переворачиваем строку
    Dim number As String
    number = StrReverse(bytes)

    ' Создаем массив в двоичном формате
    Dim mas(1 To 8) As String
    For i = 1 To 8
        mas(i) = dec2bin(Asc(Mid(number, i, 1)), 8)
    Next i

    ' Cоздаем 64 битную строку
    Dim s As String
    s = Join(mas, "")

    ' Разделяем на биты знака, экспоненты и мантиссы
    Dim znak As String
    Dim exponent As String
    Dim mant As String
    znak = Mid(s, 1, 1)
    exponent = Mid(s, 2, 11)
    mant = Mid(s, 13, 52)

    ' Находим множители
    Dim a As Double, b As Double, c As Double
    a = (-1) ^ bin2dec(znak)
    b = 2 ^ (bin2dec(exponent) - STEP - 1023)
    c = 1 + bin2dec(mant) / (2 ^ 52)

    ' Считаем результат
     bytes2double = a * b * c * MNO

End Function
