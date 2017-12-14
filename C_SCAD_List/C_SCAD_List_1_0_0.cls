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

' ������������ �� ���� ���������� ������������� �������� ���������� �����
' valueArray - ������ � ��������� ����������, ������������� �� �����������
' value - ��������, ��� ������� ���������� �����
' resultArray - ������ ��������
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

' ���������� ������� ��������, �������� �������� >= value
' valueArray - ������ � ��������� ����������, ������������� �� �����������
' value - ��������, ��� ������� ���������� �����
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

' ���������� ����� �� ��� ������� ���� �����
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