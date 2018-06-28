VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Excel_Worksheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const VERSION = "1.2.0"

Public Function findColumnWithText(sheet As Excel.Worksheet, row As Integer, text As String) As Integer

    Dim Rng As Excel.Range
    Set Rng = sheet.Rows(row).Find(what:=text, LookIn:=xlValues, lookAt:=xlWhole, MatchCase:=True)
    
    If Not Rng Is Nothing Then
        findColumnWithText = Rng.column
    Else
        findColumnWithText = 0
    End If
    
End Function

Public Function findRowWithText(sheet As Excel.Worksheet, column As Integer, text As String) As Integer

    Dim Rng As Excel.Range
    Set Rng = sheet.Columns(column).Find(what:=text, LookIn:=xlValues, lookAt:=xlWhole, MatchCase:=True)
    
    If Not Rng Is Nothing Then
        findRowWithText = Rng.row
    Else
        findRowWithText = 0
    End If
    
End Function

' ������� ��������� Collection �������� �� ����������� ��������/����� �� ����� Microsoft Excel
' (Excel.Worksheet) sheet - ������ �����
' (integer) rowWithNames - ����� ������ � ������� ��������
' (variant) names - ������ (string) � ������� ����������� ��������. ������ ������� �������
'                   ��� ������� � �������. ���� ����� �� �����, �� ������ ���� ������ ""
'                   ���� �� �� ������, �� ����� � ������� ������ ���� �����������, �����
'                   ���������� ������ ����� �������� ���������� � ����������� �������.
'                   ������ ������ ������ �� ������ ������ � ������� � ������ ������
'                   ������������� ���-�� ��������� � names >= 2 
' (variant) newNames - ������ (string) � ������ ������� ������ � ����������� �������� �� ������,
'                   ���� ��� �� ������ ��������� � ��������� ������� ��������
'                   LBound, UBound names � newNames ������ ���������

Public Function getCollectionOfRows(sheet As Excel.Worksheet, rowWithNames As Integer, _
                 names As Variant, Optional newNames As Variant = NULL) As C_ImprovedCollection

    If (UBound(names) - LBound(names)) = 0 Then Exit Function 

    Dim columns() As Integer
    ReDim columns(LBound(names) To UBound(names))

    Dim i As Long
    For i=LBound(names) To UBound(names)
        columns(i) = findColumnWithText(sheet, rowWithNames, CStr(names(i)))
    Next i

    Dim key As String
    Dim result As New C_ImprovedCollection
    Dim dic As Dictionary
    Dim row As Long
    Dim name As String
    
    row = rowWithNames + 1
    
    Do While sheet.Cells(row, columns(LBound(columns) + 1)).Value <> ""

        Set dic = New Dictionary

        For i = LBound(names) + 1 To UBound(names)
            if IsNull(newNames) Then name = names(i) Else name = newNames(i)
            dic.Item(name) = sheet.Cells(row, columns(i)).Value
        Next i

        If names(LBound(names)) = "" Then
            result.Add dic
        Else
            key = CStr(sheet.Cells(row, columns(LBound(columns))).Value)
            if result.Exists(key) Then  result.Remove(key)
            result.Add dic, key
        End If

        
        Set dic = Nothing

        row = row + 1

    Loop

    Set getCollectionOfRows = result
    Set result = Nothing

End Function

' ������� ��������� ����������� �������/������ �� �������� � Collection �� ����� Microsoft Excel
' (Excel.Worksheet) sheet - ������ �����
' (integer) rowWithNames - ����� ������ � ������� ��������
' (C_ImprovedCollection) col - �������� ��������� ��������
' (variant) namesInSheet - ������ (string) � ������� ����������� ��������, ������� ����� ����� �� �������
' (variant) namesInCollection - ������ (string) � ������� ������ � ������� � ������, ���� ��� �� ���������
'                               � ������� �� �����
' LBound, UBound namesInSheet � namesInCollection ������ ���������

Public Sub setCollectionOfRows(sheet As Excel.Worksheet, rowWithNames As Integer, col As C_ImprovedCollection, _
                 namesInSheet As Variant, Optional namesInCollection As Variant = NULL) 

    ' ���� ������ �������� �� �����, ��������������� namesInSheet
    Dim columnIndexInSheetArray() As Integer
    ReDim Preserve columnIndexInSheetArray(LBound(namesInSheet) To UBound(namesInSheet))

    Dim i As Long
    Dim columnIndex As Integer
    For i = LBound(namesInSheet) To UBound(namesInSheet)

        columnIndex = findColumnWithText(sheet, rowWithNames, CStr(namesInSheet(i)))
        
        If columnIndex = 0 Then
            MsgBox "������� � ������ " & namesInSheet(i) & " �� ������ !!!"
            Exit Sub
        End If

        columnIndexInSheetArray(i) = columnIndex
    Next i

    Dim names As Variant
    If IsNull(namesInCollection) Then
        names = namesInSheet
    Else
        names = namesInCollection
    End If

    Dim dic As Dictionary
    Dim k As Integer
    For i = 1 To col.Count

        Set dic = col(i)

        For k = LBound(names) To UBound(names)
            sheet.Cells(rowWithNames + i, columnIndexInSheetArray(k)) = dic.Item(names(k))
        Next k

    Next i

End Sub