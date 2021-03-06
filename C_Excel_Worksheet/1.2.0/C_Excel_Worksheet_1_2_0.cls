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

' ������� ��������� ����������� �������/������ �� �������� � Collection �� ����� Microsoft Excel, ��������� Range
' ����� ��������� �� ������ ���� ���������
'
' (Excel.Worksheet) sheet - ������ �����
' (integer) rowWithNames - ����� ������ � ������� ��������
' (C_ImprovedCollection) col - �������� ��������� ��������
' (boolean) isLastItemArray - �������� �� ��������� ������� ��������. ���� �� ��, ������� ����� ��� ��������
'                               ������� � ������ ������
' (variant) namesInSheet - ������ (string) � ������� ����������� ��������, ������� ����� ����� �� �������
' (variant) namesInCollection - ������ (string) � ������� ������ � ������� � ������, ���� ��� �� ���������
'                               � ������� �� �����
' LBound, UBound namesInSheet � namesInCollection ������ ���������

Public Sub setCollectionOfRowsUsingRange(sheet As Excel.Worksheet, rowWithNames As Integer, col As C_ImprovedCollection, _
                 isLastItemArray As Boolean, namesInSheet As Variant, Optional namesInCollection As Variant = NULL) 

    Dim oldStatusBar As Boolean
    oldStatusBar = Application.DisplayStatusBar 
    Application.DisplayStatusBar = True 
    Application.ScreenUpdating = False

    Dim StartTime As Double
    StartTime = Timer

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
    Dim j As Long
    Dim c As Long
    Dim minIndex As Integer
    Dim valuesToBeWritten() As Variant
    Dim minCell As String
    Dim countOfColumns As Integer
    Dim countInArray As Integer, maxCountInArray As Integer
    Dim dicArr As Variant

    ' ������������ ���-�� ��������� � ������
    countOfColumns = UBound(names)-LBound(names) + 1

    If isLastItemArray = True Then

        maxCountInArray = 0

        For i = 1 To col.Count
            countInArray = UBound(col(i).Item(names(UBound(names))))-LBound(col(i).Item(names(UBound(names)))) +1

            If countInArray > maxCountInArray Then
                maxCountInArray = countInArray
            End If

        Next i

        countOfColumns = countOfColumns + maxCountInArray - 1

    End If

    ' �������� ������ ���������� �������
    ReDim valuesToBeWritten(1 To col.Count, 1 To countOfColumns)

    ' ���������� ����� �������
    minIndex = Application.Min(columnIndexInSheetArray)
    
    ' ����� ��������� ������
    For i = 1 To col.Count

        Set dic = col(i)

        Application.StatusBar = "���������� ����� - " & CStr(Round(i/col.Count * 100)) & "%"

        c = 1

        For k = LBound(names) To UBound(names) - 1
            valuesToBeWritten(i, c) = dic.item(names(k))
            c = c + 1
        Next k

        If isLastItemArray = False Then
            valuesToBeWritten(i, c) = dic.item(names(k))
        Else
        
            dicArr = dic.item(names(k))

            For j = LBound(dicArr) To UBound(dicArr)
                valuesToBeWritten(i, c) = dicArr(j)
                c = c + 1
            Next j

        End If
        
        Set dic = Nothing

    Next i

    ' ����� ������ � Range
    minCell = sheet.Cells(rowWithNames + 1, minIndex).Address(False, False)
    sheet.Range(minCell).Resize(UBound(valuesToBeWritten, 1), UBound(valuesToBeWritten, 2)).Value  = valuesToBeWritten

    Debug.Print "setCollectionOfRowsUsingRange - " & Round(Timer - StartTime, 2) & " seconds", vbInformation

    Application.StatusBar = False
    Application.DisplayStatusBar = oldStatusBar
    Application.ScreenUpdating = True

End Sub


' ������� ��������� ����������� �������/������ �� �������� � Collection �� ����� Microsoft Excel, ��������� Cells
' (Excel.Worksheet) sheet - ������ �����
' (integer) rowWithNames - ����� ������ � ������� ��������
' (C_ImprovedCollection) col - �������� ��������� ��������
' (boolean) isLastItemArray - �������� �� ��������� ������� ��������. ���� �� ��, ������� ����� ��� ��������
'                               ������� � ������ ������
' (variant) namesInSheet - ������ (string) � ������� ����������� ��������, ������� ����� ����� �� �������
' (variant) namesInCollection - ������ (string) � ������� ������ � ������� � ������, ���� ��� �� ���������
'                               � ������� �� �����
' LBound, UBound namesInSheet � namesInCollection ������ ���������

Public Sub setCollectionOfRowsUsingCells(sheet As Excel.Worksheet, rowWithNames As Integer, col As C_ImprovedCollection, _
                 isLastItemArray As Boolean, namesInSheet As Variant, Optional namesInCollection As Variant = NULL) 

    Dim oldStatusBar As Boolean
    oldStatusBar = Application.DisplayStatusBar 
    Application.DisplayStatusBar = True 
    Application.ScreenUpdating = False

    Dim StartTime As Double
    StartTime = Timer

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
    Dim j As Long
    Dim c As Long

    For i = 1 To col.Count

        Application.StatusBar = "���������� ����� - " & CStr(Round(i/col.Count * 100)) & "%"

        Set dic = col(i)

        For k = LBound(names) To UBound(names)-1
            sheet.Cells(rowWithNames + i, columnIndexInSheetArray(k)) = dic.Item(names(k))
        Next k

        If isLastItemArray = False Then
            sheet.Cells(rowWithNames + i, columnIndexInSheetArray(k)) = dic.Item(names(k))
        Else

            c = 0
            For j = LBound(dic.Item(names(k))) To UBound(dic.Item(names(k)))
                sheet.Cells(rowWithNames + i, columnIndexInSheetArray(k ) + c) = dic.Item(names(k))(j)
                c = c + 1
            Next j

        End If
    Next i

    Debug.Print "setCollectionOfRowsUsingCells - " & Round(Timer - StartTime, 2) & " seconds", vbInformation

    Application.StatusBar = False
    Application.DisplayStatusBar = oldStatusBar
    Application.ScreenUpdating = True

End Sub