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

Const VERSION = "1.3.0"

Public Function findColumnWithText(sheet As Excel.Worksheet, row As Long, text As String) As Long

    Dim Rng As Excel.Range
    Set Rng = sheet.Rows(row).Find(what:=text, LookIn:=xlValues, lookAt:=xlWhole, MatchCase:=True)
    
    If Not Rng Is Nothing Then
        findColumnWithText = Rng.column
    Else
        findColumnWithText = 0
    End If
    
End Function

Public Function findRowWithText(sheet As Excel.Worksheet, column As Long, text As String) As Long

    Dim Rng As Excel.Range
    Set Rng = sheet.Columns(column).Find(what:=text, LookIn:=xlValues, lookAt:=xlWhole, MatchCase:=True)
    
    If Not Rng Is Nothing Then
        findRowWithText = Rng.row
    Else
        findRowWithText = 0
    End If
    
End Function

' Функция формирует массив индексов столцов с необходимыми имена в заголовке таблицы
' (Excel.Worksheet) sheet - Объект листа
' (Long) rowWithNames - номер строки с именами столбцов
' (variant) names - Массив (string) с именами необходимых столбцов. 
Private Function getColumnIndexes( _
    sheet As Excel.Worksheet, _
    rowWithNames As Long, _
    names As Variant _
) As Long()

    Dim columns() As Long
    ReDim columns(LBound(names) To UBound(names))

    Dim i As Long
    For i=LBound(names) To UBound(names)
        columns(i) = findColumnWithText(sheet, rowWithNames, CStr(names(i)))
    Next i

    getColumnIndexes = columns

End Function

' Функция формирует словарь из необходимых столбцов на листе Microsoft Excel для определенной строки
' используя индексы необходимых столбцов
' (Excel.Worksheet) sheet - Объект листа
' (Long) row - номер строки с данными
' (Long) columns - массив номеров столцов
' (Variant) names -  массив (String) с именами столбцов 
'      LBound, UBound names и columns должны совпадать

Private Function getRowUsingIndexes( _
    sheet As Excel.Worksheet, _
    row As Long, _
    columns() As Long, _
    names As Variant _
) As Dictionary

    Dim dic As New Dictionary
    Dim i As Long

    For i = LBound(names) To UBound(names)
        dic.Item(names(i)) = sheet.Cells(row, columns(i)).Value
    Next i

    Set getRowUsingIndexes = dic
    Set dic = Nothing

End Function

' Функция формирует словарь из необходимых столбцов на листе Microsoft Excel для определенной строки
' (Excel.Worksheet) sheet - Объект листа
' (Long) rowWithNames - номер строки с именами столбцов
' (Long) rowWithData - номер строки с данными
' (variant) names - Массив (string) с именами необходимых столбцов. 
' (variant) newNames - Массив (string) с новыми именами ключей в создаваемых словарях на случай,
'                   если они не должны совпадать с исходными именами столбцов
'                   LBound, UBound names и newNames должны совпадать

Public Function getRow(sheet As Excel.Worksheet, rowWithNames As Long, rowWithData As Long, _
                 names As Variant, Optional newNames As Variant = NULL) As Dictionary

    Dim columns() As Long
    columns = getColumnIndexes(sheet, rowWithNames, names)

    if IsNull(newNames) Then
        Set getRow = getRowUsingIndexes(sheet, rowWithData, columns, names)
    Else
        Set getRow = getRowUsingIndexes(sheet, rowWithData, columns, newNames)
    End If

End Function


' Функция формирует Collection словарей из необходимых столбцов/строк на листе Microsoft Excel
' (Excel.Worksheet) sheet - Объект листа
' (integer) rowWithNames - номер строки с именами столбцов
' (variant) names - Массив (string) с именами необходимых столбцов. Первый элемент массива
'                   имя столбца с ключами. Если ключи не нужны, он должен быть пустым ""
'                   Если он не пустой, то ключи в столбце должны быть уникальными, иначе
'                   последующе строки будут заменять предыдущие с одинаковыми ключами.
'                   Скрипт читает строки до пустой клетки в столбце с вторым ключом
'                   Следовательно кол-во элементов в names >= 2 
' (variant) newNames - Массив (string) с новыми именами ключей в создаваемых словарях на случай,
'                   если они не должны совпадать с исходными именами столбцов
'                   LBound, UBound names и newNames должны совпадать

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

' Функция заполняет необходимые столбцы/строки из словарей в Collection на листе Microsoft Excel, используя Range
' Между столбцами не должно быть пропусков
'
' (Excel.Worksheet) sheet - Объект листа
' (Long) rowWithNames - номер строки с именами столбцов
' (C_ImprovedCollection) col - исходная коллекция словарей
' (boolean) isLastItemArray - является ли последний элемент массивом. Если да то, функция пишет все элементы
'                               массива в ячейки правее
' (variant) namesInSheet - Массив (string) с именами необходимых столбцов, которые будут взяты из словаря
' (variant) namesInCollection - Массив (string) с именами ключей в словаре в случае, если они не совпадают
'                               с ключами на листе
' LBound, UBound namesInSheet и namesInCollection должны совпадать

Public Sub setCollectionOfRowsUsingRange(sheet As Excel.Worksheet, rowWithNames As Long, col As C_ImprovedCollection, _
                 isLastItemArray As Boolean, namesInSheet As Variant, Optional namesInCollection As Variant = NULL) 

    Dim oldStatusBar As Boolean
    oldStatusBar = Application.DisplayStatusBar 
    Application.DisplayStatusBar = True 
    Application.ScreenUpdating = False

    Dim StartTime As Double
    StartTime = Timer

    ' Ищем номера столбцов на листе, соответствующие namesInSheet
    Dim columnIndexInSheetArray() As Long
    ReDim Preserve columnIndexInSheetArray(LBound(namesInSheet) To UBound(namesInSheet))

    Dim i As Long
    Dim columnIndex As Long
    For i = LBound(namesInSheet) To UBound(namesInSheet)

        columnIndex = findColumnWithText(sheet, rowWithNames, CStr(namesInSheet(i)))
        
        If columnIndex = 0 Then
            MsgBox "Столбец с именем " & namesInSheet(i) & " не найден !!!"
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
    Dim k As Long
    Dim j As Long
    Dim c As Long
    Dim minIndex As Long
    Dim valuesToBeWritten() As Variant
    Dim minCell As String
    Dim countOfColumns As Long
    Dim countInArray As Long, maxCountInArray As Long
    Dim dicArr As Variant

    ' Рассчитываем кол-во элементов в строке
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

    ' Изменяем размер двумерного массива
    ReDim valuesToBeWritten(1 To col.Count, 1 To countOfColumns)

    ' Определяем левый столбец
    minIndex = Application.Min(columnIndexInSheetArray)
    
    ' Пишем двумерный массив
    For i = 1 To col.Count

        Set dic = col(i)

        Application.StatusBar = "Заполнение строк - " & CStr(Round(i/col.Count * 100)) & "%"

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

    ' Пишем массив в Range
    minCell = sheet.Cells(rowWithNames + 1, minIndex).Address(False, False)
    sheet.Range(minCell).Resize(UBound(valuesToBeWritten, 1), UBound(valuesToBeWritten, 2)).Value  = valuesToBeWritten

    ' Debug.Print "setCollectionOfRowsUsingRange - " & Round(Timer - StartTime, 2) & " seconds", vbInformation

    Application.StatusBar = False
    Application.DisplayStatusBar = oldStatusBar
    Application.ScreenUpdating = True

End Sub


' Функция заполняет необходимые столбцы/строки из словарей в Collection на листе Microsoft Excel, используя Cells
' (Excel.Worksheet) sheet - Объект листа
' (Long) rowWithNames - номер строки с именами столбцов
' (C_ImprovedCollection) col - исходная коллекция словарей
' (boolean) isLastItemArray - является ли последний элемент массивом. Если да то, функция пишет все элементы
'                               массива в ячейки правее
' (variant) namesInSheet - Массив (string) с именами необходимых столбцов, которые будут взяты из словаря
' (variant) namesInCollection - Массив (string) с именами ключей в словаре в случае, если они не совпадают
'                               с ключами на листе
' LBound, UBound namesInSheet и namesInCollection должны совпадать
' (Long) rowStartFrom - номер строки, с которой нужно начинать запись на лист 

Public Sub setCollectionOfRowsUsingCells(sheet As Excel.Worksheet, rowWithNames As Long, col As C_ImprovedCollection, _
                isLastItemArray As Boolean, namesInSheet As Variant, Optional namesInCollection As Variant = NULL, _
                Optional rowStartFrom As Long = 0) 

    On Error Goto Handler

    Dim oldStatusBar As Boolean
    oldStatusBar = Application.DisplayStatusBar 
    Application.DisplayStatusBar = True 
    Application.ScreenUpdating = False

    Dim StartTime As Double
    StartTime = Timer

    ' Ищем номера столбцов на листе, соответствующие namesInSheet
    Dim columnIndexInSheetArray() As Long
    ReDim Preserve columnIndexInSheetArray(LBound(namesInSheet) To UBound(namesInSheet))

    Dim i As Long
    Dim columnIndex As Long
    For i = LBound(namesInSheet) To UBound(namesInSheet)

        columnIndex = findColumnWithText(sheet, rowWithNames, CStr(namesInSheet(i)))
        
        If columnIndex = 0 Then
            MsgBox "Столбец с именем " & namesInSheet(i) & " не найден !!!"
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
    Dim k As Long
    Dim j As Long
    Dim c As Long

    Dim row As Long
    if rowStartFrom > 0 Then
        row = rowStartFrom - 1 
    Else
        row = rowWithNames
    End If

    For i = 1 To col.Count

        Application.StatusBar = "Заполнение строк - " & CStr(Round(i/col.Count * 100)) & "%"

        Set dic = col(i)

        For k = LBound(names) To UBound(names)-1
            sheet.Cells(row + i, columnIndexInSheetArray(k)) = dic.Item(names(k))
        Next k

        If isLastItemArray = False Then
            sheet.Cells(row + i, columnIndexInSheetArray(k)) = dic.Item(names(k))
        Else

            c = 0
            For j = LBound(dic.Item(names(k))) To UBound(dic.Item(names(k)))
                sheet.Cells(row + i, columnIndexInSheetArray(k ) + c) = dic.Item(names(k))(j)
                c = c + 1
            Next j

        End If
    Next i

    ' Debug.Print "setCollectionOfRowsUsingCells - " & Round(Timer - StartTime, 2) & " seconds", vbInformation

    Application.StatusBar = False
    Application.DisplayStatusBar = oldStatusBar
    Application.ScreenUpdating = True
    Exit Sub

    Handler:
    Application.ScreenUpdating = True
    MsgBox "Error during writing to Excel Worksheet '" & sheet.name & "'"

End Sub

' Функция проверяет пустые ли столбцы/строки
' (Excel.Worksheet) sheet - Объект листа
' (Long) rowWithNames - номер строки с именами столбцов
' (variant) namesInSheet - Массив (string) с именами необходимых столбцов
' (Long) rowStartFrom - номер строки, с которой нужно начинать проверку
' (Long) countOfRows - кол-во строк, которое необходимо проверить

Public Function checkIfRowsEmpty( _
    sheet As Excel.Worksheet, _
    rowWithNames As Long, _
    namesInSheet As Variant, _
    rowStartFrom As Long, _
    countOfRows As Long _
) As Boolean

    On Error Goto Handler

    ' Ищем номера столбцов на листе, соответствующие namesInSheet
    Dim columnIndexInSheetArray() As Long
    ReDim Preserve columnIndexInSheetArray(LBound(namesInSheet) To UBound(namesInSheet))

    Dim i As Long
    Dim k As Long
    Dim columnIndex As Long

    For i = LBound(namesInSheet) To UBound(namesInSheet)

        columnIndex = findColumnWithText(sheet, rowWithNames, CStr(namesInSheet(i)))
        
        If columnIndex = 0 Then
            MsgBox "Столбец с именем " & namesInSheet(i) & " не найден !!!"
            GoTo Handler
        End If

        columnIndexInSheetArray(i) = columnIndex
    Next i

    For i = rowStartFrom To rowStartFrom + countOfRows
        For k = LBound(namesInSheet) To UBound(namesInSheet)
            if Not IsEmpty(sheet.Cells(i, columnIndexInSheetArray(k))) Then
                Goto Handler
            End If
        Next k
    Next i

    checkIfRowsEmpty = True
    Exit Function

    Handler:
        checkIfRowsEmpty = False

End Function

' Формируем словарь ключ-значение из листа
' (Excel.Worksheet) sheet - Объект листа
' (Long) rowWithNames - номер строки с именами столбцов
' (Long) colWithKeys - номер колонки с ключами словаря
' (Long) colWithValues - номер колонки со значениями словаря
Public Function getKeyValueDictionary(sheet As Excel.Worksheet, rowWithNames As Long, _ 
                colWithKeys As Long, colWithValues As Long) As Dictionary
    
    Dim name As String
    Dim value As String
    Dim row As Long 
    
    Set getKeyValueDictionary = New Dictionary
    row = rowWithNames + 1

    Do Until IsEmpty(sheet.Cells(row, colWithKeys))
        name = CStr(sheet.Cells(row, colWithKeys).value)
        value = CStr(sheet.Cells(row, colWithValues).value)
        getKeyValueDictionary.Item(name) = value
        
        row = row + 1
    Loop

End Function