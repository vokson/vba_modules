VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_File"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

CONST VERSION = "1.1.0"

Private fso As New FileSystemObject

Public Function getExtensionOfFile(fileName As String) As String
   getExtensionOfFile = Right(fileName, Len(fileName) - InStrRev(fileName, "."))
End Function

Public Function getFileNameFromFullPath(path As String) As String
   getFileNameFromFullPath = Right(path, Len(path) - InStrRev(path, DIRECTORY_SEPARATOR))
End Function

Public Sub writeFileInUnicode(fullPath As String, text As String)
    Dim fs, a
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(fullPath, True, True) ' Second True is Unicode format
    a.Write (text)
    a.Close
End Sub

' Проверяем является ли длина пути возможной для сохранения в Windows
' fullPath - полный путь к файлу, включая имя файла
Public Function checkLengthOfFullPath(fullPath As String, Optional limit As Integer = 255) As Boolean
    If Len(fullPath) > limit Then
        checkLengthOfFullPath = False
    Else
        checkLengthOfFullPath = True
    End If
End Function

' Урезает полный путь до необходимой длины, укорачивая имя файла
' Если укорачивание возможно, возвращает новый полный путь
' Если нет, возвращает пустую строку
' minLengthOfBaseName - мин кол-во символов, до которого можно укорачивать файл. Например, "abcde.doc"
Public Function cutFullPath(fullPath As String, Optional minLengthOfBaseName As Integer = 10, Optional limit As Integer = 255) As String
    ' Если полный путь укладывается в limit, то не изменяем его
    if Len(fullPath) <= limit Then
        cutFullPath = fullPath
        Exit Function
    End If

    ' Если не укладывается, разбираем на составляющие
    Dim baseName As String
    baseName = fso.GetBaseName(fullPath)

    Dim extensionName As String
    extensionName = fso.GetExtensionName(fullPath)

    Dim dotLength As Integer
    dotLength = 1
    If extensionName = "" Then
        dotLength = 0
    End If

    Dim pathWithoutFilename As String
    pathWithoutFilename = left(fullPath, Len(fullPath) - Len(baseName) - Len(extensionName) - dotLength)

    ' Если длина директории слишком длинная, чтобы обрезанный путь уместился в limit
    If Len(pathWithoutFilename) + Len(extensionName) + minLengthOfBaseName + dotLength > limit Then
        cutFullPath = ""
        Exit Function
    End If

    ' Обрезаем путь
    Dim shortBaseName As String
    shortBaseName = left(baseName, limit - dotLength - Len(extensionName) - Len(pathWithoutFilename))

    ' Формируем новый укороченный полный путь
    cutFullPath = pathWithoutFilename & shortBaseName
    If Not extensionName = "" Then
        cutFullPath = cutFullPath & "." & extensionName
    End If
    
End Function

Public Sub writeFileInUtf8(fullPath As String, cText As String)
    On Error GoTo errHandler
    Dim fsT As Object

    'Create Stream object
    Set fsT = CreateObject("ADODB.Stream")

    'Specify stream type - we want To save text/string data.
    fsT.Type = 2

    'Specify charset For the source text data.
    fsT.Charset = "utf-8"

    'Open the stream And write binary data To the object
    fsT.Open
    fsT.writetext cText

    'Save binary data To disk
    fsT.SaveToFile fullPath, 2

    GoTo finish

errHandler:
    MsgBox (Err.Description)
    Exit Sub

finish:
End Sub

Public Function readFile(fullPath As String) As String
    Dim fso As New FileSystemObject
    Dim JsonTS As TextStream
    
    Set JsonTS = fso.OpenTextFile(fullPath, ForReading)
        readFile = JsonTS.ReadAll
    JsonTS.Close
End Function

Public Sub copyFile(source As String, destination As String)
    Dim fso As Object
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    
    fso.copyFile source, destination
End Sub

Public Function isFileExists(fullPath As String) As Boolean
    If Dir(fullPath) <> "" Then
        isFileExists = True
    Else
        isFileExists = False
    End If
End Function

Public Function replaceRestrictedSymbols(nameOfFile As String, replacement As String) As String
    Const restrictedSymbols = "\|/:*?<>+%!@"
    
    Dim result As String
    result = nameOfFile
    
    Dim i As Long
    For i = 1 To Len(restrictedSymbols)
        result = Replace(result, Mid(restrictedSymbols, i, 1), replacement)
    Next i
    
    result = Replace(result, """", replacement) ' Replace " with _

    replaceRestrictedSymbols = result
End Function

' Возвращает размер файла в байтах
Public Function getFileSize(path As String) As Long
    On Error GoTo Handler

    Dim fs, f

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(path)
    getFileSize = f.size

Handler:
    getFileSize = 0

End Function



Public Sub test()
    
    Dim tests As New Dictionary

    Dim data As Dictionary

    ' Full path shorter than 255    
    Set data = New Dictionary
    data.item("fullPath") = "C:\abc\def.ext"
    data.item("answer") = "C:\abc\def.ext"
    data.item("min") = 10
    data.item("limit") = 50
    Set tests.item("cutFullPath: TEST 01") = data

    ' Path without filename longer than 255
    Set data = New Dictionary
    data.item("fullPath") = "C:\qwertyuiopqwertyuiopqwertyuiopqwertyuiopqwertyuiop\def.ext"
    data.item("answer") = ""
    data.item("min") = 10
    data.item("limit") = 50
    Set tests.item("cutFullPath: TEST 02") = data
    
    ' Shorten basename less than min
    Set data = New Dictionary
    data.item("fullPath") = "C:\qwertyuiopqwertyuiopqwertyuiopqwertyuiopq\def.ext"
    data.item("answer") = ""
    data.item("min") = 10
    data.item("limit") = 50
    Set tests.item("cutFullPath: TEST 03") = data

    'Shorten basename >= min
    Set data = New Dictionary
    data.item("fullPath") = "C:\qwertyuiopqwertyuiopqwertyuiopqwertyuiopq\def.ext"
    data.item("answer") = "C:\qwertyuiopqwertyuiopqwertyuiopqwertyuiopq\d.ext"
    data.item("min") = 1
    data.item("limit") = 50
    Set tests.item("cutFullPath: TEST 04") = data

    'Shorten basename >= min and file without extension
    Set data = New Dictionary
    data.item("fullPath") = "C:\qwertyuiopqwertyuiopqwertyuiopqwertyuiopq\abcdefg"
    data.item("answer") = "C:\qwertyuiopqwertyuiopqwertyuiopqwertyuiopq\abcde"
    data.item("min") = 1
    data.item("limit") = 50
    Set tests.item("cutFullPath: TEST 05") = data

    ' Full path is for server    
    Set data = New Dictionary
    data.item("fullPath") = "\\server\abc\def.ext"
    data.item("answer") = "\\server\abc\def.ext"
    data.item("min") = 10
    data.item("limit") = 50
    Set tests.item("cutFullPath: TEST 06") = data

    'Shorten basename >= min with russian symbols
    Set data = New Dictionary
    data.item("fullPath") = "C:\йцукенгшщзйцукенгшщзйцукенгшщзйцукенгшщзй\def.ext"
    data.item("answer") =   "C:\йцукенгшщзйцукенгшщзйцукенгшщзйцукенгшщзй\d.ext"
    data.item("min") = 1
    data.item("limit") = 50
    Set tests.item("cutFullPath: TEST 07") = data

    Dim name
    Dim value As Variant
    Dim countOfCorrect As Integer
    countOfCorrect = 0

    For Each name In tests.Keys
        Debug.Print name

        Set value = tests.item(name)
        If cutFullPath(value.item("fullPath"), value.Item("min"), value.Item("limit")) = value.item("answer") Then
            Debug.Print "PASSED"
            countOfCorrect = countOfCorrect + 1
        Else
            Debug.Print "FAILED"
        End If
    Next
    
    Debug.Print "Success: " & Str(countOfCorrect) & " /" & Str(tests.Count) & " = " & Format(countOfCorrect / tests.Count * 100, "#,##0.00") & "%"
End Sub