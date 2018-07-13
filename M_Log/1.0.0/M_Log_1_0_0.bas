Attribute VB_Name = "M_Log"
Option Explicit

Const PATH_SEPARATOR = "\"
Const LOG_FILENAME = "log.txt"
Const IS_LOG_ACTIVE = True

Dim pathToLogFile As String

' Очищает лог файл

Public Sub clearLog(Optional fullPath As String = "")

    if Not IS_LOG_ACTIVE Then Exit Sub

    If fullPath <> "" Then
        pathToLogFile = fullPath
    else
        pathToLogFile = Application.ActiveWorkbook.Path & PATH_SEPARATOR & LOG_FILENAME
    End If

    Dim fs, f
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.CreateTextFile(pathToLogFile, True)

    f.Close

    Set fs = Nothing
    Set f = Nothing

End Sub

' Пишет в лог файл строку
' Перед записью необходимо очистить лог файл для присвоения правильного пути лог файла
Public Sub log(text As String)

   if Not IS_LOG_ACTIVE Then Exit Sub

   Dim fso, f

   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.OpenTextFile(pathToLogFile, 8)

   f.WriteLine text
   f.Close

   Set fso = Nothing
   Set f = Nothing

End Sub


