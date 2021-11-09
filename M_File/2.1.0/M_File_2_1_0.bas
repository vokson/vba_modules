Attribute VB_Name = "M_File"
Option Explicit

' Operation with files
' Ver. 2.1.0

Function getExtensionOfFile(fileName As String) As String
   getExtensionOfFile = Right(fileName, Len(fileName) - InStrRev(fileName, "."))
End Function

Function getFileNameFromFullPath(path As String) As String
   getFileNameFromFullPath = Right(path, Len(path) - InStrRev(path, DIRECTORY_SEPARATOR))
End Function

Sub writeFileInUnicode(fullPath As String, text As String)
    Dim fs, a
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(fullPath, True, True) ' Second True is Unicode format
    a.Write (text)
    a.Close
End Sub

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

Function readFile(fullPath As String) As String
    Dim fso As New FileSystemObject
    Dim JsonTS As TextStream
    
    Set JsonTS = fso.OpenTextFile(fullPath, ForReading)
        readFile = JsonTS.ReadAll
    JsonTS.Close
End Function

Sub copyFile(source As String, destination As String)
    Dim fso As Object
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    
    fso.copyFile source, destination
End Sub

Function isFileExists(fullPath As String) As Boolean
    If Dir(fullPath) <> "" Then
        isFileExists = True
    Else
        isFileExists = False
    End If
End Function

Function replaceRestrictedSymbols(nameOfFile As String, replacement As String) As String
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



