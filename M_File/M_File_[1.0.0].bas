Attribute VB_Name = "M_File"
Option Explicit

' Module M_File
' Operation with files
' Ver. 1.0.1

Function getExtensionOfFile(fileName As String) As String
   getExtensionOfFile = Right(fileName, Len(fileName) - InStrRev(fileName, "."))
End Function

Function getFileNameFromFullPath(path As String) As String
   getFileNameFromFullPath = Right(path, Len(path) - InStrRev(path, DIRECTORY_SEPARATOR))
End Function

Sub writeFile(fullPath As String, text As String)
    Dim fs, a
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(fullPath, True)
    a.Write (text)
    a.Close
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



