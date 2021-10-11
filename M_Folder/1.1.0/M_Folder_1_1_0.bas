Attribute VB_Name = "M_Folder"
Option Explicit
' Module M_Folder
' Operation with directories
' Ver. 1.0.1

Public Const DIRECTORY_SEPARATOR = "\"

Public Sub makeDirectory(FolderPath As String, Optional ignoreBeggingPart As String = "")
    Dim x, i As Integer, strPath As String
    x = Split(FolderPath, DIRECTORY_SEPARATOR)

    For i = 0 To UBound(x)
        strPath = strPath & x(i) & DIRECTORY_SEPARATOR

        if InStr(1, ignoreBeggingPart, strPath) = 0 Then
            If Not isFolderExists(strPath) Then MkDir strPath
        End If

    Next i
End Sub

Function isFolderExists(FolderPath As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("scripting.filesystemobject")
    isFolderExists =  fso.FolderExists(FolderPath)
End Function

Function getPathFromFullPath(path As String) As String
   getPathFromFullPath = Left(path, InStrRev(path, DIRECTORY_SEPARATOR) - 1)
End Function

' Search files in folder and subfolders
' strFolder - Path to folder
' strFileSpec - Mask of file
' bIncludeSubfolders - Is subfolders included
Public Sub findFilesInDirectory(colFiles As Collection, _
                             strFolder As String, _
                             strFileSpec As String, _
                             bIncludeSubfolders As Boolean)

    Dim strTemp As String
    Dim colFolders As New Collection
    Dim vFolderName As Variant

    'Add files in strFolder matching strFileSpec to colFiles
    strFolder = TrailingSlash(strFolder)
    strTemp = Dir(strFolder & strFileSpec)
    Do While strTemp <> vbNullString
        colFiles.Add strFolder & strTemp
        strTemp = Dir
    Loop

    If bIncludeSubfolders Then
        'Fill colFolders with list of subdirectories of strFolder
        strTemp = Dir(strFolder, vbDirectory)
        Do While strTemp <> vbNullString
            If (strTemp <> ".") And (strTemp <> "..") Then
                If (GetAttr(strFolder & strTemp) And vbDirectory) <> 0 Then
                    colFolders.Add strTemp
                End If
            End If
            strTemp = Dir
        Loop

        'Call RecursiveDir for each subfolder in colFolders
        For Each vFolderName In colFolders
            Call findFilesInDirectory(colFiles, strFolder & vFolderName, strFileSpec, True)
        Next vFolderName
    End If

End Sub
' Search files in folder and subfolders
' strFolder - Path to folder
' strFileSpec - Mask of file
' bIncludeSubfolders - Is subfolders included
Public Sub findFoldersInDirectory(colFolders As Collection, _
                             strFolder As String, _
                             bIncludeSubfolders As Boolean)
     
    Dim foldersInCurrentDirectory As New Collection
    'Fill colFolders with list of subdirectories of strFolder
    strFolder = TrailingSlash(strFolder)
    
    Dim strTemp As String
    strTemp = Dir(strFolder, vbDirectory)
    Do While strTemp <> vbNullString
        If (strTemp <> ".") And (strTemp <> "..") Then
            If (GetAttr(strFolder & strTemp) And vbDirectory) <> 0 Then
                colFolders.Add strFolder & strTemp
                foldersInCurrentDirectory.Add strTemp
            End If
        End If
        strTemp = Dir
    Loop
    
    If bIncludeSubfolders Then
        Dim vFolderName As Variant
        For Each vFolderName In foldersInCurrentDirectory
            Call findFoldersInDirectory(colFolders, strFolder & vFolderName, True)
        Next vFolderName
    End If
    
End Sub
'
'Public Sub test1()
'    Dim colFolders As New Collection
'
'    Call findFoldersInDirectory(colFolders, "D:\120", True)
'
'    Dim vFolderName As Variant
'    For Each vFolderName In colFolders
'            Debug.Print vFolderName
'    Next vFolderName
'End Sub

'Public Sub test2()
'    Dim colFiles As New Collection
'
'    Call findFilesInDirectory(colFiles, "D:\130", "*", False)
'
'    Dim vFileName As Variant
'    For Each vFileName In colFiles
'            Debug.Print vFileName
'    Next vFileName
'End Sub


Public Function TrailingSlash(strFolder As String) As String
    If Len(strFolder) > 0 Then
        If Right(strFolder, 1) = "\" Then
            TrailingSlash = strFolder
        Else
            TrailingSlash = strFolder & "\"
        End If
    End If
End Function

' Функция формирует имя папки для размещения нумерованных объектов
' Например, для объекта MakeFolderNameForNumberedObjects(333, 5, 100) = "00301-00400"
' (Long) number - номер объекта
' (Integer) countOfDigits - кол-во символов в номере
' (Integer) countOfObjectsInFolder - кол-во объектов в папке 
Public Function MakeFolderNameForNumberedObjects(number As Long, countOfDigits As Integer, countOfObjectsInFolder As Integer) As String
    Dim min As Long
    Dim max As Long

    min = (number \ countOfObjectsInFolder) * countOfObjectsInFolder + 1
    max = (number \ countOfObjectsInFolder + 1) * countOfObjectsInFolder

    If (number Mod countOfObjectsInFolder = 0) Then
        min = min - countOfObjectsInFolder
        max = max - countOfObjectsInFolder
    End If

    Dim pattern As String
    Dim i As Integer
    For i = 1 To countOfDigits
        pattern = pattern & "0"
    Next i

    MakeFolderNameForNumberedObjects = Format(min, pattern) & "-" & Format(max, pattern)
End Function