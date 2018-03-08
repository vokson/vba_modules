Attribute VB_Name = "M_VBA_Library_v2"
Option Explicit

Const FOLDER_WITH_VBA_MODULES = "vba_modules"
Const DIRECTORY_SEPARATOR = "\"

Public Sub importLibraryList()

    Dim ModuleName
    Dim Version As String
    Dim listOfModules As Dictionary
    Set listOfModules = getListOfRequiredModules()
    
    For Each ModuleName In listOfModules.Keys
        Version = listOfModules.Item(ModuleName)
        Call importRequiredModule(CStr(ModuleName), Version)
    Next
    
End Sub

Public Sub importRequiredModule(ModuleName As String, Version As String)
    If isModuleExist(ModuleName, Version) = True Then Call DeleteModule(ModuleName, Version)
    Call ImportModule(ModuleName, Version)
End Sub

Private Sub ImportModule(ModuleName As String, Version As String)
   
   Dim nameWithoutExtension As String
   nameWithoutExtension = Application.ActiveWorkbook.path & DIRECTORY_SEPARATOR & _
        FOLDER_WITH_VBA_MODULES & DIRECTORY_SEPARATOR & ModuleName & DIRECTORY_SEPARATOR & _
        combineModuleNameWithVersion(ModuleName, Version)
   
   If Dir(nameWithoutExtension & ".bas") <> "" Then
        Application.VBE.ActiveVBProject.VBComponents.Import (nameWithoutExtension & ".bas")
        
   ElseIf Dir(nameWithoutExtension & ".cls") <> "" Then
        Application.VBE.ActiveVBProject.VBComponents.Import (nameWithoutExtension & ".cls")
        
   Else
        MsgBox "Module " & ModuleName & " is NOT found."
        
   End If
   
End Sub

Private Sub DeleteModule(ModuleName As String, Version As String)
    Dim ModuleNameWithVersion As String
'    ModuleNameWithVersion = combineModuleNameWithVersion(ModuleName, Version)
    ModuleNameWithVersion = ModuleName
    
    With Application.VBE.ActiveVBProject.VBComponents
        If .Item(ModuleNameWithVersion).Type <> 100 Then ' vbext_ct_Document
            .Item(ModuleNameWithVersion).name = ModuleNameWithVersion & "_OLD"
            .Remove .Item(ModuleNameWithVersion & "_OLD")
        Else
            .Remove .Item(ModuleNameWithVersion)
        End If
    End With
End Sub

Private Function isModuleExist(ModuleName As String, Version As String)
    On Error Resume Next
    
    isModuleExist = False
    
    Dim moduleType As Integer
    moduleType = Application.VBE.ActiveVBProject.VBComponents.Item(ModuleName).Type
'    moduleType = Application.VBE.ActiveVBProject.VBComponents.Item(combineModuleNameWithVersion(ModuleName, Version)).Type
    
    If Err.number = 0 Then isModuleExist = True
    
End Function

Private Function combineModuleNameWithVersion(ModuleName As String, Version As String) As String
    combineModuleNameWithVersion = ModuleName & "_" & Version
End Function

Private Function isRevisionCorrect(rev As String) As Boolean
On Error GoTo ErrorHandler

    Dim i As Integer
    Dim symbol as String

    isRevisionCorrect = False

    For i = 1 To Len(rev)

        symbol = Mid(rev, i, 1)

        Select Case symbol
            Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", "*"
                
            Case Else
                Exit Function
        End Select

    Next i

    Dim strVer As String
    Dim intVer As Integer
    Dim versionArray() As String
    versionArray = Split(rev,".")

    if (UBound(versionArray)-LBound(versionArray)+1) <> 3 Then
        Exit Function
    End If

    For i = LBound(versionArray) To UBound(versionArray)
        If versionArray(i) <> "*" Then
            strVer = versionArray(i)
            intVer = CInt(strVer)
            
            If (Trim(CStr(intVer)) <> intVer) Or  (intVer < 0) Then
               
            End If
        End If
    Next i

    isRevisionCorrect = True

ErrorHandler:
End Function

Public Sub testIsRevisionCorrect()
    Dim count As Integer
    count = 0

    Dim test As New Dictionary

    test.Item("0.0.0") = True
    test.Item("1.0.*") = True
    test.Item("1.*.1") = True
    test.Item("*.1.1") = True
    test.Item("1.*.*") = True
    test.Item("*.*.1") = True
    test.Item("*.1.*") = True
    test.Item("*.*.*") = True
    test.Item("-1.1.1") = False
    test.Item(".1.1") = False
    test.Item("1.1") = False
    test.Item("1") = False
    test.Item("1.1.1.1") = False
    test.Item("1.$.1") = False
    test.Item("1.A.1") = False
    test.Item("23.43.12") = True

    Dim varKey As variant
    For Each varKey In test.Keys
        if isRevisionCorrect(CStr(varKey)) = test.Item(varKey) Then
            count = count + 1
        Else   
            Debug.Print "Test No." & Str(count + 1) & " - FAILED"

            Debug.Print varKey
            Debug.Print test.Item(varKey)
            Exit Sub
        End If
    Next

    Debug.Print Str(count) & " tests PASSED"

End Sub


