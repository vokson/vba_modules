Attribute VB_Name = "M_VBA_Library_v1"
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


