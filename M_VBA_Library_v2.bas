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
               Exit Function
            End If
        End If
    Next i

    isRevisionCorrect = True

ErrorHandler:
End Function

Private Function isRevisionWithRuleCorrect(rev As String) As Boolean

    isRevisionWithRuleCorrect = False

    Dim pos As Integer
    Dim minPos As Integer
    minPos = 1000

    Dim allowedSymbols As Variant
    allowedSymbols = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", "*")

    Dim i As Integer
    For i = LBound(allowedSymbols) To UBound(allowedSymbols)
        pos = InStr(rev, allowedSymbols(i))
        if pos > 0 And pos < minPos Then
            minPos = pos
        End If
    Next i

    If minPos = 1000 Then
        Exit Function
    End If

    Dim rulePart As String
    rulePart = Left(rev, minPos - 1)

    Select Case rulePart
        Case "", "=", ">", "<", ">=", "<="
            
        Case Else
            Exit Function
    End Select

    Dim revPart As String
    revPart = Mid(rev, minPos)

    isRevisionWithRuleCorrect = isRevisionCorrect(revPart)

End Function

Private Function makeDictionaryRule(rev As String) As Dictionary
    Dim result As New Dictionary

    Dim allowedRules As Variant
    allowedRules = Array(">=", "<=", "=", "<", ">")

    Dim i As Integer
    Dim pos As Integer
    Dim revPart As String

    For i = LBound(allowedRules) To UBound(allowedRules)
        pos = InStr(rev, allowedRules(i))
        if pos > 0  Then
            result.Item("RULE") = allowedRules(i)
            revPart = Mid(rev, Len(allowedRules(i))+1)
            Exit For
        End If
    Next i

    If pos = 0 Then
        result.Item("RULE") = "="
        revPart = rev
    End if

    Dim versionArray() As String
    versionArray = Split(revPart,".")

    result.Item("MAJOR") = versionArray(0)
    result.Item("MINOR") = versionArray(1)
    result.Item("PATCH") = versionArray(2)

    Set makeDictionaryRule = result
    Set result = Nothing
End Function

Private Function isRevEqual( _
    majorOriginal As String, _
    minorOriginal As String, _
    patchOriginal As String, _
    majorTest As String, _
    minorTest As String, _
    patchTest As String _
)
    If _
        (majorOriginal = majorTest Or majorOriginal ="*" Or majorTest = "*") And _
        (minorOriginal = minorTest Or minorOriginal ="*" Or minorTest = "*") And _
        (patchOriginal = patchTest Or patchOriginal ="*" Or patchTest = "*") _
    Then
        isRevEqual = True
    Else
        isRevEqual = False
    End If

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

Public Sub testIsRevisionWithRuleCorrect()
    Dim count As Integer
    count = 0

    Dim test As New Dictionary

    test.Item("1.2.3") = True
    test.Item("=1.2.3") = True
    test.Item(">1.2.3") = True
    test.Item("<1.2.3") = True
    test.Item(">=1.2.3") = True
    test.Item("<=1.2.3") = True
    test.Item("^1.2.3") = False
    test.Item("A1.2.3") = False
    test.Item("<<1.2.3") = False
    test.Item("A1.2.3") = False
    test.Item("<1.2.3<") = False
    test.Item(">1.>2.3") = False
    test.Item("1>1.2.3") = False
    test.Item(">1,2.3") = False

    

    Dim varKey As variant
    For Each varKey In test.Keys
        if isRevisionWithRuleCorrect(CStr(varKey)) = test.Item(varKey) Then
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

Public Sub testMakeDictionaryRule()
    Dim count As Integer
    count = 0

    Dim correct1 As New Dictionary
    correct1.Item("RULE") = "="
    correct1.Item("MAJOR") = "1"
    correct1.Item("MINOR") = "2"
    correct1.Item("PATCH") = "3"

    Dim correct2 As New Dictionary
    correct2.Item("RULE") = "="
    correct2.Item("MAJOR") = "1"
    correct2.Item("MINOR") = "2"
    correct2.Item("PATCH") = "3"

    Dim correct3 As New Dictionary
    correct3.Item("RULE") = ">"
    correct3.Item("MAJOR") = "1"
    correct3.Item("MINOR") = "2"
    correct3.Item("PATCH") = "3"

    Dim correct4 As New Dictionary
    correct4.Item("RULE") = "<"
    correct4.Item("MAJOR") = "1"
    correct4.Item("MINOR") = "2"
    correct4.Item("PATCH") = "3"

    Dim correct5 As New Dictionary
    correct5.Item("RULE") = ">="
    correct5.Item("MAJOR") = "1"
    correct5.Item("MINOR") = "2"
    correct5.Item("PATCH") = "3"

    Dim correct6 As New Dictionary
    correct6.Item("RULE") = "<="
    correct6.Item("MAJOR") = "1"
    correct6.Item("MINOR") = "2"
    correct6.Item("PATCH") = "3"

    Dim correct7 As New Dictionary
    correct7.Item("RULE") = "<="
    correct7.Item("MAJOR") = "*"
    correct7.Item("MINOR") = "*"
    correct7.Item("PATCH") = "*"

    Dim test As New Dictionary

    Set test.Item("1.2.3") = correct1
    Set test.Item("=1.2.3") = correct2
    Set test.Item(">1.2.3") = correct3
    Set test.Item("<1.2.3") = correct4
    Set test.Item(">=1.2.3") = correct5
    Set test.Item("<=1.2.3") = correct6
    Set test.Item("<=*.*.*") = correct7
    

    Dim varKey As variant
    Dim result As Dictionary
    Dim major As String
    Dim minor As String
    Dim patch As String
    Dim rule As String
    For Each varKey In test.Keys

        Set result = makeDictionaryRule(CStr(varKey))
        rule = test.Item(varKey).Item("RULE")
        major = test.Item(varKey).Item("MAJOR")
        minor = test.Item(varKey).Item("MINOR")
        patch = test.Item(varKey).Item("PATCH")

        if _
            result.Item("RULE") = rule And _
            result.Item("MAJOR") = major And _
            result.Item("MINOR") = minor And _
            result.Item("PATCH") = patch _
        Then
            count = count + 1
        Else   
            Debug.Print "Test No." & Str(count + 1) & " - FAILED"

            Debug.Print varKey
            Exit Sub
        End If
    Next

    Debug.Print Str(count) & " tests PASSED"

    Set correct1 = Nothing
    Set correct2 = Nothing
    Set correct3 = Nothing
    Set correct4 = Nothing
    Set correct5 = Nothing
    Set correct6 = Nothing
    Set correct7 = Nothing
    Set result = Nothing

End Sub

Public Sub testIsRevEqual()
    Dim count As Integer
    count = 0

    Dim test As New Dictionary

    test.Item("1.2.3|1.2.3") = True
    test.Item("1.2.3|1.2.*") = True
    test.Item("1.2.3|1.*.3") = True
    test.Item("1.2.3|1.*.3") = True
    test.Item("1.2.3|*.2.3") = True
    test.Item("1.2.3|*.*.3") = True
    test.Item("1.2.3|1.*.*") = True
    test.Item("1.2.3|*.2.*") = True
    test.Item("1.2.3|*.*.*") = True
    test.Item("1.2.3|1.3.3") = False
    test.Item("1.2.3|*.3.3") = False
    test.Item("1.2.3|1.3.*") = False
    test.Item("1.2.3|*.3.*") = False
    test.Item("1.2.3|*.3.*") = False
    

    Dim varKey As variant
    Dim revisions() As String
    Dim originalRevs() As String
    Dim testRevs() As String
    For Each varKey In test.Keys

        revisions = Split(CStr(varKey),"|")
        originalRevs = Split(revisions(0),".")
        testRevs = Split(revisions(1),".")

        if isRevEqual( _
            originalRevs(0), originalRevs(1), originalRevs(2), _
            testRevs(0), testRevs(1), testRevs(2) _
        ) = test.Item(varKey) Then
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


