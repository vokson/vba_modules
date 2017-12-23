VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_NIPIGAZ_Validator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Function validateFileName(text As String) As String
    On Error GoTo ErrorHandler

    Dim dic As Dictionary
    Set dic = decodeFileName(text)

    If dic.Item("IS_VALID") = False Then GoTo ErrorHandler

    dic.Item("CONTRACT") = validateContract(dic.Item("CONTRACT"))
    dic.Item("TITLE_1") = validateNumber(dic.Item("TITLE_1"), 1)
    dic.Item("TITLE_2") = validateNumber(dic.Item("TITLE_2"), 1)
    dic.Item("TITLE_3") = validateNumber(dic.Item("TITLE_3"), 1)
    dic.Item("TITLE_4") = validateNumber(dic.Item("TITLE_4"), 2)
    dic.Item("TITLE_5") = validateNumber(dic.Item("TITLE_5"), 3)

    dic.Item("PACKAGE_CODE") = validatePackageCodeWithNumber(dic.Item("PACKAGE_CODE"))
    dic.Item("DOCUMENT_TYPE_CODE") = validateDocumentTypeCode(dic.Item("DOCUMENT_TYPE_CODE"))
    dic.Item("SEQUENCE_NUMBER") = validateNumber(dic.Item("SEQUENCE_NUMBER"), 4)

    dic.Item("REVISION") = validateRevision(dic.Item("REVISION"))
    dic.Item("LANGUAGE") = validateLanguage(dic.Item("LANGUAGE"))
    dic.Item("CRS") = validateCRS(dic.Item("CRS"))
    dic.Item("ATTACHMENT") = validateAttachment(dic.Item("ATTACHMENT"))
    dic.Item("EXTENSION") = validateExtension(dic.Item("EXTENSION"))

    validateFileName = encodeFileName(dic)
    Exit Function

ErrorHandler:
    validateFileName = ""
End Function


Public Function validateDocumentName(text As String) As String
    On Error GoTo ErrorHandler

    Dim dic As Dictionary
    Set dic = decodeDocumentName(text)

    If dic.Item("IS_VALID") = False Then GoTo ErrorHandler

    dic.Item("CONTRACT") = validateContract(dic.Item("CONTRACT"))
    dic.Item("TITLE_1") = validateNumber(dic.Item("TITLE_1"), 1)
    dic.Item("TITLE_2") = validateNumber(dic.Item("TITLE_2"), 1)
    dic.Item("TITLE_3") = validateNumber(dic.Item("TITLE_3"), 1)
    dic.Item("TITLE_4") = validateNumber(dic.Item("TITLE_4"), 2)
    dic.Item("TITLE_5") = validateNumber(dic.Item("TITLE_5"), 3)

    dic.Item("PACKAGE_CODE") = validatePackageCodeWithNumber(dic.Item("PACKAGE_CODE"))
    dic.Item("DOCUMENT_TYPE_CODE") = validateDocumentTypeCode(dic.Item("DOCUMENT_TYPE_CODE"))
    dic.Item("SEQUENCE_NUMBER") = validateNumber(dic.Item("SEQUENCE_NUMBER"), 4)

    validateDocumentName = encodeDocumentName(dic)
    Exit Function

ErrorHandler:
    validateDocumentName = ""
End Function

Private Function encodeFileName(dic As Dictionary) As String

    encodeFileName = encodeDocumentName(dic) & "_"
    encodeFileName = encodeFileName & dic.Item("REVISION") & "_"
    encodeFileName = encodeFileName & dic.Item("LANGUAGE")

    If dic.Item("CRS") <> "" Then  encodeFileName = encodeFileName & "_" &  dic.Item("CRS")
    If dic.Item("ATTACHMENT") <> "" Then  encodeFileName = encodeFileName & "_" &  dic.Item("ATTACHMENT")
       
    encodeFileName = encodeFileName & "." & dic.Item("EXTENSION")

End Function

Private Function encodeDocumentName(dic As Dictionary) As String
    encodeDocumentName = dic.Item("CONTRACT") & "-"
    encodeDocumentName = encodeDocumentName & dic.Item("TITLE_1") & "."
    encodeDocumentName = encodeDocumentName & dic.Item("TITLE_2") & "."
    encodeDocumentName = encodeDocumentName & dic.Item("TITLE_3") & "."
    encodeDocumentName = encodeDocumentName & dic.Item("TITLE_4") & "."
    encodeDocumentName = encodeDocumentName & dic.Item("TITLE_5") & "-"
    encodeDocumentName = encodeDocumentName & dic.Item("PACKAGE_CODE") & "."
    encodeDocumentName = encodeDocumentName & dic.Item("DOCUMENT_TYPE_CODE") & "-"
    encodeDocumentName = encodeDocumentName & dic.Item("SEQUENCE_NUMBER")
End Function

Private Function decodeFileName(name As String) As Dictionary
    Dim dic As New Dictionary
    Dim pos As Integer
    Dim s As String

    s = name
    dic.Item("IS_VALID") = True

    ' Document Name
    pos = InStr(1, s, "_")
    If pos = 0 Then
        dic.Item("IS_VALID") = False
        GoTo Finish
    End If
    Set dic = decodeDocumentName(Left(s, pos-1))
    s = Mid(s, pos + 1)

    'Revision
    pos = InStr(1, s, "_")
    If pos = 0 Then
        dic.Item("IS_VALID") = False
        GoTo Finish
    End If
    dic.Item("REVISION") = Left(s, pos-1)
    s = Mid(s, pos+1)

    'Language
    pos = InStr(1, s, "_")
    if pos = 0 Then pos = InStr(1, s, ".")

    If pos = 0 Then
        dic.Item("IS_VALID") = False
        GoTo Finish
    End If
    dic.Item("LANGUAGE") = Left(s, pos-1)
    s = Mid(s, pos+1)

    'Text Before Extension
    pos = InStr(1, s, ".")
    dic.Item("TEXT_BEFORE_EXTENSION") = ""
    If pos > 0 Then
        dic.Item("TEXT_BEFORE_EXTENSION") = Left(s, pos-1)
        s = Mid(s, pos+1)
    End If

    'CRS

    if Len(dic.Item("TEXT_BEFORE_EXTENSION")) > 0 Then

        pos = InStr(1, dic.Item("TEXT_BEFORE_EXTENSION"), "_")

        If pos > 0 Then
            dic.Item("CRS") = Left(dic.Item("TEXT_BEFORE_EXTENSION"), pos-1)
            dic.Item("ATTACHMENT") = Mid(dic.Item("TEXT_BEFORE_EXTENSION"), pos+1)

        ElseIf  InStr(1, dic.Item("TEXT_BEFORE_EXTENSION"), "CRS") > 0 Then
            dic.Item("CRS") = dic.Item("TEXT_BEFORE_EXTENSION")
            dic.Item("ATTACHMENT") = ""

        Else
            dic.Item("CRS") = ""
            dic.Item("ATTACHMENT") = dic.Item("TEXT_BEFORE_EXTENSION")
        End If

    End If

   'Extension
     If s = "" Then
        dic.Item("IS_VALID") = False
        GoTo Finish
    End If
    dic.Item("EXTENSION") = s

Finish:
    Set decodeFileName = dic
    Set dic = Nothing
End Function

Private Function decodeDocumentName(name As String) As Dictionary
    Dim dic As New Dictionary
    Dim pos As Integer
    Dim s As String

    s = name
    dic.Item("IS_VALID") = True

    'Contract
    pos = InStr(1, s, "-")
    If pos = 0 Then
        dic.Item("IS_VALID") = False
        GoTo Finish
    End If
    dic.Item("CONTRACT") = Left(s, pos-1)
    s = Mid(s, pos+1)

    'Title_1
    pos = InStr(1, s, ".")
    If pos = 0 Then
        dic.Item("IS_VALID") = False
        GoTo Finish
    End If
    dic.Item("TITLE_1") = Left(s, pos-1)
    s = Mid(s, pos+1)

    'Title_2
    pos = InStr(1, s, ".")
    If pos = 0 Then
        dic.Item("IS_VALID") = False
        GoTo Finish
    End If
    dic.Item("TITLE_2") = Left(s, pos-1)
    s = Mid(s, pos+1)

    'Title_3
    pos = InStr(1, s, ".")
    If pos = 0 Then
        dic.Item("IS_VALID") = False
        GoTo Finish
    End If
    dic.Item("TITLE_3") = Left(s, pos-1)
    s = Mid(s, pos+1)

    'Title_4
    pos = InStr(1, s, ".")
    If pos = 0 Then
        dic.Item("IS_VALID") = False
        GoTo Finish
    End If
    dic.Item("TITLE_4") = Left(s, pos-1)
    s = Mid(s, pos+1)

    'Title_5
    pos = InStr(1, s, "-")
    If pos = 0 Then
        dic.Item("IS_VALID") = False
        GoTo Finish
    End If
    dic.Item("TITLE_5") = Left(s, pos-1)
    s = Mid(s, pos+1)

    'Package code
    pos = InStr(1, s, ".")
    If pos = 0 Then
        dic.Item("IS_VALID") = False
        GoTo Finish
    End If
    dic.Item("PACKAGE_CODE") = Left(s, pos-1)
    s = Mid(s, pos+1)

    'Document Type Code
    pos = InStr(1, s, "-")
    If pos = 0 Then
        dic.Item("IS_VALID") = False
        GoTo Finish
    End If
    dic.Item("DOCUMENT_TYPE_CODE") = Left(s, pos-1)
    s = Mid(s, pos+1)

    'Sequence Number
     If s = "" Then
        dic.Item("IS_VALID") = False
        GoTo Finish
    End If
    dic.Item("SEQUENCE_NUMBER") = s

Finish:
    Set decodeDocumentName = dic
    Set dic = Nothing

End Function

' Get proper value, if value is correct
' Get "", if value is not correct
Private Function validateContract(text As String) As String
    Dim number As Integer
    On Error GoTo ErrorHandler

    number = CInt(text)
    if number = 55 Then
        validateContract = Format(number, "0000")
        Exit Function
    End If

ErrorHandler:
    validateContract = ""

End Function

Private Function validateNumber(text As String, countOfSymbols As Integer) As String
    Dim number As Integer
    On Error GoTo ErrorHandler

    number = CInt(text)

    Dim valueFormat As String
    Dim i As Integer
    valueFormat = ""
    For i = 1 To countOfSymbols
        valueFormat = valueFormat & "0"
    Next i

    if (Format(number, valueFormat) = text) and (Len(text) = countOfSymbols) Then
        validateNumber = text
        Exit Function
    End If

ErrorHandler:
    validateNumber = ""
End Function


Private Function validatePackageCodeWithNumber(text As String) As String
    Dim code As String
    Dim number As Integer
    Dim strNumber As String
    On Error GoTo ErrorHandler

    number = 0

    if Len(text) = 2 Then
        code = text
    ElseIf Len(text) > 2 Then
        If IsNumeric(Mid(text, 3)) Then
            code = Mid(text, 1, 2)
            number = CInt(Mid(text, 3))
        Else
            code = Mid(text, 1, 3)
            number = CInt(Mid(text, 4))
        End If
    End If

    strNumber = ""
    if number > 0 Then strNumber = CStr(number)

    if validatePackageCode(code) & strNumber = text Then
        validatePackageCodeWithNumber = text
        Exit Function
    End If

ErrorHandler:
    validatePackageCodeWithNumber = ""
End Function

Private Function validatePackageCode(text As String) As String
    validatePackageCode = validateStringFromArray(text, _
        Array("VK", "NV", "NVK", "NK", "PT", "TM", "THM", "TI", "VS", "GSV", "GSN", "TS", "TT", "OE", "AD", "GP", "GT", "IZT", _
                         "PJ", "AR", "AS", "KJ", "KM", "VL", "RZA", "EPO", "EPZ", "ER", "ERZ", "EOE", "EHZ", "TH", "AK", "AKZ", "ANK", _
                          "ANV", "AE", "OV", "OVO", "OVK", "ASK", "ATH", "SSL", "OS", "AUP", "SS", "RT", "PS", "SM") _
    )
End Function

Private Function validateDocumentTypeCode(text As String) As String
     validateDocumentTypeCode = validateStringFromArray(text, _
        Array("OD", "PL", "KP", "PR", "CP", "DT", "ID", "SC", "RR", "CJ", "LT", "LS", "LR", "OS", "UO", "PID", "PFD", _
                             "IS", "SP", "OL", "IZ", "NI", "PU", "TDA", "TT", "PO", "PI") _
    )
End Function

Public Function validateRevision(text As String) As String
    validateRevision = validateStringFromArray(text, _
        Array("VD", "SD", "A1", "B1") _
    )

    if validateRevision = "" Then  validateRevision = validateNumber(text, 2)

End Function

Public Function validateLanguage(text As String) As String
    validateLanguage = validateStringFromArray(text, _
        Array("RU", "EN", "ER") _
    )
End Function

Private Function validateExtension(text As String) As String
    validateExtension = validateStringFromArray(text, _
        Array("dwg", "DWG", "pdf", "PDF", "doc", "DOC", "docx", "DOCX", "xls", "XLS", "xlsx", "XLSX", "zip", "ZIP") _
    )
End Function

Private Function validateCRS(text As String) As String
    validateCRS = validateStringFromArray(text, Array("CRS", "ACRS") )
End Function

Private Function validateStringFromArray(text As String, validCodes) As String
    On Error GoTo ErrorHandler

    Dim i As Integer
    For i = LBound(validCodes) To UBound(validCodes)
        if validCodes(i) = text Then
            validateStringFromArray = text
            Exit Function
        End If
    Next i

ErrorHandler:
    validateStringFromArray = ""
End Function

Private Function validateAttachment(text As String) As String
    On Error GoTo ErrorHandler

    Dim att As String
    Dim number As String

    att = Left(text, 3)
    number = Mid(text, 4)

    If (att = "Att" Or att = "ATT") And IsNumeric(number) Then
        validateAttachment = text
        Exit Function
    End If

ErrorHandler:
    validateAttachment = ""
End Function