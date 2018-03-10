VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_TCM_Validator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private ENGLISH_ALPHABET

Private Sub Class_Initialize()
    ENGLISH_ALPHABET = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", _
     "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
End Sub

Public Function validateDocumentName(text As String) As String
    On Error GoTo ErrorHandler

    Dim dic As Dictionary
    Set dic = decodeDocumentName(text)

    If dic.Item("IS_VALID") = False Then GoTo ErrorHandler

    dic.Item("CONTRACT") = validateContract(dic.Item("CONTRACT"))
    dic.Item("CRS") = validateCRS(dic.Item("CRS"))
    dic.Item("DISCIPLINE_CODE") = validateDisciplineCode(dic.Item("DISCIPLINE_CODE"))
    dic.Item("DOCUMENT_TYPE_CODE") = validateDocumentTypeCode(dic.Item("DOCUMENT_TYPE_CODE"))
    dic.Item("SEQUENCE_NUMBER") = validateNumber(dic.Item("SEQUENCE_NUMBER"), 12)

    validateDocumentName = encodeDocumentName(dic)
    Exit Function

ErrorHandler:
    validateDocumentName = ""
End Function

Private Function encodeDocumentName(dic As Dictionary) As String
    encodeDocumentName = dic.Item("CONTRACT") & "-"
    If dic.Item("CRS") <> "" Then  encodeDocumentName = encodeDocumentName & dic.Item("CRS") & "-"
    encodeDocumentName = encodeDocumentName & dic.Item("DISCIPLINE_CODE") & "-"
    encodeDocumentName = encodeDocumentName & dic.Item("DOCUMENT_TYPE_CODE") & "-"
    encodeDocumentName = encodeDocumentName & dic.Item("SEQUENCE_NUMBER")
End Function


Private Function decodeDocumentName(name As String) As Dictionary
  Dim dic As New Dictionary
  Dim pos As Integer

  On Error GoTo ErrorHandler

  dic.Item("IS_VALID") = True

  Dim parts() As String
  parts = Split(name, "-")

  if UBound(parts) < 3 Or UBound(parts) > 4 Then GoTo ErrorHandler

  pos = 0
  'Contract
  dic.Item("CONTRACT") = parts(pos)
  pos = pos + 1

  'CRS
  If UBound(parts) = 4 Then
    dic.Item("CRS") = parts(pos)
    pos = pos + 1
  End If

  'Discipline Code
  dic.Item("DISCIPLINE_CODE") = parts(pos)
  pos = pos + 1

  'Document Type Code
  dic.Item("DOCUMENT_TYPE_CODE") = parts(pos)
  pos = pos + 1

  'Sequence Number
  dic.Item("SEQUENCE_NUMBER") = parts(pos)

ErrorHandler:
  Set decodeDocumentName = dic
  Set dic = Nothing

End Function

Private Function validateContract(text As String) As String
  validateContract = text
  if text <> "3971" Then validateContract = ""
End Function

Private Function validateCRS(text As String) As String
  validateCRS = text
'   if text <> "C" And text <> "CA" Then validateCRS = ""
  if text <> "CA" Then validateCRS = ""
End Function

Private Function validateDisciplineCode(text As String) As String
    Dim firstSymbol As String
    Dim secondSymbol As String
    On Error GoTo ErrorHandler

    If Len(text) <> 2 Then GoTo ErrorHandler

    firstSymbol = validateStringFromArray( Mid(text,1,1), _
        Array("I", "Z", "Q", "H", "M", "P", "L", "O", "Y", "A", _
             "X", "N", "K", "J", "V", "S", "W", "D", "C", "G", "F", "R") _
    )

    secondSymbol = validateStringFromArray( Mid(text,2,1), _
        Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", _
             "N", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z") _
    )

    if (firstSymbol & secondSymbol) = text Then
        validateDisciplineCode = text
        Exit Function
    End If

ErrorHandler:
    validateDisciplineCode = ""
End Function

Private Function validateDocumentTypeCode(text As String) As String
    Dim firstSymbol As String
    Dim secondSymbol As String
    On Error GoTo ErrorHandler

    If Len(text) <> 2 Then GoTo ErrorHandler

    firstSymbol = validateStringFromArray( Mid(text,1,1), ENGLISH_ALPHABET)
    secondSymbol = validateStringFromArray( Mid(text,2,1), ENGLISH_ALPHABET)

    if (firstSymbol & secondSymbol) = text Then
        validateDocumentTypeCode = text
        Exit Function
    End If

ErrorHandler:
    validateDocumentTypeCode = ""
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

Private Function validateNumber(text As String, countOfSymbols As Integer) As String
    On Error GoTo ErrorHandler

    if Len(text) <> countOfSymbols Then GoTo ErrorHandler

    Dim i As Integer
    For i = 1 to Len(text)
        If Not IsNumeric(CInt(Mid(text,i,1))) Then GoTo ErrorHandler
    Next i

    validateNumber = text
    Exit Function

ErrorHandler:
    validateNumber = ""
End Function