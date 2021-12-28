VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Outlook_Application"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const VERSION = "1.1.1"

Private m_outlookApp As Object
Private clsFile As Object
  
Private Sub Class_Initialize() 
  On Error Resume Next

  Set clsFile = New C_File
  Set m_outlookApp = GetObject(, "Outlook.Application")
  
  If TypeName(m_outlookApp) = "Nothing" Then
      Set m_outlookApp = CreateObject("Outlook.Application")
      m_outlookApp.Visible = True
  End If

End Sub 

Private Sub Class_Terminate()
  Set m_outlookApp = Nothing
End Sub

Public Function GetOutlook() As Object
  Set GetOutlook = m_outlookApp
End Function

Public Function GetSelectedLetters() As C_ImprovedCollection
  Set GetSelectedLetters = New C_ImprovedCollection
  Dim message As Object

  For Each message In m_outlookApp.ActiveExplorer.Selection
    ' 43 = olMail
    If message.Class = 43 Then
      GetSelectedLetters.Add message
    End If
  Next
 
End Function

' Функция определяет адрес текущего пользователя по умолчанию и возвращает его
Public Function getAddressOfCurrentUser() As String
  With m_outlookApp.Session
      
    If .CurrentUser.AddressEntry.Type = "EX" Then
        getAddressOfCurrentUser = .CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress
    Else
        getAddressOfCurrentUser = .CurrentUser.address
    End If
  End With

End Function

' Функция ищет объект Account по SMTP адресу и возвращает его
' (String) address - email адрес
' Account заменен на Object из-за late-binding
Public Function getAccountByAddress(address As String) As Object
  With m_outlookApp.Session

    Dim id As Object
    For Each id In .Accounts

      if id.SmtpAddress = address Then
        Set getAccountByAddress = id
        Exit Function
      End If

    Next

  End With

  Set getAccountByAddress = Nothing
End Function

' Функция проверяет является ли картинка вставленной в текст HTML письма
' (MailItem.Attachment) att - объект вложения
Public Function isEmbeddedAttachment(att As Object) As Boolean
    Const PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"

    Dim ext As String
    ext = LCase(clsFile.getExtension(att.FileName))
    if Not( _
        ext = "bmp" Or ext = "jpg" Or ext = "jpeg" Or _
        ext = "png" Or ext = "tif" Or ext = "tiff" Or _
        ext = "eps" Or ext = "ps" Or ext = "svg" Or _
        ext = "swf" Or ext = "wmf" Or ext = "gif" Or _
        ext = "ico" Or ext = "cur" Or ext = "webp" Or _
        ext = "jfif" Or ext = "pjpeg" Or ext = "pjp" Or _
        ext = "apng" Or ext = "avif" _
    ) Then
        isEmbeddedAttachment = False
        Exit Function
    End If

    Dim PropAccessor As Object ' PropertyAccessor
    Set PropAccessor = Att.PropertyAccessor
    isEmbeddedAttachment = (PropAccessor.GetProperty(PR_ATTACH_CONTENT_ID) <> "")
End Function