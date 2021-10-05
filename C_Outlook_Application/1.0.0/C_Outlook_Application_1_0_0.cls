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

Const VERSION = "1.0.0"

Private m_outlookApp As Object
  
Private Sub Class_Initialize() 
  On Error Resume Next

  Set m_outlookApp = GetObject(, "Outlook.Application")
  
  If TypeName(m_outlookApp) = "Nothing" Then
      Set m_outlookApp = CreateObject("Outlook.Application")
      m_outlookApp.Visible = True
  End If

End Sub 

Private Sub Class_Terminate()
  Set m_outlookApp = Nothing
End Sub

Public Function GetSelectedLetters() As C_ImprovedCollection
  Set GetSelectedLetters = New C_ImprovedCollection
  Dim message As Object

  For Each message In m_outlookApp.ActiveExplorer.Selection
    If message.Class = olMail Then
      GetSelectedLetters.Add message
    End If
  Next
 
End Function


