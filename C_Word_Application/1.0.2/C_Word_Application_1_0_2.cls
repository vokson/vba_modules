VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Word_Application"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const VERSION = "1.0.2"

Private m_wordApp As Object
  
Private Sub Class_Initialize() 
  On Error Resume Next

  Set m_wordApp = GetObject(, "Word.Application")
  
  If TypeName(m_wordApp) = "Nothing" Then
      Set m_wordApp = CreateObject("Word.Application")
      m_wordApp.Visible = True
  End If

End Sub 

Private Sub Class_Terminate()
  Set m_wordApp = Nothing
End Sub

Public Function GetApplication() As Object
  Set GetApplication = m_wordApp
End Function

Public Sub replace(range As Object, needle As String, replacement As String)
  With range.Find
    .text = needle
    While .Execute
      range.Select
      m_wordApp.Selection.TypeText text:=replacement
      range.Collapse Direction:=0 ' wdCollapseEnd
    Wend
  End With
End Sub
