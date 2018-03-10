VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_String"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function deleteDoubleSpaces(s As String) As String
    Dim pos As Integer
    Dim text As String

    text = s

    pos = InStr(1, text, "  ", vbBinaryCompare)
    Do While pos > 0
        text = Replace(text, "  ", " ")
        pos = InStr(1, text, "  ", vbBinaryCompare)
    Loop

    deleteDoubleSpaces = text
End Function

Function replaceNewStringSymbolWith(s As String, replaceText As String) As String
    Dim text As String
    text = s

    text = Replace(text, vbNewLine, replaceText)
    text = Replace(text, vbLf, replaceText)
    text = Replace(text, vbCr, replaceText)

    replaceNewStringSymbolWith = text
End Function