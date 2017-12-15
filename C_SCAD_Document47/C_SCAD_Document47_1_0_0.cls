VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_SCAD_Document47"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Make dictionary with groups of elements
Function getGroups(s As String) As Dictionary
  Dim result As New Dictionary
  
  Dim math As New C_Math
  Dim list As New C_SCAD_List
  
  Dim text As String
  text = Trim(s)
  text = str.replaceNewStringSymbolWith(text, " ")
  text = str.deleteDoubleSpaces(text)

  text = Mid(text, InStr(1, text, "(47/", vbBinaryCompare) + 4)
  text = Left(text, InStr(1, text, "/)", vbBinaryCompare) - 1)

  Dim textArray() As String
  textArray = Split(text, "/")

  Dim i As Integer
  Dim groupText As String
  Dim nameText As String
  Dim elementsText As String
  For i = LBound(textArray) To UBound(textArray)
    groupText = textArray(i)

    nameText = Mid(text, InStr(1, text, """", vbBinaryCompare) + 1)
    nameText = Left(nameText, InStr(1, text, """", vbBinaryCompare) - 1)

    elementsText = Mid(text, InStr(1, text, ":", vbBinaryCompare) + 1)

    result.Items(nameText) = list..makeArrayFromList(elementsText)
  Next i

  Set math = Nothing
  Set list = Nothing

End Function


