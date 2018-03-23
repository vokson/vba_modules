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
Function getGroups(s As String, code As Integer) As Dictionary
  Dim result As New Dictionary
  
  Dim math As New C_Math
  Dim str As New C_String
  Dim list As New C_SCAD_List
  
  Dim text As String
  text = Trim(s)
  text = str.replaceNewStringSymbolWith(text, " ")
  text = str.deleteDoubleSpaces(text)

  text = Mid(text, InStr(1, text, "(" & CStr(code) & "/", vbBinaryCompare) + 4)
  text = Left(text, InStr(1, text, ")", vbBinaryCompare) - 1)
  If Mid(text, Len(text) - 1, 1) = "/" Then text = Left(text, Len(text) - 2) Else text = Left(text, Len(text) - 1)

  Dim groupArray() As String
  groupArray = Split(text, "/")

  Dim textArray() As String
  Dim i As Integer

  ' Сливаем группы обратно, если в имени группы был слеш "/"
  For i = LBound(groupArray) To UBound(groupArray)

    groupArray(i) = Trim(groupArray(i))

    If (Not Not textArray) <> 0 Then ' Array has been initialized, so you're good to go.

      If "Name=" = Left(groupArray(i), 5) Then
        Redim Preserve textArray(UBound(textArray) + 1)
        textArray(UBound(textArray)) = groupArray(i)
      Else 
        textArray(UBound(textArray)) = textArray(UBound(textArray)) & "/" & groupArray(i)
      End If

    Else ' Array IS NOT initialized
      Redim textArray(0)
      textArray(0) = groupArray(i)
    End If

  Next i




  
  Dim groupText As String
  Dim nameText As String
  Dim elementsText As String
  For i = LBound(textArray) To UBound(textArray)
    groupText = textArray(i)

    nameText = Mid(groupText, InStr(1, groupText, """", vbBinaryCompare) + 1)
    nameText = Left(nameText, InStr(1, nameText, """", vbBinaryCompare) - 1)

    elementsText = Mid(groupText, InStr(1, groupText, ":", vbBinaryCompare) + 1)

    result.Item(nameText) = list.makeArrayFromList(elementsText)
  Next i

  Set getGroups = result

  Set math = Nothing
  Set str = Nothing
  Set list = Nothing

End Function


