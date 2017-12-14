VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_SCAD_List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Make array of elements' numbers from standard SCAD list
' For example "7538 r 7610 8" => Array(7538, 7546, ..., 7602, 7610)
Function makeArrayFromList(s As String) As Long()
  Dim math As New C_Math
  
  Dim result() As Long
  Dim additive() As Long
'  result = Array()

  Dim str As New C_String
  
  Dim text As String
  text = Trim(s)
  text = str.replaceNewStringSymbolWith(text, " ")
  text = str.deleteDoubleSpaces(text)

  Dim textArray() As String
  textArray = Split(text)

  Dim elements() As Long

  Dim i As Long: i = LBound(textArray)
  Dim pos As Integer
  Dim pitch As Long
  Dim min As Long
  Dim max As Long

  Do Until i > UBound(textArray)
    
    If InStr(1, textArray(i), "-", vbBinaryCompare) > 0 Then
      pos = InStr(1, textArray(i), "-", vbBinaryCompare)
      min = CLng(Left(textArray(i), pos - 1))
      max = CLng(Mid(textArray(i), pos + 1))
      additive = math.makeArrayWithStep(min, max, 1)
      i = i + 1

    ElseIf textArray(i) = "r" Then
      min = result(UBound(result))
      max = CLng(result(i + 1))
      pitch = CLng(result(i + 2))
      additive = math.makeArrayWithStep(min + pitch, max, pitch)
      i = i + 3

    Else
      additive = Array(CLng(textArray(i)))
      i = i + 1

    End If
    
    If (Not Not result) <> 0 Then
        result = math.mergeTwoArrays(result, additive)
    Else
        result = additive
    End If

  Loop

  makeArrayFromList = result

  Set math = Nothing

End Function


