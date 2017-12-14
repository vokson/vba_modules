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
Function makeArrayFromList(s As String) As Variant
    
    Dim str As New C_String
    
    Dim text As String
    text = Trim(s)
    text = str.replaceNewStringSymbolWith(text, " ")
    text = str.deleteDoubleSpaces(text)

    Dim textArray() As String
    textArray = Split(text)

    Dim elements() As Long

    Dim i As Long
    For i = LBound(textArray) To UBound(textArray)

    Next i

End Function

