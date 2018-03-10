VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_SCAD_Document4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Make string from array of dictionaries with nodes
' Format of node
' X As Double
' Y As Double
' Z As Double
Function writeNodes(ByRef nodes() As Dictionary) As string
  Dim i As Long
  Dim node As Dictionary
  Dim x As Double
  Dim y As Double
  Dim z As Double
  Dim text As String

  text = "(4/"

  For i=LBound(nodes) To UBound(nodes)
    Set node = nodes(i)
    text = text & node.Item("X") & " " & node.Item("Y") & " " & node.Item("Z") & "/"
  Next i

  text = text & ")"
  writeNodes = text

End Function


