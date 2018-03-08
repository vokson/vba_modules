VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_SAP2000_SCAD_Converter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const LOADS_SHEET = "LOADS"
Const CORRESPONDANCE_SHEET = "CORRESPONDANCE"
Const LOADS_SHEET_ROW_WITH_NAMES= 2
Const CORRESPONDANCE_SHEET_ROW_WITH_NAMES= 1

Private Function getDicWithProperties(sheetWithLoads As Excel.Worksheet, name1 As String, name2 As String) As Dictionary
    Dim xls As C_Excel_Worksheet
    Set xls = New C_Excel_Worksheet

    Dim col1 As Integer
    Dim col2 As Integer
    Dim row As Integer
    col1 = xls.findColumnWithText(sheetWithLoads, CORRESPONDANCE_SHEET_ROW_WITH_NAMES, name1)
    col2 = xls.findColumnWithText(sheetWithLoads, CORRESPONDANCE_SHEET_ROW_WITH_NAMES, name2)

    row = CORRESPONDANCE_SHEET_ROW_WITH_NAMES + 1
    Dim dic As New Dictionary
    Do While sheetWithLoads.Cells(row, col1).Value <> ""
        dic.item(sheetWithLoads.Cells(row, col1).Value) = sheetWithLoads.Cells(row, col2).Value
        row = row + 1
    Loop

    Set getDicWithProperties = dic
    Set xls = Nothing
    Set dic = Nothing
End Function

Public Sub calculateSCADloads(book As Excel.Workbook)

    Dim loadNames As New Dictionary
    Dim loadSafetyFactors As New Dictionary
    Dim loadMults As New Dictionary
    Dim nodes As New Dictionary
    Dim sheetWithLoads As Excel.Worksheet
    Dim sheetWithFactors As Excel.Worksheet

    Set sheetWithLoads = book.Worksheets(LOADS_SHEET)
    Set sheetWithFactors = book.Worksheets(CORRESPONDANCE_SHEET)

    Set loadNames = getDicWithProperties(sheetWithFactors, "SAP_LOAD", "SCAD_LOAD")
    Set loadSafetyFactors = getDicWithProperties(sheetWithFactors, "SAP_LOAD", "SAFETY_FACTOR")
    Set loadMults = getDicWithProperties(sheetWithFactors, "SAP_LOAD", "MULT")
    Set nodes = getDicWithProperties(sheetWithFactors, "SAP_NODE", "SCAD_NODE")

    Dim colJoint As Integer
    Dim colNode As Integer
    Dim colSapLoad As Integer
    Dim colScadLoad As Integer
    Dim colF1 As Integer
    Dim colF2 As Integer
    Dim colF3 As Integer
    Dim colM1 As Integer
    Dim colM2 As Integer
    Dim colM3 As Integer
    Dim colFX As Integer
    Dim colFY As Integer
    Dim colFZ As Integer
    Dim colMX As Integer
    Dim colMY As Integer
    Dim colMZ As Integer
    Dim row As Integer

    Dim xls As C_Excel_Worksheet
    Set xls = New C_Excel_Worksheet

    colJoint = xls.findColumnWithText(sheetWithLoads, LOADS_SHEET_ROW_WITH_NAMES, "JOINT")
    colSapLoad = xls.findColumnWithText(sheetWithLoads, LOADS_SHEET_ROW_WITH_NAMES, "SAP_LOAD")

    colF1 = xls.findColumnWithText(sheetWithLoads, LOADS_SHEET_ROW_WITH_NAMES, "F1")
    colF2 = xls.findColumnWithText(sheetWithLoads, LOADS_SHEET_ROW_WITH_NAMES, "F2")
    colF3 = xls.findColumnWithText(sheetWithLoads, LOADS_SHEET_ROW_WITH_NAMES, "F3")
    colM1 = xls.findColumnWithText(sheetWithLoads, LOADS_SHEET_ROW_WITH_NAMES, "M1")
    colM2 = xls.findColumnWithText(sheetWithLoads, LOADS_SHEET_ROW_WITH_NAMES, "M2")
    colM3 = xls.findColumnWithText(sheetWithLoads, LOADS_SHEET_ROW_WITH_NAMES, "M3")

    colNode = xls.findColumnWithText(sheetWithLoads, LOADS_SHEET_ROW_WITH_NAMES, "NODE")
    colScadLoad = xls.findColumnWithText(sheetWithLoads, LOADS_SHEET_ROW_WITH_NAMES, "SCAD_LOAD")
    colFX = xls.findColumnWithText(sheetWithLoads, LOADS_SHEET_ROW_WITH_NAMES, "FX")
    colFY = xls.findColumnWithText(sheetWithLoads, LOADS_SHEET_ROW_WITH_NAMES, "FY")
    colFZ = xls.findColumnWithText(sheetWithLoads, LOADS_SHEET_ROW_WITH_NAMES, "FZ")
    colMX = xls.findColumnWithText(sheetWithLoads, LOADS_SHEET_ROW_WITH_NAMES, "MX")
    colMY = xls.findColumnWithText(sheetWithLoads, LOADS_SHEET_ROW_WITH_NAMES, "MY")
    colMZ = xls.findColumnWithText(sheetWithLoads, LOADS_SHEET_ROW_WITH_NAMES, "MZ")

    row = LOADS_SHEET_ROW_WITH_NAMES + 1
    Do While sheetWithLoads.Cells(row, colJoint).Value <> ""

        sheetWithLoads.Cells(row, colNode).Value = nodes.Item(sheetWithLoads.Cells(row, colJoint).Value)
        sheetWithLoads.Cells(row, colScadLoad).Value = loadNames.Item(sheetWithLoads.Cells(row, colSapLoad).Value)

        sheetWithLoads.Cells(row, colFX).Value = sheetWithLoads.Cells(row, colF1).Value * _
            loadSafetyFactors.Item(sheetWithLoads.Cells(row, colSapLoad).Value) * _
            loadMults.Item(sheetWithLoads.Cells(row, colSapLoad).Value)

        sheetWithLoads.Cells(row, colFY).Value = sheetWithLoads.Cells(row, colF2).Value * _
            loadSafetyFactors.Item(sheetWithLoads.Cells(row, colSapLoad).Value) * _
            loadMults.Item(sheetWithLoads.Cells(row, colSapLoad).Value)

        sheetWithLoads.Cells(row, colFZ).Value = sheetWithLoads.Cells(row, colF3).Value * _
            loadSafetyFactors.Item(sheetWithLoads.Cells(row, colSapLoad).Value) * _
            loadMults.Item(sheetWithLoads.Cells(row, colSapLoad).Value)

        sheetWithLoads.Cells(row, colMX).Value = sheetWithLoads.Cells(row, colM1).Value * _
            loadSafetyFactors.Item(sheetWithLoads.Cells(row, colSapLoad).Value) * _
            loadMults.Item(sheetWithLoads.Cells(row, colSapLoad).Value)

        sheetWithLoads.Cells(row, colMY).Value = sheetWithLoads.Cells(row, colM2).Value * _
            loadSafetyFactors.Item(sheetWithLoads.Cells(row, colSapLoad).Value) * _
            loadMults.Item(sheetWithLoads.Cells(row, colSapLoad).Value)

        sheetWithLoads.Cells(row, colMZ).Value = sheetWithLoads.Cells(row, colM3).Value * _
            loadSafetyFactors.Item(sheetWithLoads.Cells(row, colSapLoad).Value) * _
            loadMults.Item(sheetWithLoads.Cells(row, colSapLoad).Value)

        row = row + 1
    Loop

    Set xls = Nothing
    Set sheetWithLoads = Nothing
    Set sheetWithFactors = Nothing
    Set loadNames =Nothing
    Set loadSafetyFactors = Nothing
    Set loadMults = Nothing
    Set nodes = Nothing
    
End Sub



Public Sub writeScadLoadsToTextFile(book As Excel.Workbook)

    Dim colNode As Integer
    Dim colScadLoad As Integer
    Dim colFX As Integer
    Dim colFY As Integer
    Dim colFZ As Integer
    Dim colMX As Integer
    Dim colMY As Integer
    Dim colMZ As Integer
    Dim row As Integer
    

    Dim xls As C_Excel_Worksheet
    Set xls = New C_Excel_Worksheet

    Dim sheetWithLoads As Excel.Worksheet
    Set sheetWithLoads = book.Worksheets(LOADS_SHEET)

    colNode = xls.findColumnWithText(sheetWithLoads, LOADS_SHEET_ROW_WITH_NAMES, "NODE")
    colScadLoad = xls.findColumnWithText(sheetWithLoads, LOADS_SHEET_ROW_WITH_NAMES, "SCAD_LOAD")
    colFX = xls.findColumnWithText(sheetWithLoads, LOADS_SHEET_ROW_WITH_NAMES, "FX")
    colFY = xls.findColumnWithText(sheetWithLoads, LOADS_SHEET_ROW_WITH_NAMES, "FY")
    colFZ = xls.findColumnWithText(sheetWithLoads, LOADS_SHEET_ROW_WITH_NAMES, "FZ")
    colMX = xls.findColumnWithText(sheetWithLoads, LOADS_SHEET_ROW_WITH_NAMES, "MX")
    colMY = xls.findColumnWithText(sheetWithLoads, LOADS_SHEET_ROW_WITH_NAMES, "MY")
    colMZ = xls.findColumnWithText(sheetWithLoads, LOADS_SHEET_ROW_WITH_NAMES, "MZ")

    Dim node As Long
    Dim load As Integer
    Dim FX As Double
    Dim FY As Double
    Dim FZ As Double
    Dim MX As Double
    Dim MY As Double
    Dim MZ As Double

On Error GoTo Handler
    
    Dim myFile As String
    myFile = Application.ActiveWorkbook.Path & "\output.txt"
    Open myFile For Output As #1

    row = LOADS_SHEET_ROW_WITH_NAMES + 1
    Do While sheetWithLoads.Cells(row, colNode).Value <> ""

        node = CLng(sheetWithLoads.Cells(row, colNode).Value)
        load = CInt(sheetWithLoads.Cells(row, colScadLoad).Value)
        FX = CDbl(sheetWithLoads.Cells(row, colFX).Value)
        FY = CDbl(sheetWithLoads.Cells(row, colFY).Value)
        FZ = CDbl(sheetWithLoads.Cells(row, colFZ).Value)
        MX = CDbl(sheetWithLoads.Cells(row, colMX).Value)
        MY = CDbl(sheetWithLoads.Cells(row, colMY).Value)
        MZ = CDbl(sheetWithLoads.Cells(row, colMZ).Value)
    
        Write #1, "0 1 Load=" & load & "  " & FX & ":" & node & " /"
        Write #1, "0 2  " & FY & ":" & node & " /"
        Write #1, "0 3  " & FZ & ":" & node & " /"
        Write #1, "0 4  " & MX & ":" & node & " /"
        Write #1, "0 5  " & MY & ":" & node & " /"
        Write #1, "0 6  " & MZ & ":" & node & " /"
        Write #1, ""
        row = row + 1
    Loop
    
Handler:
    Close #1

        
    
'     Dim node As Long, load As String, c As Integer, i As Long, r As Long
'     Dim oRange As Range
'     Set oRange = Selection
    
    
'     Dim LOADS As New CImprovedCollection, oLoad As CLoad, oCol As CImprovedCollection
    
'     For r = 1 To oRange.Rows.Count
    
'         load = CStr(oRange.Cells(r, 2).Value)
        
'         If Not LOADS.Exists(load) Then
'             Set oCol = New CImprovedCollection
'             LOADS.Add oCol, load
'         End If
        
'         node = CLng(oRange.Cells(r, 1).Value)
'         Select Case (node)
'             Case 338: node = 1239
'             Case 359: node = 1240
'             Case 395: node = 1241
'             Case 428: node = 1242
'             Case 431: node = 1243
'             Case 347: node = 1004
'             Case 348: node = 999
'             Case 374: node = 1110
'             Case 389: node = 990
'             Case 423: node = 1019
'             Case 424: node = 992
'             Case 429: node = 1024
'             Case 430: node = 1029
'             Case 432: node = 1149
'             Case 433: node = 1167
'             Case 434: node = 1158
'             Case 435: node = 1176
'             Case 436: node = 1034
'             Case 437: node = 1039
'             Case 438: node = 1115
'             Case 439: node = 1120
'             Case 441: node = 1125
'             Case 447: node = 1130
'             Case 453: node = 1140
'             Case 459: node = 1135
'             Case 464: node = 1044
'             Case 465: node = 1049
'             Case 466: node = 1246
'             Case 467: node = 1247
'             Case 468: node = 1054
'             Case 469: node = 1059
'             Case 470: node = 1212
'             Case 471: node = 1185
'             Case 472: node = 5874
'             Case 473: node = 5865
'             Case 474: node = 1064
'             Case 475: node = 1069
'             Case 476: node = 6231
'             Case 477: node = 6236
'             Case 478: node = 6359
'             Case 479: node = 6364
'             Case 480: node = 1222
'             Case 481: node = 1217
'             Case Else: Debug.Print "Wrong node number " & node & " !!!"
'         End Select
        
'         Set oLoad = New CLoad
'         oLoad.load = load
'         oLoad.node = node
'         oLoad.FX = CDbl(oRange.Cells(r, 3).Value)
'         oLoad.FY = CDbl(oRange.Cells(r, 4).Value)
'         oLoad.FZ = CDbl(oRange.Cells(r, 5).Value)
'         oLoad.MX = CDbl(oRange.Cells(r, 6).Value)
'         oLoad.MY = CDbl(oRange.Cells(r, 7).Value)
'         oLoad.MZ = CDbl(oRange.Cells(r, 8).Value)
        
'         LOADS.item(load).Add oLoad
'         Debug.Print r
'     Next r
    
'     Dim item1, item2, isFirst As Boolean
    
'     On Error GoTo Handler
    
'     Dim myFile As String
'     myFile = Application.ActiveWorkbook.Path & "\output.txt"
'     Open myFile For Output As #1
    
    
'     For Each item1 In LOADS
'         isFirst = True
    
'         For Each item2 In item1
'             If isFirst = True Then
'                 Write #1, "0 1 Load=" & item2.load & "  " & item2.FX & ":" & item2.node & " /"
'                 Write #1, "0 2  " & item2.FY & ":" & item2.node & " /"
'                 Write #1, "0 3  " & item2.FZ & ":" & item2.node & " /"
'                 Write #1, "0 4  " & item2.MX & ":" & item2.node & " /"
'                 Write #1, "0 5  " & item2.MY & ":" & item2.node & " /"
'                 Write #1, "0 6  " & item2.MZ & ":" & item2.node & " /"
'                 isFirst = False
'             Else
'                 Write #1, "0 1  " & item2.FX & ":" & item2.node & " /"
'                 Write #1, "0 2  " & item2.FY & ":" & item2.node & " /"
'                 Write #1, "0 3  " & item2.FZ & ":" & item2.node & " /"
'                 Write #1, "0 4  " & item2.MX & ":" & item2.node & " /"
'                 Write #1, "0 5  " & item2.MY & ":" & item2.node & " /"
'                 Write #1, "0 6  " & item2.MZ & ":" & item2.node & " /"
'             End If
'         Next
        
'         Write #1, ""
'     Next
    
' Handler:
'     Close #1
    
End Sub
