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

Const VERSION = "2.0.0"

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
    Dim loadMult_F1 As New Dictionary
    Dim loadMult_F2 As New Dictionary
    Dim loadMult_F3 As New Dictionary
    Dim loadMult_M1 As New Dictionary
    Dim loadMult_M2 As New Dictionary
    Dim loadMult_M3 As New Dictionary
    Dim loadMults As New Dictionary
    Dim nodes As New Dictionary
    Dim forces As New Dictionary
    Dim sheetWithLoads As Excel.Worksheet
    Dim sheetWithFactors As Excel.Worksheet

    Set sheetWithLoads = book.Worksheets(LOADS_SHEET)
    Set sheetWithFactors = book.Worksheets(CORRESPONDANCE_SHEET)

    Set loadNames = getDicWithProperties(sheetWithFactors, "SAP_LOAD", "SCAD_LOAD")
    Set loadSafetyFactors = getDicWithProperties(sheetWithFactors, "SAP_LOAD", "SAFETY_FACTOR")

    Set loadMults.Item("F1") =  getDicWithProperties(sheetWithFactors, "SAP_LOAD", "MULT_F1")
    Set loadMults.Item("F2") =  getDicWithProperties(sheetWithFactors, "SAP_LOAD", "MULT_F2")
    Set loadMults.Item("F3") =  getDicWithProperties(sheetWithFactors, "SAP_LOAD", "MULT_F3")
    Set loadMults.Item("M1") =  getDicWithProperties(sheetWithFactors, "SAP_LOAD", "MULT_M1")
    Set loadMults.Item("M2") =  getDicWithProperties(sheetWithFactors, "SAP_LOAD", "MULT_M2")
    Set loadMults.Item("M3") =  getDicWithProperties(sheetWithFactors, "SAP_LOAD", "MULT_M3")
    

    Set nodes = getDicWithProperties(sheetWithFactors, "SAP_NODE", "SCAD_NODE")
    Set forces = getDicWithProperties(sheetWithFactors, "SAP_FORCE", "SCAD_FORCE")

    Dim colJoint As Integer
    Dim colNode As Integer
    Dim colSapLoad As Integer
    Dim colScadLoad As Integer
    Dim colForceLeft As Integer
    Dim colForceRight As Integer
    Dim sapLoad As String
    Dim row As Integer

    Dim xls As C_Excel_Worksheet
    Set xls = New C_Excel_Worksheet

    colJoint = xls.findColumnWithText(sheetWithLoads, LOADS_SHEET_ROW_WITH_NAMES, "JOINT")
    colSapLoad = xls.findColumnWithText(sheetWithLoads, LOADS_SHEET_ROW_WITH_NAMES, "SAP_LOAD")
    colNode = xls.findColumnWithText(sheetWithLoads, LOADS_SHEET_ROW_WITH_NAMES, "NODE")
    colScadLoad = xls.findColumnWithText(sheetWithLoads, LOADS_SHEET_ROW_WITH_NAMES, "SCAD_LOAD")

    row = LOADS_SHEET_ROW_WITH_NAMES + 1
    Do While sheetWithLoads.Cells(row, colJoint).Value <> ""

        sapLoad = sheetWithLoads.Cells(row, colSapLoad).Value
        sheetWithLoads.Cells(row, colNode).Value = nodes.Item(sheetWithLoads.Cells(row, colJoint).Value)
        sheetWithLoads.Cells(row, colScadLoad).Value = loadNames.Item(sapLoad)

        Dim key
        For Each key In forces.Keys
            colForceleft = xls.findColumnWithText(sheetWithLoads, LOADS_SHEET_ROW_WITH_NAMES, CStr(key))
            colForceRight = xls.findColumnWithText(sheetWithLoads, LOADS_SHEET_ROW_WITH_NAMES, forces.Item(key))

            sheetWithLoads.Cells(row, colForceRight).Value = sheetWithLoads.Cells(row, colForceLeft).Value * _
                loadSafetyFactors.Item(sapLoad) * loadMults.Item(CStr(key)).Item(sapLoad)
                ' loadMult_F1.Item(sapLoad) * _
                ' loadMult_F2.Item(sapLoad) * _
                ' loadMult_F3.Item(sapLoad) * _
                ' loadMult_M1.Item(sapLoad) * _
                ' loadMult_M2.Item(sapLoad) * _
                ' loadMult_M3.Item(sapLoad)
        Next
        
        row = row + 1
    Loop

    Set xls = Nothing
    Set sheetWithLoads = Nothing
    Set sheetWithFactors = Nothing
    Set loadNames =Nothing
    Set loadSafetyFactors = Nothing

    Set loadMult_F1 = Nothing
    Set loadMult_F2 = Nothing
    Set loadMult_F3 = Nothing
    Set loadMult_M1 = Nothing
    Set loadMult_M2 = Nothing
    Set loadMult_M3 = Nothing
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

End Sub
