VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Borehole_Database"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' База данных скважин
Private pDatabase As Dictionary

Private Sub Class_Initialize()
    Set pDatabase = New Dictionary
End Sub

Private Sub Class_Terminate()
    Set pDatabase = Nothing
End Sub

Public Sub addBorehole (ByVal brh As C_Borehole)
    Dim newBrh As C_Borehole
    Set newBrh = brh.DeepCopy()
    Set pDatabase.Item(newBrh.nameOfBorehole) = newBrh
End Sub

Public Function isExist(name As String) As Boolean
    isExist = pDatabase.Exists(name)
End Function

Public Function getBorehole(name As String) As C_Borehole
    if isExist(name) = True Then
        Set getBorehole = pDatabase.Item(name)
    Else
        MsgBox "There is not borehole " & name & " in database", vbOK, "ERROR"
    End If
End Function


Private Sub importBoreholeFromExcelSheet(sheet As Excel.Worksheet, ByVal row As Integer, ByVal col As Integer)
    Dim brh As New C_Borehole
    Dim name As String
    Dim value As Double

    Dim lastRow As Integer
    lastRow = sheet.Cells(sheet.Rows.count, col).End(xlUp).row

    Dim isRowWithLayer As Boolean
    isRowWithLayer = False

    brh.nameOfBorehole = CStr(sheet.Cells(row, col + 1))
    row = row + 1

    Do Until IsEmpty(sheet.Cells(row, col)) Or isRowWithLayer = True
    
        name = CStr(sheet.Cells(row, col).value)
        value = CDbl(sheet.Cells(row, col + 1).value)
        
        If name = "LAYERS" Then
            isRowWithLayer = True
        Else
            Select Case (name)
                Case "TOP": brh.topElevation = value
                Case "UGW": brh.waterDepth = value
            End Select
        End If
        
        row = row + 1
    Loop
    
    If isRowWithLayer = True Then
        Do Until row > lastRow
            Call brh.addLayer(CStr(sheet.Cells(row, col+1)), CDbl(sheet.Cells(row, col)))
            row = row + 1
        Loop
    End If

    Call addBorehole(brh)
    Set brh = Nothing
    
End Sub


Public Sub loadFromExcelSheet (sheet As Excel.Worksheet, rowWithNames As Integer)
    Dim column As Integer

    column = 1
    Do While sheet.Cells(rowWithNames, column).value <> "" 
        Call importBoreholeFromExcelSheet(sheet, rowWithNames, column)
        column = column + 2
    Loop
    
End Sub


