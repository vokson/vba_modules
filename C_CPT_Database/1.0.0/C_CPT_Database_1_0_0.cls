VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_CPT_Database"
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

Public Sub addCpt (ByVal cpt As C_CPT)
    Dim newCpt As C_CPT
    Set newCpt = cpt.DeepCopy()
    Set pDatabase.Item(newCpt.nameOfBorehole) = newCpt
End Sub

Public Function isExist(name As String) As Boolean
    isExist = pDatabase.Exists(name)
End Function

Public Function getCpt(name As String) As C_CPT
    if isExist(name) = True Then
        Set getCpt = pDatabase.Item(name)
    Else
        MsgBox "There is not CPT " & name & " in database", vbOK, "ERROR"
    End If
End Function


Private Sub importCptFromExcelSheet(sheet As Excel.Worksheet, ByVal row As Integer, ByVal col As Integer)
    Dim cpt As New C_CPT
    Dim name As String
    Dim value As Double

    Dim lastRow As Integer
    lastRow = sheet.Cells(sheet.Rows.count, col).End(xlUp).row

    Dim isRowWithLayer As Boolean
    isRowWithLayer = False

    cpt.nameOfBorehole = CStr(sheet.Cells(row, col + 1))
    row = row + 1

    Do Until IsEmpty(sheet.Cells(row, col)) Or isRowWithLayer = True
    
        name = CStr(sheet.Cells(row, col).value)
        value = CDbl(sheet.Cells(row, col + 1).value)
        
        If name = "LAYERS" Then
            isRowWithLayer = True
            row = row + 1 ' Желтая строка с заголовками
        Else
            Select Case (name)
                Case "TOP": cpt.topElevation = value
            End Select
        End If
        
        row = row + 1
    Loop
    
    If isRowWithLayer = True Then
        Do Until row > lastRow
            Call cpt.addLayer( _
                CDbl(sheet.Cells(row, col)), _
                CDbl(sheet.Cells(row, col+1)), _
                CDbl(sheet.Cells(row, col+2)) _
            )
            row = row + 1
        Loop
    End If

    Call addCpt(cpt)
    Set cpt = Nothing
    
End Sub


Public Sub loadFromExcelSheet (sheet As Excel.Worksheet, rowWithNames As Integer)
    Dim column As Integer

    column = 1
    Do While sheet.Cells(rowWithNames, column).value <> "" 
        Call importCptFromExcelSheet(sheet, rowWithNames, column)
        column = column + 3
    Loop
    
End Sub


