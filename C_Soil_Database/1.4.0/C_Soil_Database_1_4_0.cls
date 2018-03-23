VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Soil_Database"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const VERSION = "1.4.0"

' База данных грунтов в Excel
Private pDatabase As Dictionary

Private Sub Class_Initialize()
    Set pDatabase = New Dictionary
End Sub

Private Sub Class_Terminate()
    Set pDatabase = Nothing
End Sub

Public Sub addSoil (ByVal soil As C_Soil)
    Dim newSoil As C_Soil
    Set newSoil = soil.DeepCopy()
    Set pDatabase.Item(soil.name) = newSoil
End Sub

Public Function isExist(name As String) As Boolean
    isExist = pDatabase.Exists(name)
End Function

Public Function getSoil(name As String) As C_Soil
    if isExist(name) = True Then
        Set getSoil = pDatabase.Item(name)
    Else
        If name = "" Then 
            MsgBox "Пустое имя грунта", vbOK, "ERROR"
        Else
            MsgBox "В базе нет грунта " & name, vbOK, "ERROR"
        End If
    End If
End Function

Public Sub loadFromExcelSheet (sheet As Excel.Worksheet, rowWithNames As Integer)
    Dim xls As New C_Excel_Worksheet

    Dim requiredColumns As Variant

    ' Кол-во переносов д.б. не более 24
    requiredColumns = Array ( _
        "name", _
        "ClassOfSoil", _
        "TypeBySize", _
        "SubtypeBySize", _
        "TypeByLiquidityIndex", _
        "TypeByDensity", _
        "YoungModulus", _
        "YoungModulusSecondary", _
        "Yn", "Y2", "Y1", _
        "Cn", "C2", "C1", _
        "Fin", "Fi2", "Fi1", _
        "e", _
        "Sr", _
        "W", _
        "WL", "WP", _
        "IP", "IL", _
        "efh", "tfh", _
        "D", _
        "dfn", _
        "TypeByFrostHeaving", _
        "Ps", _
        "SeismicCategory" _
    )

    Dim i As Integer
    Dim col As Integer
    Dim columns As New Dictionary
    For i = LBound(requiredColumns) To UBound(requiredColumns)

        col = xls.findColumnWithText(sheet, rowWithNames, CStr(requiredColumns(i)))

        if col = 0 Then
            MsgBox "Обязательная колонка " & requiredColumns(i) & " НЕ найдена"
        Else
            columns.Item(requiredColumns(i)) = col
        End If

    Next i

    Dim row As Integer
    row = rowWithNames + 1
    Do While sheet.Cells(row, columns("name")).value <> ""

        Dim soil As New C_Soil
        soil.name = sheet.Cells(row, columns("name"))
        soil.ClassOfSoil = sheet.Cells(row, columns("ClassOfSoil"))
        soil.TypeBySize = sheet.Cells(row, columns("TypeBySize"))
        soil.SubtypeBySize = sheet.Cells(row, columns("SubtypeBySize"))
        soil.TypeByLiquidityIndex = sheet.Cells(row, columns("TypeByLiquidityIndex"))
        soil.TypeByDensity = sheet.Cells(row, columns("TypeByDensity"))
        soil.YoungModulus = sheet.Cells(row, columns("YoungModulus"))
        soil.YoungModulusSecondary = sheet.Cells(row, columns("YoungModulusSecondary"))
        soil.SoilDensity = sheet.Cells(row, columns("Yn"))
        soil.SoilDensity_1 = sheet.Cells(row, columns("Y1"))
        soil.SoilDensity_2 = sheet.Cells(row, columns("Y2"))
        soil.SpecificCohesion = sheet.Cells(row, columns("Cn"))
        soil.SpecificCohesion_1 = sheet.Cells(row, columns("C1"))
        soil.SpecificCohesion_2 = sheet.Cells(row, columns("C2"))
        soil.InternalFrictionAngle = sheet.Cells(row, columns("Fin"))
        soil.InternalFrictionAngle_1 = sheet.Cells(row, columns("Fi1"))
        soil.InternalFrictionAngle_2 = sheet.Cells(row, columns("Fi2"))
        soil.VoidRatio = sheet.Cells(row, columns("e"))
        soil.DegreeOfSaturation = sheet.Cells(row, columns("Sr"))
        soil.WaterContent = sheet.Cells(row, columns("W"))
        soil.PlasticLimit = sheet.Cells(row, columns("WP"))
        soil.LiquidLimit = sheet.Cells(row, columns("WL"))
        soil.PlasticityIndex = sheet.Cells(row, columns("IP"))
        soil.LiquidityIndex = sheet.Cells(row, columns("IL"))

        soil.FrostHeaveRate = sheet.Cells(row, columns("efh"))
        soil.ShearHeavingForce = sheet.Cells(row, columns("tfh"))
        soil.GranulationFactor = sheet.Cells(row, columns("D"))
        soil.FrozenDepth = sheet.Cells(row, columns("dfn"))
        soil.TypeByFrostHeaving = sheet.Cells(row, columns("TypeByFrostHeaving"))

        soil.SolidParticlesDensity = sheet.Cells(row, columns("Ps"))
        soil.SeismicCategory = sheet.Cells(Row, columns("SeismicCategory"))

        Call addSoil(soil)
        

        row = row + 1
    Loop

End Sub

Private Function getProperty(name As String, rowWithValues As Integer) As Variant
    Dim column As Integer
    Dim row As Integer
    
    Dim xls As Object
    Set xls = New C_Excel_Worksheet
    
    row = xls.findRowWithText(Application.ActiveSheet, 1, "NAMES")
    column = xls.findColumnWithText(Application.ActiveSheet, row, name)
    
    getProperty = Application.ActiveSheet.Cells(rowWithValues, column).value
End Function

Public Sub test()

    Call importLibraryList
    
    Dim xls As Object
    Set xls = New C_Excel_Worksheet
    
    Dim row As Integer
    row = xls.findRowWithText(Application.ActiveSheet, 1, "NAMES") + 1
    
    Do While ActiveSheet.Cells(row, 1).value <> ""
        testRow (row)
        row = row + 1
    Loop
    
End Sub





