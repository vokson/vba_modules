Attribute VB_Name = "BOREHOLES_FUNCTIONS_VER001"
Option Explicit

Public BOREHOLES As New Scripting.Dictionary

Const BOREHOLE_ROW_WITH_NAMES = 1
Const BOREHOLE_SHEET_NAME = "BRH"

Function getBoreholeTopElevation(nameOfBorehole As String) As Double
    Dim brh As borehole
    Set brh = getBoreholeByName(nameOfBorehole)

    getBoreholeTopElevation = brh.topElevation
End Function

Function getBoreholeSoilNameAtDepth(nameOfBorehole As String, depth As Double) As String
    Dim brh As borehole
    Set brh = getBoreholeByName(nameOfBorehole)

    getBoreholeSoilNameAtDepth = brh.getSoilNameAtDepth(depth)
End Function

Public Function getBoreholeByName(nameOfBorehole As String) As borehole

    If BOREHOLES.Exists(nameOfBorehole) Then
        Set getBoreholeByName = BOREHOLES.Item(nameOfBorehole)
        Exit Function
    End If
    
    Dim Rng As Range, columnWithNames As Integer
    Set Rng = Worksheets(BOREHOLE_SHEET_NAME).Rows(1).Find(what:=nameOfBorehole, LookIn:=xlValues, lookAt:=xlWhole, MatchCase:=True)
    
    If Not Rng Is Nothing Then
        
        columnWithNames = Rng.column - 1
        
        Dim newBorehole As New borehole, layerDepth() As Double, layerName() As String
        Dim name As String, value As Variant, row As Integer
        newBorehole.name = nameOfBorehole
        row = 1 + BOREHOLE_ROW_WITH_NAMES
        
        With Worksheets(BOREHOLE_SHEET_NAME)
        
            Dim isRowWithLayer As Boolean: isRowWithLayer = False
            Do Until IsEmpty(.Cells(row, columnWithNames)) Or isRowWithLayer = True
            
                name = .Cells(row, columnWithNames).value
                value = .Cells(row, columnWithNames + 1).value
                
                If name = "LAYERS" Then
                    isRowWithLayer = True
                Else
                    Select Case (name)
                        Case "TOP": newBorehole.topElevation = CDbl(value)
                        Case "UGW": newBorehole.waterDepth = CDbl(value)
                    End Select
                End If
                
                row = row + 1
            Loop
            
            If isRowWithLayer = True Then
                Dim lastRow As Integer, i As Integer
                lastRow = .Cells(.Rows.count, columnWithNames).End(xlUp).row
                
                ReDim layerDepth(lastRow - row): ReDim layerName(lastRow - row)
                
                For i = LBound(layerDepth) To UBound(layerDepth)
                    layerDepth(i) = .Cells(row + i, columnWithNames)
                    layerName(i) = .Cells(row + i, columnWithNames + 1)
                Next i
                
                newBorehole.layerDepth = layerDepth
                newBorehole.layerName = layerName
            End If
        
        End With
        
        Set getBoreholeByName = newBorehole
        Set BOREHOLES.Item(nameOfBorehole) = newBorehole
        
    End If
    
End Function
