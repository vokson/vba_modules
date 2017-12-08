Attribute VB_Name = "CPT_FUNCTIONS_VER001"
Option Explicit

Public CPTS As New Scripting.Dictionary

Const CPT_ROW_WITH_NAMES = 1
Const CPT_SHEET_NAME = "CPT"

Function getCptTopElevation(nameOfCpt As String) As Double
    Dim well As C_CPT
    Set well = getCptByName(nameOfCpt)

    getCptTopElevation = well.topElevation
End Function

Function getCptFrontResistanceAtDepth(nameOfCpt As String, depth As Double) As String
    Dim cpt As C_CPT
    Set cpt = getCptByName(nameOfCpt)

    getCptFrontResistanceAtDepth = cpt.getFrontResistanceAtDepth(depth)
End Function

Function getCptSideFrictionAtDepth(nameOfCpt As String, depth As Double) As String
    Dim cpt As C_CPT
    Set cpt = getCptByName(nameOfCpt)

    getCptSideFrictionAtDepth = cpt.getSideFrictionAtDepth(depth)
End Function

Public Function getCptByName(nameOfCpt As String) As C_CPT

    If CPTS.Exists(nameOfCpt) Then
        Set getCptByName = CPTS.Item(nameOfCpt)
        Exit Function
    End If
    
    Dim Rng As Range, columnWithDepth As Integer
    Set Rng = Worksheets(CPT_SHEET_NAME).Rows(1).Find(what:=nameOfCpt, LookIn:=xlValues, lookAt:=xlWhole, MatchCase:=True)
    
    If Not Rng Is Nothing Then
        
        columnWithDepth = Rng.column - 1
        
        Dim newCpt As New C_CPT, layerDepth() As Double, layerSideFriction() As Double, layerFrontResistance() As Double
        Dim depth As Double, sideFriction As Double, frontResistance As Double, row As Integer
        Dim name As String
        newCpt.name = nameOfCpt
        row = 1 + CPT_ROW_WITH_NAMES
        
        With Worksheets(CPT_SHEET_NAME)
        
            Dim isRowWithLayer As Boolean: isRowWithLayer = False
            Do Until IsEmpty(.Cells(row, columnWithDepth)) Or isRowWithLayer = True
            
                name = CStr(.Cells(row, columnWithDepth).value)
'                sideFriction = .Cells(row, columnWithDepth + 1).value
                
                If name = "LAYERS" Then
                    isRowWithLayer = True
                Else
                    Select Case (name)
                        Case "TOP": newCpt.topElevation = CDbl(.Cells(row, columnWithDepth + 1).value)
                    End Select
                End If
                
                row = row + 1
            Loop
            
            'Пропускаем желтую строку с названиями
            row = row + 1
            
            If isRowWithLayer = True Then
                Dim lastRow As Integer, i As Integer
                lastRow = .Cells(.Rows.count, columnWithDepth).End(xlUp).row
                
                ReDim layerDepth(lastRow - row): ReDim layerSideFriction(lastRow - row)
                ReDim layerFrontResistance(lastRow - row)
                
                For i = LBound(layerDepth) To UBound(layerDepth)
                    layerDepth(i) = CDbl(.Cells(row + i, columnWithDepth))
                    layerFrontResistance(i) = CDbl(.Cells(row + i, columnWithDepth + 1))
                    layerSideFriction(i) = CDbl(.Cells(row + i, columnWithDepth + 2))
                Next i
                
                newCpt.layerDepth = layerDepth
                newCpt.layerFrontResistance = layerFrontResistance
                newCpt.layerSideFriction = layerSideFriction
            End If
        
        End With
        
        Set getCptByName = newCpt
        Set CPTS.Item(nameOfCpt) = newCpt
        
    End If
    
End Function

Function getCptDepthArrayBtwDepth(nameOfCpt As String, depth1 As Double, depth2 As Double) As Double()

    Dim cpt As C_CPT
    Set cpt = getCptByName(nameOfCpt)

    getCptDepthArrayBtwDepth = cpt.getDepthArrayBtwDepth(depth1, depth2)
    
End Function

Function getCptSideFrictionArrayBtwDepth(nameOfCpt As String, depth1 As Double, depth2 As Double) As Double()

    Dim cpt As C_CPT
    Set cpt = getCptByName(nameOfCpt)

    getCptSideFrictionArrayBtwDepth = cpt.getSideFrictionArrayBtwDepth(depth1, depth2)
    
End Function

Function getCptFrontResistanceArrayBtwDepth(nameOfCpt As String, depth1 As Double, depth2 As Double) As Double()

    Dim cpt As C_CPT
    Set cpt = getCptByName(nameOfCpt)

    getCptFrontResistanceArrayBtwDepth = cpt.getFrontResistanceArrayBtwDepth(depth1, depth2)
    
End Function

Sub testGetCptArrayBtwDepth()
    Dim a, b, c, i As Integer
    Dim name As String, d1 As Double, d2 As Double

    name = "CPT-148I": d1 = 0.99: d2 = 1.21
    
    a = getCptDepthArrayBtwDepth(name, d1, d2)
    b = getCptFrontResistanceArrayBtwDepth(name, d1, d2)
    c = getCptSideFrictionArrayBtwDepth(name, d1, d2)
    
    Dim sizeA As Integer, sizeB As Integer, sizeC As Integer
    sizeA = UBound(a) - LBound(a)
    sizeB = UBound(a) - LBound(a)
    sizeC = UBound(a) - LBound(a)
    
    If (sizeA = sizeB And sizeB = sizeC) Then
        Debug.Print "IS ALL ARRAYS WITH SAME SIZE - OK"
    Else
        Debug.Print "IS ALL ARRAYS WITH SAME SIZE - FAIL"
    End If
    
    For i = LBound(a) To UBound(a)
        Debug.Print a(i) & " - " & b(i) & " - " & c(i)
    Next i
End Sub
