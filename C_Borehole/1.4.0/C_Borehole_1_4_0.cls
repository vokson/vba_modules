VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Borehole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const VERSION = "1.4.0"

Private math As C_Math
Private pName As String
Private pTopElevation As Double 'Абсолютная отметка (м)
Private pWaterDepth As Double ' Глубина воды (м)
Private pLayerDepth() As Double ' Глубина низа слоя грунта (м)
Private pLayerName() As String ' Имя слоя грунта (м)
Private pCountOfLayers As Integer

Private Sub Class_Initialize() 
    Set math = New C_Math
    pCountOfLayers = 0
    pWaterDepth = -1000
End Sub 

' Имя скважины
Property Let nameOfBorehole(word As String)
    pName = word
End Property

Property Get nameOfBorehole() As String
    nameOfBorehole = pName
End Property

' Отметка устья скважины
Property Let topElevation(elevation As Double)
     pTopElevation = elevation
End Property

Property Get topElevation() As Double
    topElevation = pTopElevation
End Property

' Уровень подземных вод
Property Let waterDepth(depth As Double)
     pWaterDepth  = depth
End Property

Property Get waterDepth() As Double
    waterDepth = pWaterDepth
End Property

Property Get countOfLayers() As Integer
    countOfLayers = pCountOfLayers
End Property

Public Function arrayOfLayerDepth() As Variant
    arrayOfLayerDepth = pLayerDepth
End Function

Public Function arrayOfLayerName() As Variant
    arrayOfLayerName = pLayerName
End Function

' Добавление слоя
Public Sub addLayer(name As String, bottomOfLayer As Double)

    If bottomOfLayer <=0 Then
        MsgBox "C_Borehole - Низ слоя " & name & " должен быть > 0. Скважина " & nameOfBorehole, vbOK, "ERROR"
        Exit Sub
    End If

    ReDim Preserve pLayerDepth(pCountOfLayers)
    ReDim Preserve pLayerName(pCountOfLayers)
    pCountOfLayers = pCountOfLayers + 1

    pLayerDepth(UBound(pLayerDepth)) = bottomOfLayer
    pLayerName(UBound(pLayerName)) = name

End Sub

' Получить имя грунта на определенной глубине
Public Property Get getSoilNameAtDepth(depth As Double) As String
    
    If depth >=  pLayerDepth(UBound(pLayerDepth)) Then
        MsgBox "C_Borehole - Увеличьте глубину скважины. Глубина " _
             & nameOfBorehole & " д.б. > " & depth, vbOK, "ERROR"
        Exit Sub
    End If

    Dim i As Integer
    For i = LBound(pLayerDepth) To UBound(pLayerDepth)
        If pLayerDepth(i) > depth Then
            getSoilNameAtDepth = pLayerName(i)
            Exit For
        End If
    Next i
    
End Property

' Есть ли вода на данной глубине
Public Property Get isWaterAtDepth(depth As Double) As Boolean
    
    If (depth >= pWaterDepth) Then
        isWaterAtDepth = True
    Else
        isWaterAtDepth = False
    End If
    
End Property

Public Function DeepCopy() As C_Borehole
    Dim other As New C_Borehole

    other.nameOfBorehole = Me.nameOfBorehole
    other.topElevation = Me.topElevation
    other.waterDepth = Me.waterDepth

    If (Not Not pLayerDepth) <> 0 Then ' Array has been initialized, so you're good to go.

        Dim i As Integer
        For i = LBound(pLayerDepth) To UBound(pLayerDepth)
            Call other.addLayer(pLayerName(i), pLayerDepth(i))
        Next i

    End if

    Set DeepCopy = other
End Function

Public Function DeepCopyWithOtherTopElevation(newTopElevation As Double, nameOfFillSoil As String)
    
    If (Not Not pLayerDepth) <> 0 Then ' Array has been initialized, so you're good to go.

        Dim other As New C_Borehole
        Dim deltaTop As Double
        Dim i As Integer

        deltaTop = Abs(newTopElevation-Me.topElevation)
        other.nameOfBorehole = Me.nameOfBorehole
        other.topElevation = newTopElevation

        If newTopElevation > Me.topElevation Then

            other.waterDepth = Me.waterDepth + deltaTop

            Call other.addLayer(nameOfFillSoil, deltaTop)

            
            For i = LBound(pLayerDepth) To UBound(pLayerDepth)
                Call other.addLayer(pLayerName(i), pLayerDepth(i) + deltaTop)
            Next i

            Set DeepCopyWithOtherTopElevation = other

        ElseIf newTopElevation < Me.topElevation Then

            other.waterDepth = math.max(0, Me.waterDepth - deltaTop)

            For i = LBound(pLayerDepth) To UBound(pLayerDepth)
                if pLayerDepth(i) > deltaTop Then
                    Call other.addLayer(pLayerName(i), pLayerDepth(i) - deltaTop)
                End If
            Next i

            Set DeepCopyWithOtherTopElevation = other

        Else
            Set DeepCopyWithOtherTopElevation = Me.DeepCopy()
        End If

    End if
    
End Function

Public Function isSameWith(other As C_Borehole) As Boolean

    isSameWith = True

    Dim precision As Double
    precision = 0.001


    If (other.nameOfBorehole <> Me.nameOfBorehole) Then isSameWith = False
    If (Abs(other.topElevation - Me.topElevation) > precision) Then isSameWith = False
    If (Abs(other.waterDepth - Me.waterDepth) > precision) Then isSameWith = False

    If (Not Not pLayerDepth) <> 0 Then ' Array has been initialized, so you're good to go.

        if pCountOfLayers <> other.countOfLayers Then
            isSameWith = False
            Exit Function
        End If
        
        Dim otherLayerDepthArray As Variant
        Dim otherLayerNameArray As Variant

        otherLayerDepthArray = other.arrayOfLayerDepth
        otherLayerNameArray = other.arrayOfLayerName

        If Not (math.isArraysSame(otherLayerDepthArray, pLayerDepth) And _
             math.isArraysSame(otherLayerNameArray, pLayerName)) Then
            
            isSameWith = False
            Exit Function

        End If

    End if

End Function

' Разделяет слой грунта на указанной глубине на 2шт
' Если указанная глубина совпадает с границей слоя в
' Пределах tolerance - ничего не делает
Public Sub splitAtDepth(depth As Double, Optional tolerance As Double = 0.01)

    if depth > pLayerDepth(UBound(pLayerDepth)) Then
        Exit Sub
    End If

    Dim i As Integer
    For i = LBound(pLayerDepth) To UBound(pLayerDepth)
        If Abs(depth-pLayerDepth(i)) < tolerance Then
            Exit Sub
        End If

        If pLayerDepth(i)-depth > tolerance Then
            Exit For
        End If
    Next i

    ReDim Preserve pLayerDepth(LBound(pLayerDepth) To UBound(pLayerDepth)+1)
    ReDim Preserve pLayerName(LBound(pLayerName) To UBound(pLayerName)+1)

    Dim k As Integer
    For k = UBound(pLayerDepth)-1 To i Step -1
        pLayerDepth(k+1) = pLayerDepth(k)
        pLayerName(k+1) = pLayerName(k)
    Next k

    pLayerDepth(i) = depth
    pCountOfLayers = UBound(pLayerDepth)-LBound(pLayerDepth)+1

End Sub

' Разделяет слой грунта на 2шт на глубине UGW
Public Sub splitAtWaterDepth()
    If pWaterDepth <> -1000 And pWaterDepth <> 0 Then Call splitAtDepth(pWaterDepth)
End Sub

' Разделяет указанный слой грунта на слои указанной толщины
' layerIndex начинается с 0
' Не оставляет слои тоньше tolerance
Public Sub splitLayerWithMaxThickness(layerIndex As Integer, maxThickness As Double, _
                                         Optional tolerance As Double = 0.01)
    
    If layerIndex > UBound(pLayerDepth) Then Exit Sub
    If maxThickness > pLayerDepth(layerIndex) Or  _
        Abs(pLayerDepth(layerIndex)-maxThickness) <= tolerance Then Exit Sub

    Dim depthOfTop As Double
    if (layerIndex = LBound(pLayerDepth)) Then depthOfTop = 0 Else depthOfTop = pLayerDepth(layerIndex-1)

    Dim countOfLayers As Integer
    countOfLayers = Fix(Me.thicknessOfLayer(layerIndex) / maxThickness)

    If (Me.thicknessOfLayer(layerIndex) - countOfLayers * maxThickness) > tolerance Then _
         countOfLayers = countOfLayers + 1

    Dim i As Integer
    For i = 1 To countOfLayers-1
        Call Me.splitAtDepth(depthOfTop + i * maxThickness, tolerance)
    Next i

End Sub

' Разделяет все слои на слои указанной толщины
' Не оставляет слои тоньше tolerance
Public Sub splitAllLayersWithMaxThickness(maxThickness As Double, Optional tolerance As Double = 0.01)
    
    Dim i As Integer : i = LBound(pLayerDepth)
    
    Do While i <= UBound(pLayerDepth)

        If thicknessOfLayer(i) > (maxThickness + tolerance) then
            Call Me.splitLayerWithMaxThickness(i, maxThickness, tolerance)
        End If

        i = i + 1
    Loop

End Sub

' Вычисляет толщину указанного слоя
' Если слоя нет, возвращает 0

Public Function thicknessOfLayer(layerIndex As Integer) As Double
    thicknessOfLayer = 0#

    If layerIndex = LBound(pLayerDepth) Then
        thicknessOfLayer = pLayerDepth(layerIndex)
        Exit Function
    End If

    If (layerIndex > LBound(pLayerDepth) And layerIndex <= UBound(pLayerDepth)) Then
        thicknessOfLayer = pLayerDepth(layerIndex) - pLayerDepth(layerIndex - 1)
    End If

End Function