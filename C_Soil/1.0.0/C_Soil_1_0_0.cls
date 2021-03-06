VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Soil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' ���������� ����� � ������������ ������� �� ���� 25100-2011

Private Const WATER_DENSITY = 1# '�/��3

Private Const SIZE_TYPE_MACROFRAGMENTAL = "����������������"
Private Const SIZE_TYPE_SAND = "�����"
Private Const SIZE_TYPE_CLAY_SANDY = "������"
Private Const SIZE_TYPE_CLAY_LOAM = "��������"
Private Const SIZE_TYPE_CLAY = "�����"

Private Const SIZE_MACROFRAGMENTAL_SUBTYPE_BOULDER = "��������"
Private Const SIZE_MACROFRAGMENTAL_SUBTYPE_COBBLES = "��������"
Private Const SIZE_MACROFRAGMENTAL_SUBTYPE_GRAVEL = "���������"

Private Const SIZE_SAND_SUBTYPE_GRAVEL = "�����������"
Private Const SIZE_SAND_SUBTYPE_COARSE = "�������"
Private Const SIZE_SAND_SUBTYPE_MIDDLE = "������� ���������"
Private Const SIZE_SAND_SUBTYPE_SMALL = "������"
Private Const SIZE_SAND_SUBTYPE_FINE = "���������"

Private Const SIZE_CLAY_SANDY_SUBTYPE_SANDY = "�����������"
Private Const SIZE_CLAY_SANDY_SUBTYPE_SILTY = "���������"

Private Const SIZE_CLAY_LOAM_SUBTYPE_LIGHT_SANDY = "������ �����������"
Private Const SIZE_CLAY_LOAM_SUBTYPE_LIGHT_SILTY = "������ ���������"
Private Const SIZE_CLAY_LOAM_SUBTYPE_HEAVY_SANDY = "������� �����������"
Private Const SIZE_CLAY_LOAM_SUBTYPE_HEAVY_SILTY = "������� ���������"

Private Const SIZE_CLAY_SUBTYPE_LIGHT_SANDY = "������ �����������"
Private Const SIZE_CLAY_SUBTYPE_LIGHT_SILTY = "������ ���������"
Private Const SIZE_CLAY_SUBTYPE_HEAVY = "�������"

Private Const SATURATION_TYPE_1 = "����� ������� �������������"
Private Const SATURATION_TYPE_2 = "������� ������� �������������"
Private Const SATURATION_TYPE_3 = "��������������"

Private Const DENSITY_TYPE_SAND_HIGH = "�������"
Private Const DENSITY_TYPE_SAND_MIDDLE = "������� ���������"
Private Const DENSITY_TYPE_SAND_LOW = "������"

Private Const LIQUIDITY_INDEX_TYPE_0 = "�������"
Private Const LIQUIDITY_INDEX_TYPE_1 = "�����������"
Private Const LIQUIDITY_INDEX_TYPE_2 = "��������������"
Private Const LIQUIDITY_INDEX_TYPE_3 = "����������"
Private Const LIQUIDITY_INDEX_TYPE_4 = "���������������"
Private Const LIQUIDITY_INDEX_TYPE_5 = "����������������"
Private Const LIQUIDITY_INDEX_TYPE_6 = "�������"

Private pSoilName As String

' ������������ �������� ������� �.1 ���� 25100-2011

Private pTypeBySize As String ' ������������� ������ �� �������
Private pSubtypeBySize As String ' ������������� ������ �� �������
Private pTypeByDegreeOfSaturation As String ' ������������� �����, ����������������� �� ������������ �������������
Private pTypeByDensity As String ' ������������� ����� �� ���������
Private pTypeByLiquidityIndex As String ' ������������� ��������� ������� �� ���������� ���������

Private pSoilDensity As Double '��������� ������ (P) - �/��3
Private pSoilDensity_1 As Double '�� �� ��� ������� �� I ������ ���������� ���������
Private pSoilDensity_2 As Double '�� �� ��� ������� �� II ������ ���������� ���������

Private pInternalFrictionAngle As Double '���� ����������� ������ (��) - �������
Private pInternalFrictionAngle_1 As Double '�� �� ��� ������� �� I ������ ���������� ���������
Private pInternalFrictionAngle_2 As Double '�� �� ��� ������� �� II ������ ���������� ���������

Private pSpecificCohesion As Double '�������� ��������� (�) - ���
Private pSpecificCohesion_1 As Double '�� �� ��� ������� �� I ������ ���������� ���������
Private pSpecificCohesion_2 As Double '�� �� ��� ������� �� II ������ ���������� ���������

Private pYoungModulus As Double ' ������ ��������� ������ (E) - ���
Private pSolidParticlesDensity As Double '��������� ������ ������ (Ps) - �/��3
Private pWaterContent As Double '��������� ��������� ������ (W) - %
Private pPlasticLimit As Double '������� ������������ (WP) - %
Private pLiquidLimit As Double '������� ��������� (WL) - %
Private topSizeLimit, bottomSizeLimit, percentsBySize

Private Sub Class_Initialize()
      topSizeLimit = Array(100000, 800, 400, 200, 100, 60, 10, 5, 2, 1, 0.5, 0.25, 0.1, 0.05, 0.01, 0.002)
      bottomSizeLimit = Array(800, 400, 200, 100, 60, 10, 5, 2, 1, 0.5, 0.25, 0.1, 0.05, 0.01, 0.002, 0)
      percentsBySize = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
End Sub
   
Public Property Get name() As String
    name = pSoilName
End Property

Public Property Let name(nameOfSoil As String)
    pSoilName = nameOfSoil
End Property

Public Property Get SoilDensity() As Double
    SoilDensity = pSoilDensity
End Property

Public Property Let SoilDensity(density As Double)
    pSoilDensity = density
End Property

Public Property Get SoilDensity_1() As Double
    SoilDensity_1 = pSoilDensity_1
End Property

Public Property Let SoilDensity_1(density As Double)
    pSoilDensity_1 = density
End Property

Public Property Get SoilDensity_2() As Double
    SoilDensity_2 = pSoilDensity_2
End Property

Public Property Let SoilDensity_2(density As Double)
    pSoilDensity_2 = density
End Property

Public Property Get InternalFrictionAngle() As Double
    InternalFrictionAngle = pInternalFrictionAngle
End Property

Public Property Let InternalFrictionAngle(angle As Double)
    pInternalFrictionAngle = angle
End Property

Public Property Get InternalFrictionAngle_1() As Double
    InternalFrictionAngle_1 = pInternalFrictionAngle_1
End Property

Public Property Let InternalFrictionAngle_1(angle As Double)
    pInternalFrictionAngle_1 = angle
End Property

Public Property Get InternalFrictionAngle_2() As Double
    InternalFrictionAngle_2 = pInternalFrictionAngle_2
End Property

Public Property Let InternalFrictionAngle_2(angle As Double)
    pInternalFrictionAngle_2 = angle
End Property

Public Property Get SpecificCohesion() As Double
    SpecificCohesion = pSpecificCohesion
End Property

Public Property Let SpecificCohesion(c As Double)
    pSpecificCohesion = c
End Property

Public Property Get SpecificCohesion_1() As Double
    SpecificCohesion_1 = pSpecificCohesion_1
End Property

Public Property Let SpecificCohesion_1(c As Double)
    pSpecificCohesion_1 = c
End Property

Public Property Get SpecificCohesion_2() As Double
    SpecificCohesion_2 = pSpecificCohesion_2
End Property

Public Property Let SpecificCohesion_2(c As Double)
    pSpecificCohesion_2 = c
End Property

Public Property Get SolidParticlesDensity() As Double
    SolidParticlesDensity = pSolidParticlesDensity
End Property

Public Property Let SolidParticlesDensity(density As Double)
    pSolidParticlesDensity = density
End Property

Public Property Get YoungModulus() As Double
    YoungModulus = pYoungModulus
End Property

Public Property Let YoungModulus(E As Double)
    pYoungModulus = E
End Property

' ��������� ������ ������ (Pd) - �/��3
' ���� 25100-2011, ������� �.8
Public Property Get DrySoilDensity() As Double
    DrySoilDensity = SoilDensity / (1 + 0.01 * WaterContent)
End Property

' ����������� ���������� (e) - �.�.
' ���� 25100-2011, ������� �.5
Public Property Get VoidRatio() As Double
    VoidRatio = (SolidParticlesDensity - DrySoilDensity) / DrySoilDensity
End Property

Public Property Get WaterContent() As Double
    WaterContent = pWaterContent
End Property

Public Property Let WaterContent(W As Double)
    pWaterContent = W
End Property

' ����������� ������������� (Sr) - �.�.
' ���� 25100-2011, ������� �.1
Public Property Get DegreeOfSaturation() As Double
    DegreeOfSaturation = 0.01 * WaterContent * SolidParticlesDensity / VoidRatio / WATER_DENSITY
End Property

Public Property Get PlasticLimit() As Double
    PlasticLimit = pPlasticLimit
End Property

Public Property Let PlasticLimit(WP As Double)
    pPlasticLimit = WP
End Property

Public Property Get LiquidLimit() As Double
    LiquidLimit = pLiquidLimit
End Property

Public Property Let LiquidLimit(WL As Double)
    pLiquidLimit = WL
End Property

' ����������� ������������� (IP) - %
' ���� 25100-2011, ������� �.17
Public Property Get PlasticityIndex() As Double
    PlasticityIndex = LiquidLimit - PlasticLimit
End Property

' ����������� ������������� (IL) - %
' ���� 25100-2011, ������� �.9
Public Property Get LiquidityIndex() As Double
    If PlasticityIndex <> 0 Then LiquidityIndex = (WaterContent - PlasticLimit) / PlasticityIndex
End Property

' ���� ������ � ������� �� ���. ������� ������ � ������ �������
Private Function findIndexByMinSize(size As Double) As Integer
    findIndexByMinSize = findIndexBySize(bottomSizeLimit, size)
End Function

' ���� ������ � ������� �� ����. ������� ������ � ������ �������
Private Function findIndexByMaxSize(size As Double) As Integer
    findIndexByMaxSize = findIndexBySize(topSizeLimit, size)
End Function

' ���� ������ � ������� �� ������� ������ � ������ �������
Private Function findIndexBySize(arr As Variant, size As Double) As Integer
    findIndexBySize = -1
    
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        If (arr(i) = size) Then
            findIndexBySize = i
            Exit Function
        End If
    Next i
    
End Function

' ������������� ���������� ���������� ������ �� ������ ������� ��������, ��
' ���� ����� ������� ���� - true, ���� ��� - false
Public Function setSize(minSize As Double, howManyPercents As Double) As Boolean
    Dim pos As Integer
    pos = findIndexByMinSize(minSize)
    
    If (pos <> -1) Then
        percentsBySize(pos) = howManyPercents
        setSize = True
    Else
        setSize = False
    End If
        
End Function

' ������������ ���������� ���������� ������ ����� size1 � size2, ��
Public Function getSumPercentOfSizeBetween(size1 As Double, size2 As Double) As Double
    Dim i As Integer
    Dim i1 As Integer
    Dim i2 As Integer
    Dim min As Double
    Dim max As Double
    
    If (size1 < size2) Then
        min = size1: max = size2
    Else
        max = size1: min = size2
    End If
    
    getSumPercentOfSizeBetween = 0
    i1 = findIndexByMaxSize(max)
    i2 = findIndexByMinSize(min)
        
    For i = i1 To i2
        getSumPercentOfSizeBetween = getSumPercentOfSizeBetween + percentsBySize(i)
    Next i
        
End Function

' ������������ ���������� ���������� ������ ������, ��� size, ��
Public Function getSumPercentOfSizeMoreThan(size As Double) As Double
  getSumPercentOfSizeMoreThan = getSumPercentOfSizeBetween(size, CDbl(topSizeLimit(LBound(topSizeLimit))))
End Function

' ������������ ���������� ���������� ������ ������, ��� size, ��
Public Function getSumPercentOfSizeLessThan(size As Double) As Double
  getSumPercentOfSizeLessThan = getSumPercentOfSizeBetween(size, CDbl(bottomSizeLimit(UBound(bottomSizeLimit))))
End Function

' ���������� �������� �� ����� �������
Private Function isCohesiveSoil() As Boolean
    isCohesiveSoil = False
    If getSumPercentOfSizeLessThan(0.002) >= 3 And PlasticityIndex >= 1 Then isCohesiveSoil = True
End Function

' ���������� �������� �� ����� ��������� ��������
Private Function isSandySoil() As Boolean
    isSandySoil = False
    If getSumPercentOfSizeBetween(0.05, 2) > 50 And PlasticityIndex < 1 Then isSandySoil = True
End Function

' ���������� �������� �� ����� ��������� ����������������
Private Function isMacrofragmentalSoil() As Boolean
    isMacrofragmentalSoil = False
    If getSumPercentOfSizeMoreThan(2) > 50 Then isMacrofragmentalSoil = True
End Function

' ��� (�������������) ������ �� ������������������� �������
Public Property Get TypeBySize() As String
    TypeBySize = pTypeBySize
End Property

' ��� (�������������) ������ �� ������������������� �������
Public Property Let TypeBySize(name As String)
    Select Case name
        Case SIZE_TYPE_MACROFRAGMENTAL, SIZE_TYPE_SAND, SIZE_TYPE_CLAY_SANDY, _
            SIZE_TYPE_CLAY_LOAM, SIZE_TYPE_CLAY:
             pTypeBySize = name
        Case Else
            Debug.Print name & " is wrong type by size for soil"
    End Select
End Property

' ������ (�������������) ������ �� ������������������� �������
Public Property Get SubtypeBySize() As String
    SubtypeBySize = pSubtypeBySize
End Property

' ������ (�������������) ������ �� ������������������� �������
Public Property Let SubtypeBySize(name As String)
    Select Case pTypeBySize
        Case SIZE_TYPE_MACROFRAGMENTAL:
            Select Case name
                Case SIZE_MACROFRAGMENTAL_SUBTYPE_BOULDER, _
                     SIZE_MACROFRAGMENTAL_SUBTYPE_COBBLES, _
                     SIZE_MACROFRAGMENTAL_SUBTYPE_GRAVEL:
                     pSubtypeBySize = name
                Case Else: GoTo Handler
            End Select
            
        Case SIZE_TYPE_SAND:
            Select Case name
                Case SIZE_SAND_SUBTYPE_GRAVEL, _
                     SIZE_SAND_SUBTYPE_COARSE, _
                     SIZE_SAND_SUBTYPE_MIDDLE, _
                     SIZE_SAND_SUBTYPE_SMALL, _
                     SIZE_SAND_SUBTYPE_FINE:
                     pSubtypeBySize = name
                Case Else: GoTo Handler
            End Select
            
        Case SIZE_TYPE_CLAY_SANDY:
            Select Case name
                Case SIZE_CLAY_SANDY_SUBTYPE_SANDY, _
                     SIZE_CLAY_SANDY_SUBTYPE_SILTY:
                     pSubtypeBySize = name
                Case Else: GoTo Handler
            End Select
            
        Case SIZE_TYPE_CLAY_LOAM:
            Select Case name
                Case SIZE_CLAY_LOAM_SUBTYPE_LIGHT_SANDY, _
                     SIZE_CLAY_LOAM_SUBTYPE_LIGHT_SILTY, _
                     SIZE_CLAY_LOAM_SUBTYPE_HEAVY_SANDY, _
                     SIZE_CLAY_LOAM_SUBTYPE_HEAVY_SILTY:
                     pSubtypeBySize = name
                Case Else: GoTo Handler
            End Select
            
        Case SIZE_TYPE_CLAY:
            Select Case name
                Case SIZE_CLAY_SUBTYPE_LIGHT_SANDY, _
                     SIZE_CLAY_SUBTYPE_LIGHT_SILTY, _
                     SIZE_CLAY_SUBTYPE_HEAVY:
                     pSubtypeBySize = name
                Case Else: GoTo Handler
            End Select
    End Select
    
    Exit Property
    
Handler:
    Debug.Print name & " is wrong subtype by size for the soil"
    
End Property

' �������������� ����������� (�������������) ������ �� ������������������� �������
' ������� �.9 � �.17 ���� 25100-2011
Private Sub AutoTypeBySize()
    pTypeBySize = "": pSubtypeBySize = ""
    
    If isMacrofragmentalSoil() = True Then
        pTypeBySize = SIZE_TYPE_MACROFRAGMENTAL
        
        If getSumPercentOfSizeMoreThan(200) > 50 Then
            pSubtypeBySize = SIZE_MACROFRAGMENTAL_SUBTYPE_BOULDER
            
        ElseIf getSumPercentOfSizeMoreThan(10) > 50 Then
            pSubtypeBySize = SIZE_MACROFRAGMENTAL_SUBTYPE_COBBLES
        
        ElseIf getSumPercentOfSizeMoreThan(2) > 50 Then
            pSubtypeBySize = SIZE_MACROFRAGMENTAL_SUBTYPE_GRAVEL
        End If
    End If
    
    If isSandySoil() = True Then
        pTypeBySize = SIZE_TYPE_SAND
        
        If getSumPercentOfSizeMoreThan(2) > 25 Then
            pSubtypeBySize = SIZE_SAND_SUBTYPE_GRAVEL
            
        ElseIf getSumPercentOfSizeMoreThan(0.5) > 50 Then
            pSubtypeBySize = SIZE_SAND_SUBTYPE_COARSE
        
        ElseIf getSumPercentOfSizeMoreThan(0.25) > 50 Then
            pSubtypeBySize = SIZE_SAND_SUBTYPE_MIDDLE
            
        ElseIf getSumPercentOfSizeMoreThan(0.1) >= 75 Then
            pSubtypeBySize = SIZE_SAND_SUBTYPE_SMALL
            
        ElseIf getSumPercentOfSizeMoreThan(0.1) < 75 Then
            pSubtypeBySize = SIZE_SAND_SUBTYPE_FINE
        End If
    End If
    
    If isCohesiveSoil() = True Then
        
        
        If PlasticityIndex >= 1 And PlasticityIndex <= 7 And getSumPercentOfSizeBetween(0.05, 2) >= 50 Then
            pTypeBySize = SIZE_TYPE_CLAY_SANDY
            pSubtypeBySize = SIZE_CLAY_SANDY_SUBTYPE_SANDY
            
        ElseIf PlasticityIndex >= 1 And PlasticityIndex <= 7 And getSumPercentOfSizeBetween(0.05, 2) < 50 Then
            pTypeBySize = SIZE_TYPE_CLAY_SANDY
            pSubtypeBySize = SIZE_CLAY_SANDY_SUBTYPE_SILTY
            
        ElseIf PlasticityIndex > 7 And PlasticityIndex <= 12 And getSumPercentOfSizeBetween(0.05, 2) >= 40 Then
            pTypeBySize = SIZE_TYPE_CLAY_LOAM
            pSubtypeBySize = SIZE_CLAY_LOAM_SUBTYPE_LIGHT_SANDY
        
        ElseIf PlasticityIndex > 7 And PlasticityIndex <= 12 And getSumPercentOfSizeBetween(0.05, 2) < 40 Then
            pTypeBySize = SIZE_TYPE_CLAY_LOAM
            pSubtypeBySize = SIZE_CLAY_LOAM_SUBTYPE_LIGHT_SILTY
            
        ElseIf PlasticityIndex > 12 And PlasticityIndex <= 17 And getSumPercentOfSizeBetween(0.05, 2) >= 40 Then
            pTypeBySize = SIZE_TYPE_CLAY_LOAM
            pSubtypeBySize = SIZE_CLAY_LOAM_SUBTYPE_HEAVY_SANDY
            
        ElseIf PlasticityIndex > 12 And PlasticityIndex <= 17 And getSumPercentOfSizeBetween(0.05, 2) < 40 Then
            pTypeBySize = SIZE_TYPE_CLAY_LOAM
            pSubtypeBySize = SIZE_CLAY_LOAM_SUBTYPE_HEAVY_SILTY
            
        ElseIf PlasticityIndex > 17 And PlasticityIndex <= 27 And getSumPercentOfSizeBetween(0.05, 2) >= 40 Then
            pTypeBySize = SIZE_TYPE_CLAY
            pSubtypeBySize = SIZE_CLAY_SUBTYPE_LIGHT_SANDY
            
        ElseIf PlasticityIndex > 17 And PlasticityIndex <= 27 And getSumPercentOfSizeBetween(0.05, 2) < 40 Then
            pTypeBySize = SIZE_TYPE_CLAY
            pSubtypeBySize = SIZE_CLAY_SUBTYPE_LIGHT_SILTY
            
        ElseIf PlasticityIndex > 27 Then
            pTypeBySize = SIZE_TYPE_CLAY
            pSubtypeBySize = SIZE_CLAY_SUBTYPE_HEAVY
        End If
        
    End If
    
End Sub

' ������������� ������ � ���������������� �� ������������ �������������
Public Property Get TypeByDegreeOfSaturation() As String
    TypeByDegreeOfSaturation = pTypeByDegreeOfSaturation
End Property

' ������������� ������ � ���������������� �� ������������ �������������
Public Property Let TypeByDegreeOfSaturation(name As String)
    Select Case name
        Case SATURATION_TYPE_1, SATURATION_TYPE_2, SATURATION_TYPE_3:
             pTypeByDegreeOfSaturation = name
        Case Else
            Debug.Print name & " is wrong type by degree of saturation for soil"
    End Select
End Property

' �������������� ����������� ���� ��������������
Public Sub AssignTypesAutomatically()
    Call AutoTypeBySize
    Call AutoTypeByDegreeOfSaturation
    Call AutoTypeByDensity
    Call AutoTypeByLiquidityIndex
End Sub

' �������������� ����������� ������ � ���������������� �� ������������ �������������
' ������� �.9 � �.17 ���� 25100-2011
Private Sub AutoTypeByDegreeOfSaturation()
    pTypeByDegreeOfSaturation = ""
    
    If TypeBySize = SIZE_TYPE_MACROFRAGMENTAL Or TypeBySize = SIZE_TYPE_SAND Then
        
        If DegreeOfSaturation > 0 And DegreeOfSaturation <= 0.5 Then
            pTypeByDegreeOfSaturation = SATURATION_TYPE_1
            
        ElseIf DegreeOfSaturation > 0.5 And DegreeOfSaturation <= 0.8 Then
            pTypeByDegreeOfSaturation = SATURATION_TYPE_2
        
        ElseIf DegreeOfSaturation > 0.8 And DegreeOfSaturation <= 1# Then
            pTypeByDegreeOfSaturation = SATURATION_TYPE_3
        End If
    End If
    
End Sub

' ������������� ������ �� ���������
Public Property Get TypeByDensity() As String
    TypeByDensity = pTypeByDensity
End Property

' ������������� ������ �� ���������
Public Property Let TypeByDensity(name As String)
    Select Case name
        Case DENSITY_TYPE_SAND_HIGH, DENSITY_TYPE_SAND_MIDDLE, DENSITY_TYPE_SAND_LOW:
             pTypeByDensity = name
        Case Else
            Debug.Print name & " is wrong type by density for sand"
    End Select
End Property

' �������������� ����������� ������ ���������
' ������� �.12 ���� 25100-2011
Private Sub AutoTypeByDensity()
    pTypeByDensity = ""
    
    If TypeBySize = SIZE_TYPE_SAND Then
        
        If SubtypeBySize = SIZE_SAND_SUBTYPE_GRAVEL Or SubtypeBySize = SIZE_SAND_SUBTYPE_COARSE _
                Or SubtypeBySize = SIZE_SAND_SUBTYPE_MIDDLE Then
                
            If VoidRatio <= 0.55 Then pTypeByDensity = DENSITY_TYPE_SAND_HIGH
            If VoidRatio > 0.55 And VoidRatio <= 0.7 Then pTypeByDensity = DENSITY_TYPE_SAND_MIDDLE
            If VoidRatio > 0.7 Then pTypeByDensity = DENSITY_TYPE_SAND_LOW
                
        End If
        
        If SubtypeBySize = SIZE_SAND_SUBTYPE_SMALL Then
            If VoidRatio <= 0.6 Then pTypeByDensity = DENSITY_TYPE_SAND_HIGH
            If VoidRatio > 0.6 And VoidRatio <= 0.75 Then pTypeByDensity = DENSITY_TYPE_SAND_MIDDLE
            If VoidRatio > 0.75 Then pTypeByDensity = DENSITY_TYPE_SAND_LOW
        End If
        
        If SubtypeBySize = SIZE_SAND_SUBTYPE_FINE Then
            If VoidRatio <= 0.6 Then pTypeByDensity = DENSITY_TYPE_SAND_HIGH
            If VoidRatio > 0.6 And VoidRatio <= 0.8 Then pTypeByDensity = DENSITY_TYPE_SAND_MIDDLE
            If VoidRatio > 0.8 Then pTypeByDensity = DENSITY_TYPE_SAND_LOW
        End If
        
    End If
    
End Sub

' ������������� ��������� ������� �� ���������� ���������
Public Property Get TypeByLiquidityIndex() As String
    TypeByLiquidityIndex = pTypeByLiquidityIndex
End Property

' ������������� ��������� ������� �� ���������� ���������
Public Property Let TypeByLiquidityIndex(name As String)

    If TypeBySize = SIZE_TYPE_CLAY_SANDY Then
        Select Case name
            Case LIQUIDITY_INDEX_TYPE_0, LIQUIDITY_INDEX_TYPE_3, LIQUIDITY_INDEX_TYPE_6:
                 pTypeByLiquidityIndex = name
            Case Else
                GoTo Handler
        End Select
    End If
    
    If TypeBySize = SIZE_TYPE_CLAY_LOAM Or TypeBySize = SIZE_TYPE_CLAY Then
        Select Case name
            Case LIQUIDITY_INDEX_TYPE_0, LIQUIDITY_INDEX_TYPE_1, LIQUIDITY_INDEX_TYPE_2, _
                 LIQUIDITY_INDEX_TYPE_4, LIQUIDITY_INDEX_TYPE_5, LIQUIDITY_INDEX_TYPE_6:
                 pTypeByLiquidityIndex = name
            Case Else
                GoTo Handler
        End Select
    End If
    
    Exit Property
    
Handler:
    Debug.Print name & " is wrong type by liquidity index"
    
End Property

' �������������� ����������� ��������� ������� �� ���������� ���������
' ������� �.19 ���� 25100-2011
Private Sub AutoTypeByLiquidityIndex()
    pTypeByLiquidityIndex = ""
    
    If TypeBySize = SIZE_TYPE_CLAY_SANDY Then
        
        If LiquidityIndex < 0 Then pTypeByLiquidityIndex = LIQUIDITY_INDEX_TYPE_0
        If LiquidityIndex >= 0 And LiquidityIndex <= 1# Then pTypeByLiquidityIndex = LIQUIDITY_INDEX_TYPE_3
        If LiquidityIndex > 1# Then pTypeByLiquidityIndex = LIQUIDITY_INDEX_TYPE_6
                
    End If
    
    If TypeBySize = SIZE_TYPE_CLAY_LOAM Or TypeBySize = SIZE_TYPE_CLAY Then
        
        If LiquidityIndex < 0 Then pTypeByLiquidityIndex = LIQUIDITY_INDEX_TYPE_0
        If LiquidityIndex >= 0 And LiquidityIndex <= 0.25 Then pTypeByLiquidityIndex = LIQUIDITY_INDEX_TYPE_1
        If LiquidityIndex > 0.25 And LiquidityIndex <= 0.5 Then pTypeByLiquidityIndex = LIQUIDITY_INDEX_TYPE_2
        If LiquidityIndex > 0.5 And LiquidityIndex <= 0.75 Then pTypeByLiquidityIndex = LIQUIDITY_INDEX_TYPE_4
        If LiquidityIndex > 0.75 And LiquidityIndex <= 1# Then pTypeByLiquidityIndex = LIQUIDITY_INDEX_TYPE_5
        If LiquidityIndex > 1# Then pTypeByLiquidityIndex = LIQUIDITY_INDEX_TYPE_6
                
    End If
    
End Sub

'
'Public salinity As String
'Public frost_heaving As String
'Public seismicity_category As String
'
'Public Dsal As Double
'Public efh As Double
'Public Ro As Double
'
'Public isBulk As Boolean





