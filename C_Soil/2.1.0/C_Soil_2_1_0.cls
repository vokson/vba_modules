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

' ���������� ��� ������� ����� � ������������ ������� �� ���� 25100-2011

Private Const WATER_DENSITY = 1# '�/��3

Private Const CLASS_ROCK = "��������"
Private Const CLASS_DISPERSAL = "����������"
Private Const CLASS_FROZEN = "�������"

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

Private Const FROST_HEAVING_TYPE_NO = "������������"
Private Const FROST_HEAVING_TYPE_LOW = "���������������"
Private Const FROST_HEAVING_TYPE_MIDDLE = "����������������"
Private Const FROST_HEAVING_TYPE_HIGH = "����������������"
Private Const FROST_HEAVING_TYPE_VERY_HIGH = "��������� ����������"

Private pSoilName As String

' ������������ �������� ������� �.1 ���� 25100-2011

Private pClassOfSoil As String ' ����� ������

Private pTypeBySize As String ' ������������� ������ �� �������
Private pSubtypeBySize As String ' ������������� ������ �� �������
Private pTypeByDegreeOfSaturation As String ' ������������� �����, ����������������� �� ������������ �������������
Private pTypeByDensity As String ' ������������� ����� �� ���������
Private pTypeByLiquidityIndex As String ' ������������� ��������� ������� �� ���������� ���������
Private pTypeByFrostHeaving As String ' ������������� ������ �� ������� �������� ������������

Private pWaterDensity As Double '��������� ������ (P) - �/��3
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
private pDegreeOfSaturation As Double '����������� ������������� (Sr) - �.�.
Private pLiquidityIndex As Double ' ���������� ��������� (IL) - %
Private pPlasticityIndex As Double ' ����� ������������ (IP) - %
Private pVoidRatio As Double ' ����������� ���������� (e) - �.�.
Private pGranulationFactor As Double ' ���������� ������������ �� (6.33) �� 22.13330.2011 (D) - �.�.
Private pFrozenDepth As Double ' ����������� ������� ����������� - �
Private pFrostHeaveRate As Double ' ������� �������� ������������ - %
Private pShearHeavingForce As Double' ����������� ���� ��������� ������� - ���
Private topSizeLimit, bottomSizeLimit, percentsBySize

Private Sub Class_Initialize()
      topSizeLimit = Array(100000, 800, 400, 200, 100, 60, 10, 5, 2, 1, 0.5, 0.25, 0.1, 0.05, 0.01, 0.002)
      bottomSizeLimit = Array(800, 400, 200, 100, 60, 10, 5, 2, 1, 0.5, 0.25, 0.1, 0.05, 0.01, 0.002, 0)
      percentsBySize = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)

      pWaterDensity = 1#
End Sub

Public Function getPropertyByName(nameOfParameter As String) As Variant

    Select Case nameOfParameter

        Case "NAME": getPropertyByName = Me.name
        Case "CLASS_OF_SOIL": getPropertyByName = Me.ClassOfSoil
        Case "TYPE_BY_SIZE": getPropertyByName = Me.TypeBySize
        Case "SUBTYPE_BY_SIZE": getPropertyByName = Me.SubtypeBySize
        Case "TYPE_BY_LIQUIDITY_INDEX": getPropertyByName = Me.TypeByLiquidityIndex
        Case "TYPE_BY_DENSITY": getPropertyByName = Me.TypeByDensity
        Case "TYPE_BY_DEGREE_OF_SATURATION": getPropertyByName = Me.TypeByDegreeOfSaturation
        Case "TYPE_BY_FROST_HEAVING": getPropertyByName = Me.TypeByFrostHeaving
        Case "YOUNG_MODULUS": getPropertyByName = Me.YoungModulus
        Case "SOIL_DENSITY": getPropertyByName = Me.SoilDensity
        Case "SOIL_DENSITY_1": getPropertyByName = Me.SoilDensity_1
        Case "SOIL_DENSITY_2": getPropertyByName =  Me.SoilDensity_2
        Case "SPECIFIC_COHESION": getPropertyByName =  Me.SpecificCohesion
        Case "SPECIFIC_COHESION_1": getPropertyByName = Me.SpecificCohesion_1
        Case "SPECIFIC_COHESION_2": getPropertyByName =  Me.SpecificCohesion_2
        Case "INTERNAL_FRICTION_ANGLE": getPropertyByName =  Me.InternalFrictionAngle
        Case "INTERNAL_FRICTION_ANGLE_1": getPropertyByName =  Me.InternalFrictionAngle_1
        Case "INTERNAL_FRICTION_ANGLE_2": getPropertyByName =  Me.InternalFrictionAngle_2
        Case "VOID_RATIO": getPropertyByName =  Me.VoidRatio
        Case "DEGREE_OF_SATURATION": getPropertyByName = Me.DegreeOfSaturation
        Case "WATER_CONTENT": getPropertyByName = Me.WaterContent
        Case "PLASTIC_LIMIT": getPropertyByName =  Me.PlasticLimit
        Case "LIQUID_LIMIT": getPropertyByName =  Me.LiquidLimit
        Case "PLASTICITY_INDEX": getPropertyByName =  Me.PlasticityIndex
        Case "LIQUIDITY_INDEX": getPropertyByName =  Me.LiquidityIndex
        Case "FROST_HEAVE_RATE": getPropertyByName =  Me.FrostHeaveRate
        Case "SHEAR_HEAVING_FORCE": getPropertyByName =  Me.ShearHeavingForce
        Case "GRANULATION_FACTOR": getPropertyByName = Me.GranulationFactor
        Case "FROZEN_DEPTH": getPropertyByName =  Me.FrozenDepth
        Case "WATER_DENSITY": getPropertyByName =  Me.waterDensity

    End Select

End Function

Public Function DeepCopy() As C_Soil
    Dim other As New C_Soil

    other.name = Me.name
    other.ClassOfSoil = Me.ClassOfSoil
    other.TypeBySize = Me.TypeBySize
    other.SubtypeBySize = Me.SubtypeBySize
    other.TypeByLiquidityIndex = Me.TypeByLiquidityIndex
    other.TypeByDensity = Me.TypeByDensity
    other.TypeByDegreeOfSaturation = Me.TypeByDegreeOfSaturation
    other.TypeByFrostHeaving = Me.TypeByFrostHeaving
    other.YoungModulus = Me.YoungModulus
    other.SoilDensity = Me.SoilDensity
    other.SoilDensity_1 = Me.SoilDensity_1
    other.SoilDensity_2 = Me.SoilDensity_2
    other.SpecificCohesion = Me.SpecificCohesion
    other.SpecificCohesion_1 = Me.SpecificCohesion_1
    other.SpecificCohesion_2 = Me.SpecificCohesion_2
    other.InternalFrictionAngle = Me.InternalFrictionAngle
    other.InternalFrictionAngle_1 = Me.InternalFrictionAngle_1
    other.InternalFrictionAngle_2 = Me.InternalFrictionAngle_2
    other.VoidRatio = Me.VoidRatio
    other.DegreeOfSaturation = Me.DegreeOfSaturation
    other.WaterContent = Me.WaterContent
    other.PlasticLimit = Me.PlasticLimit
    other.LiquidLimit = Me.LiquidLimit
    other.PlasticityIndex = Me.PlasticityIndex
    other.LiquidityIndex = Me.LiquidityIndex
    other.FrostHeaveRate = Me.FrostHeaveRate
    other.ShearHeavingForce = Me.ShearHeavingForce
    other.GranulationFactor = Me.GranulationFactor
    other.FrozenDepth = Me.FrozenDepth
    other.waterDensity = Me.waterDensity

    Set DeepCopy = other
End Function

Public Property Get waterDensity() As Double
    waterDensity = pWaterDensity
End Property

Public Property Let waterDensity(value As Double)
    pWaterDensity = value
End Property
   
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

Public Property Get VoidRatio() As Double
    VoidRatio = pVoidRatio
End Property

Public Property Let VoidRatio(e As Double)
    pVoidRatio = e
End Property

' ����������� ���������� (e) - �.�.
' ���� 25100-2011, ������� �.5
Public Sub AutoVoidRatio()
    pVoidRatio = (SolidParticlesDensity - DrySoilDensity) / DrySoilDensity
End Sub

Public Property Get WaterContent() As Double
    WaterContent = pWaterContent
End Property

Public Property Let WaterContent(W As Double)
    pWaterContent = W
End Property

Public Property Get DegreeOfSaturation() As Double
    DegreeOfSaturation = pDegreeOfSaturation
End Property

Public Property Let DegreeOfSaturation(value As Double)
    pDegreeOfSaturation = value
End Property

' ����������� ������������� (Sr) - �.�.
' ���� 25100-2011, ������� �.1
Public Sub AutoDegreeOfSaturation()
    pDegreeOfSaturation = 0.01 * WaterContent * SolidParticlesDensity / VoidRatio / WATER_DENSITY
End Sub

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

Public Property Get PlasticityIndex() As Double
    PlasticityIndex = pPlasticityIndex
End Property

Public Property Let PlasticityIndex(IP As Double)
    pPlasticityIndex = IP
End Property

' ����������� ������������� (IP) - %
' ���� 25100-2011, ������� �.17
Public Sub AutoPlasticityIndex()
    pPlasticityIndex = LiquidLimit - PlasticLimit
End Sub

Public Property Get LiquidityIndex() As Double
    LiquidityIndex = pLiquidityIndex
End Property

Public Property Let LiquidityIndex(IL As Double)
    pLiquidityIndex = IL
End Property

' ����������� ������������� (IL) - %
' ���� 25100-2011, ������� �.9
Public Sub AutoLiquidityIndex()
    If PlasticityIndex <> 0 Then pLiquidityIndex = (WaterContent - PlasticLimit) / PlasticityIndex
End Sub

Public Property Get GranulationFactor() As Double
    GranulationFactor = pGranulationFactor
End Property

Public Property Let GranulationFactor(D As Double)
    pGranulationFactor = D
End Property

Public Property Get FrozenDepth() As Double
    FrozenDepth = pFrozenDepth
End Property

Public Property Let FrozenDepth(depth As Double)
    pFrozenDepth = depth
End Property

Public Property Get FrostHeaveRate() As Double
    FrostHeaveRate = pFrostHeaveRate
End Property

Public Property Let FrostHeaveRate(value As Double)
    pFrostHeaveRate = value
End Property

Public Property Get ShearHeavingForce() As Double
    ShearHeavingForce = pShearHeavingForce
End Property

Public Property Let ShearHeavingForce(value As Double)
    pShearHeavingForce = value
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

' ����� ������
Public Property Get ClassOfSoil() As String
    ClassOfSoil = pClassOfSoil
End Property

' ����� ������
Public Property Let ClassOfSoil(classname As String)
    Select Case classname
        Case CLASS_DISPERSAL, CLASS_FROZEN:
             pClassOfSoil = classname
        Case Else
            Debug.Print classname & " is wrong class for soil"
    End Select
End Property

' ��� (�������������) ������ �� ������������������� �������
Public Property Get TypeBySize() As String
    TypeBySize = pTypeBySize
End Property

' ��� (�������������) ������ �� ������������������� �������
Public Property Let TypeBySize(name As String)
    Select Case ClassOfSoil

        Case CLASS_DISPERSAL, CLASS_FROZEN:
            Select Case name
                Case SIZE_TYPE_MACROFRAGMENTAL, SIZE_TYPE_SAND, SIZE_TYPE_CLAY_SANDY, _
                    SIZE_TYPE_CLAY_LOAM, SIZE_TYPE_CLAY:
                    pTypeBySize = name
                Case Else
                    Debug.Print name & " is wrong type by size for dispersal or frozen soils"
            End Select

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
            Debug.Print name & " ������������ ��� �� ���������� ������������� ��� ������"
    End Select
End Property

' ������������� ������ �� ������� �������� ������������
Public Property Get TypeByFrostHeaving() As String
    TypeByFrostHeaving = pTypeByFrostHeaving
End Property

' ������������� ������ �� ������� �������� ������������
Public Property Let TypeByFrostHeaving(name As String)
    Select Case name
        Case FROST_HEAVING_TYPE_NO, FROST_HEAVING_TYPE_LOW, _
             FROST_HEAVING_TYPE_MIDDLE, FROST_HEAVING_TYPE_HIGH, _
             FROST_HEAVING_TYPE_VERY_HIGH:
             pTypeByFrostHeaving = name
        Case Else
            Debug.Print name & " ������������ ��� �� ������� �������� ������������"
    End Select
End Property

' �������������� ����������� ���� �� ������� �������� ������������
' ������� �.27 ���� 25100-2011
Private Sub AutoTypeByFrostHeaving()

    pTypeByFrostHeaving = ""
        
    If FrostHeaveRate < 1# Then
        pTypeByFrostHeaving = FROST_HEAVING_TYPE_NO
        
    ElseIf FrostHeaveRate <= 3.5 Then
        pTypeByFrostHeaving = FROST_HEAVING_TYPE_LOW
    
    ElseIf FrostHeaveRate <= 7# Then
        pTypeByFrostHeaving = FROST_HEAVING_TYPE_MIDDLE

    ElseIf FrostHeaveRate <= 10# Then
        pTypeByFrostHeaving = FROST_HEAVING_TYPE_HIGH
    
    Else
        pTypeByFrostHeaving = FROST_HEAVING_TYPE_VERY_HIGH
    End If
    
End Sub

' �������������� ����������� ���� ��������������
Public Sub AssignTypesAutomatically()
   
    Call AutoPlasticityIndex
    Call AutoLiquidityIndex
    Call AutoVoidRatio
    Call AutoDegreeOfSaturation
    Call AutoTypeBySize
    Call AutoTypeByDegreeOfSaturation
    Call AutoTypeByDensity
    Call AutoTypeByLiquidityIndex
    Call AutoTypeByFrostHeaving
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





