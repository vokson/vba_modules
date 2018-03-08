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

' Дисперсный или мерзлый грунт с определением свойств по ГОСТ 25100-2011

Private Const WATER_DENSITY = 1# 'г/см3

Private Const CLASS_ROCK = "СКАЛЬНЫЙ"
Private Const CLASS_DISPERSAL = "ДИСПЕРСНЫЙ"
Private Const CLASS_FROZEN = "МЕРЗЛЫЙ"

Private Const SIZE_TYPE_MACROFRAGMENTAL = "КРУПНООБЛОМОЧНЫЙ"
Private Const SIZE_TYPE_SAND = "ПЕСОК"
Private Const SIZE_TYPE_CLAY_SANDY = "СУПЕСЬ"
Private Const SIZE_TYPE_CLAY_LOAM = "СУГЛИНОК"
Private Const SIZE_TYPE_CLAY = "ГЛИНА"

Private Const SIZE_MACROFRAGMENTAL_SUBTYPE_BOULDER = "ВАЛУННЫЙ"
Private Const SIZE_MACROFRAGMENTAL_SUBTYPE_COBBLES = "ГАЛЕЧНЫЙ"
Private Const SIZE_MACROFRAGMENTAL_SUBTYPE_GRAVEL = "ГРАВИЙНЫЙ"

Private Const SIZE_SAND_SUBTYPE_GRAVEL = "ГРАВЕЛИСТЫЙ"
Private Const SIZE_SAND_SUBTYPE_COARSE = "КРУПНЫЙ"
Private Const SIZE_SAND_SUBTYPE_MIDDLE = "СРЕДНЕЙ КРУПНОСТИ"
Private Const SIZE_SAND_SUBTYPE_SMALL = "МЕЛКИЙ"
Private Const SIZE_SAND_SUBTYPE_FINE = "ПЫЛЕВАТЫЙ"

Private Const SIZE_CLAY_SANDY_SUBTYPE_SANDY = "ПЕСЧАНИСТАЯ"
Private Const SIZE_CLAY_SANDY_SUBTYPE_SILTY = "ПЫЛЕВАТАЯ"

Private Const SIZE_CLAY_LOAM_SUBTYPE_LIGHT_SANDY = "ЛЕГКИЙ ПЕСЧАНИСТЫЙ"
Private Const SIZE_CLAY_LOAM_SUBTYPE_LIGHT_SILTY = "ЛЕГКИЙ ПЫЛЕВАТЫЙ"
Private Const SIZE_CLAY_LOAM_SUBTYPE_HEAVY_SANDY = "ТЯЖЕЛЫЙ ПЕСЧАНИСТЫЙ"
Private Const SIZE_CLAY_LOAM_SUBTYPE_HEAVY_SILTY = "ТЯЖЕЛЫЙ ПЫЛЕВАТЫЙ"

Private Const SIZE_CLAY_SUBTYPE_LIGHT_SANDY = "ЛЕГКАЯ ПЕСЧАНИСТАЯ"
Private Const SIZE_CLAY_SUBTYPE_LIGHT_SILTY = "ЛЕГКАЯ ПЫЛЕВАТАЯ"
Private Const SIZE_CLAY_SUBTYPE_HEAVY = "ТЯЖЕЛАЯ"

Private Const SATURATION_TYPE_1 = "МАЛОЙ СТЕПЕНИ ВОДОНАСЫЩЕНИЯ"
Private Const SATURATION_TYPE_2 = "СРЕДНЕЙ СТЕПЕНИ ВОДОНАСЫЩЕНИЯ"
Private Const SATURATION_TYPE_3 = "ВОДОНАСЫЩЕННЫЙ"

Private Const DENSITY_TYPE_SAND_HIGH = "ПЛОТНЫЙ"
Private Const DENSITY_TYPE_SAND_MIDDLE = "СРЕДНЕЙ ПЛОТНОСТИ"
Private Const DENSITY_TYPE_SAND_LOW = "РЫХЛЫЙ"

Private Const LIQUIDITY_INDEX_TYPE_0 = "ТВЕРДЫЙ"
Private Const LIQUIDITY_INDEX_TYPE_1 = "ПОЛУТВЕРДЫЙ"
Private Const LIQUIDITY_INDEX_TYPE_2 = "ТУГОПЛАСТИЧНЫЙ"
Private Const LIQUIDITY_INDEX_TYPE_3 = "ПЛАСТИЧНЫЙ"
Private Const LIQUIDITY_INDEX_TYPE_4 = "МЯГКОПЛАСТИЧНЫЙ"
Private Const LIQUIDITY_INDEX_TYPE_5 = "ТЕКУЧЕПЛАСТИЧНЫЙ"
Private Const LIQUIDITY_INDEX_TYPE_6 = "ТЕКУЧИЙ"

Private pSoilName As String

' Наименования согласно Таблице Ж.1 ГОСТ 25100-2011

Private pClassOfSoil As String ' Класс грунта

Private pTypeBySize As String ' Разновидность грунта по составу
Private pSubtypeBySize As String ' Разновидность грунта по составу
Private pTypeByDegreeOfSaturation As String ' Разновидность песка, крупнообломочного по коэффициенту водонасыщения
Private pTypeByDensity As String ' Разновидность песка по плотности
Private pTypeByLiquidityIndex As String ' Разновидность глинистых грунтов по показателю текучести

Private pSoilDensity As Double 'Плотность грунта (P) - г/см3
Private pSoilDensity_1 As Double 'То же для расчета по I группе предельных состояний
Private pSoilDensity_2 As Double 'То же для расчета по II группе предельных состояний

Private pInternalFrictionAngle As Double 'Угол внутреннего трения (Фи) - градусы
Private pInternalFrictionAngle_1 As Double 'То же для расчета по I группе предельных состояний
Private pInternalFrictionAngle_2 As Double 'То же для расчета по II группе предельных состояний

Private pSpecificCohesion As Double 'Удельное сцепление (С) - кПа
Private pSpecificCohesion_1 As Double 'То же для расчета по I группе предельных состояний
Private pSpecificCohesion_2 As Double 'То же для расчета по II группе предельных состояний

Private pYoungModulus As Double ' Модуль упругости грунта (E) - МПа
Private pSolidParticlesDensity As Double 'Плотность частиц грунта (Ps) - г/см3
Private pWaterContent As Double 'Природная влажность грунта (W) - %
Private pPlasticLimit As Double 'Граница раскатывания (WP) - %
Private pLiquidLimit As Double 'Граница текучести (WL) - %
private pDegreeOfSaturation As Double 'Коэффициент водонасыщения (Sr) - д.е.
Private pLiquidityIndex As Double ' Показатель текучести (IL) - %
Private pPlasticityIndex As Double ' Число пластичности (IP) - %
Private pVoidRatio As Double ' Коэффициент пористости (e) - д.е.
Private topSizeLimit, bottomSizeLimit, percentsBySize

Private Sub Class_Initialize()
      topSizeLimit = Array(100000, 800, 400, 200, 100, 60, 10, 5, 2, 1, 0.5, 0.25, 0.1, 0.05, 0.01, 0.002)
      bottomSizeLimit = Array(800, 400, 200, 100, 60, 10, 5, 2, 1, 0.5, 0.25, 0.1, 0.05, 0.01, 0.002, 0)
      percentsBySize = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
End Sub

Public Function DeepCopy() As C_Soil
    Dim other As New C_Soil

    other.name = Me.name
    other.ClassOfSoil = Me.ClassOfSoil
    other.TypeBySize = Me.TypeBySize
    other.SubtypeBySize = Me.SubtypeBySize
    other.TypeByLiquidityIndex = Me.TypeByLiquidityIndex
    other.TypeByDensity = Me.TypeByDensity
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

    Set DeepCopy = other
End Function
   
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

' Плотность сухого грунта (Pd) - г/см3
' ГОСТ 25100-2011, формула А.8
Public Property Get DrySoilDensity() As Double
    DrySoilDensity = SoilDensity / (1 + 0.01 * WaterContent)
End Property

Public Property Get VoidRatio() As Double
    VoidRatio = pVoidRatio
End Property

Public Property Let VoidRatio(e As Double)
    pVoidRatio = e
End Property

' Коэффициент пористости (e) - д.е.
' ГОСТ 25100-2011, формула А.5
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

' Коэффициент водонасыщения (Sr) - д.е.
' ГОСТ 25100-2011, формула А.1
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

' Коэффициент водонасыщения (IP) - %
' ГОСТ 25100-2011, формула А.17
Public Sub AutoPlasticityIndex()
    pPlasticityIndex = LiquidLimit - PlasticLimit
End Sub

Public Property Get LiquidityIndex() As Double
    LiquidityIndex = pLiquidityIndex
End Property

Public Property Let LiquidityIndex(IL As Double)
    pLiquidityIndex = IL
End Property

' Коэффициент водонасыщения (IL) - %
' ГОСТ 25100-2011, формула А.9
Public Sub AutoLiquidityIndex()
    If PlasticityIndex <> 0 Then pLiquidityIndex = (WaterContent - PlasticLimit) / PlasticityIndex
End Sub

' Ищет индекс в массиве по мин. размеру частиц в данной фракции
Private Function findIndexByMinSize(size As Double) As Integer
    findIndexByMinSize = findIndexBySize(bottomSizeLimit, size)
End Function

' Ищет индекс в массиве по макс. размеру частиц в данной фракции
Private Function findIndexByMaxSize(size As Double) As Integer
    findIndexByMaxSize = findIndexBySize(topSizeLimit, size)
End Function

' Ищет индекс в массиве по размеру частиц в данной фракции
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

' Устанавливает процентное содержание частиц по нижней границе размеров, мм
' Если такая граница есть - true, если нет - false
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

' Кумулятивное процентное содержание частиц между size1 и size2, мм
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

' Кумулятивное процентное содержание частиц больше, чем size, мм
Public Function getSumPercentOfSizeMoreThan(size As Double) As Double
  getSumPercentOfSizeMoreThan = getSumPercentOfSizeBetween(size, CDbl(topSizeLimit(LBound(topSizeLimit))))
End Function

' Кумулятивное процентное содержание частиц меньше, чем size, мм
Public Function getSumPercentOfSizeLessThan(size As Double) As Double
  getSumPercentOfSizeLessThan = getSumPercentOfSizeBetween(size, CDbl(bottomSizeLimit(UBound(bottomSizeLimit))))
End Function

' Определяет является ли грунт связным
Private Function isCohesiveSoil() As Boolean
    isCohesiveSoil = False
    If getSumPercentOfSizeLessThan(0.002) >= 3 And PlasticityIndex >= 1 Then isCohesiveSoil = True
End Function

' Определяет является ли грунт несвязным песчаным
Private Function isSandySoil() As Boolean
    isSandySoil = False
    If getSumPercentOfSizeBetween(0.05, 2) > 50 And PlasticityIndex < 1 Then isSandySoil = True
End Function

' Определяет является ли грунт несвязным крупнообломочным
Private Function isMacrofragmentalSoil() As Boolean
    isMacrofragmentalSoil = False
    If getSumPercentOfSizeMoreThan(2) > 50 Then isMacrofragmentalSoil = True
End Function

' Класс грунта
Public Property Get ClassOfSoil() As String
    ClassOfSoil = pClassOfSoil
End Property

' Класс грунта
Public Property Let ClassOfSoil(classname As String)
    Select Case classname
        Case CLASS_DISPERSAL, CLASS_FROZEN:
             pClassOfSoil = classname
        Case Else
            Debug.Print classname & " is wrong class for soil"
    End Select
End Property

' Вид (разновидность) грунта по гранулометрическому составу
Public Property Get TypeBySize() As String
    TypeBySize = pTypeBySize
End Property

' Вид (разновидность) грунта по гранулометрическому составу
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

' Подвид (разновидность) грунта по гранулометрическому составу
Public Property Get SubtypeBySize() As String
    SubtypeBySize = pSubtypeBySize
End Property

' Подвид (разновидность) грунта по гранулометрическому составу
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

' Автоматическое определение (разновидность) грунта по гранулометрическому составу
' Таблица Б.9 и Б.17 ГОСТ 25100-2011
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

' Разновидность песков и крупнообломочных по коэффициенту водонасыщения
Public Property Get TypeByDegreeOfSaturation() As String
    TypeByDegreeOfSaturation = pTypeByDegreeOfSaturation
End Property

' Разновидность песков и крупнообломочных по коэффициенту водонасыщения
Public Property Let TypeByDegreeOfSaturation(name As String)
    Select Case name
        Case SATURATION_TYPE_1, SATURATION_TYPE_2, SATURATION_TYPE_3:
             pTypeByDegreeOfSaturation = name
        Case Else
            Debug.Print name & " is wrong type by degree of saturation for soil"
    End Select
End Property

' Автоматическое определение всех разновидностей
Public Sub AssignTypesAutomatically()
    Call AutoDegreeOfSaturation
    Call AutoPlasticityIndex
    Call AutoLiquidityIndex
    Call AutoVoidRatio
    Call AutoTypeBySize
    Call AutoTypeByDegreeOfSaturation
    Call AutoTypeByDensity
    Call AutoTypeByLiquidityIndex
End Sub

' Автоматическое определение песков и крупнообломочных по коэффициенту водонасыщения
' Таблица Б.9 и Б.17 ГОСТ 25100-2011
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

' Разновидность песков по плотности
Public Property Get TypeByDensity() As String
    TypeByDensity = pTypeByDensity
End Property

' Разновидность песков по плотности
Public Property Let TypeByDensity(name As String)
    Select Case name
        Case DENSITY_TYPE_SAND_HIGH, DENSITY_TYPE_SAND_MIDDLE, DENSITY_TYPE_SAND_LOW:
             pTypeByDensity = name
        Case Else
            Debug.Print name & " is wrong type by density for sand"
    End Select
End Property

' Автоматическое определение песков плотности
' Таблица Б.12 ГОСТ 25100-2011
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

' Разновидность глинистых грунтов по показателю текучести
Public Property Get TypeByLiquidityIndex() As String
    TypeByLiquidityIndex = pTypeByLiquidityIndex
End Property

' Разновидность глинистых грунтов по показателю текучести
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

' Автоматическое определение глинистых грунтов по показателю текучести
' Таблица Б.19 ГОСТ 25100-2011
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





