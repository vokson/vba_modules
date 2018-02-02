VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_SP63_13330_2012_Tables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public GradeAnchor As Variant '// Класс арматуры анкера
Public Aan1 As Double '//
Public GradeSteel As Variant '// Марка стали пластины
Public GradeConcrete As Variant '// Класс бетона
Public Rbt_gb As Variant '//
Public tSteel As Double '// Толщина пластины (мм)
Public nan_Anchor As Integer '// Колличество рядов анкеров вдоль направления сдвигающей силы
Public n_Anchor As Integer '// Колличество анкеров в ряду
Public Z_Anchor As Double '// Расстояние между крайними рядами анкеров
Public Y_Anchor As Double '// Растояние между крайними анкерами в растянутом ряду
Public dAnchor As Double '// Диаметр анкера (мм)
Public La_Anchor As Double '// Длина анкера (в свету между пластинами)
Public Lpl_Anchor As Double '// Длина пластины усиления анкеров
Public tpl_Anchor As Double '// Толщина пластины усиления анкеров
Public Ha_Anchor As Double '// Толщина бетонного элемента
Public lzZ_Anchor As Double '// Расстояние между анкерами вдоль усилия
Public lzY_Anchor As Double '// Расстояние между анкерами поперек усилия
Public Cy_Anchor As Variant '// Расстояние от оси крайнего анкера до грани бетонного элемента по горизонтали (минимальное)
Public Cz_Anchor As Variant '// Расстояние от оси крайнего анкера до грани бетонного элемента по вертикали (минимальное)
Public kConcreteGB As Variant '// Коэффициент условий работы бетона
Public kConcreteN As Variant '// Коэффициент условий работы бетона
Public kConcreteLay As Variant '// Коэффициент условий работы бетона
Public D1 As Variant '// мм
Public D2 As Variant '// мм
Public ea As Double '// По п.8.1.7 СП 63.13330.2012: ea = max[A/30; 10 mm/мм] 



Public Sub Table_Ini()

Dim sheet As Variant
sheet = "Charts"

GradeAnchor = Worksheets(sheet).Cells(5, 6)
GradeSteel = Worksheets(sheet).Cells(7, 6)
GradeConcrete = Worksheets(sheet).Cells(10, 6)
tSteel = Worksheets(sheet).Cells(23, 5)
nan_Anchor = Worksheets(sheet).Cells(23, 6)
n_Anchor = Worksheets(sheet).Cells(23, 7)
lzZ_Anchor = Worksheets(sheet).Cells(23, 8)
lzY_Anchor = Worksheets(sheet).Cells(23, 9)
dAnchor = Worksheets(sheet).Cells(23, 10)
La_Anchor = Worksheets(sheet).Cells(23, 11)
Lpl_Anchor = Worksheets(sheet).Cells(23, 12)
tpl_Anchor = Worksheets(sheet).Cells(23, 13)
Ha_Anchor = Worksheets(sheet).Cells(23, 14)
kConcreteGB = Worksheets(sheet).Cells(13, 3)
kConcreteN = Worksheets(sheet).Cells(14, 3)
kConcreteLay = Worksheets(sheet).Cells(15, 3)
D1 = Worksheets(sheet).Cells(45, 3)
D2 = Worksheets(sheet).Cells(45, 6)

Z_Anchor = lzZ_Anchor * (nan_Anchor - 1)
Y_Anchor = lzY_Anchor * (n_Anchor - 1)
Cz_Anchor = D1 + (Worksheets(sheet).Cells(23, 3) - lzZ_Anchor * (nan_Anchor - 1)) / 2
Cy_Anchor = D2 + (Worksheets(sheet).Cells(23, 4) - lzY_Anchor * (n_Anchor - 1)) / 2
ea = Application.Max(Worksheets(sheet).Cells(23, 3) / 30, 10) / 1000
End Sub

Public Function Rs(clls As Variant)

Dim row As Integer, sheet As Variant
row = 3
sheet = "Factored parameters"

Do While Worksheets(sheet).Cells(row, 1) <> ""
    If Worksheets(sheet).Cells(row, 1) = clls Then
        Rs = Worksheets(sheet).Cells(row, 2)
        Exit Function
    End If
    row = row + 1
Loop
End Function

Public Function Aanj_Anchor(clls As Variant) '// Cуммарная площадь поперечного сечения анкеров крайнего ряда

Dim row As Integer, sheet As Variant
row = 1
sheet = "Assortment"

Do While Worksheets(sheet).Cells(row, 1) <> ""
    If Worksheets(sheet).Cells(row, 1) = clls Then
        Aan1 = Worksheets(sheet).Cells(row, 2) / 100
        Aanj_Anchor = n_Anchor * Aan1 / 10000
        Exit Function
    End If
    row = row + 1
Loop
End Function

Public Function Rb_gb(clls As Variant)

Dim row As Integer, sheet As Variant, rb As Double
row = 17
sheet = "Factored parameters"

Do While Worksheets(sheet).Cells(row, 1) <> ""
    If Worksheets(sheet).Cells(row, 1) = clls Then
        rb = Worksheets(sheet).Cells(row, 2)
        Rb_gb = Gb(kConcreteGB, kConcreteN, kConcreteLay) * rb
        Rbt_gb = Worksheets(sheet).Cells(row, 5) * Gb(kConcreteGB, kConcreteN, kConcreteLay)
        Exit Function
    End If
    row = row + 1
Loop
End Function

Public Function Gb(kgb As Variant, kn As Variant, klay As Variant)
Dim a As Double, b As Double, c As Double
    If kgb = "железобетонный элемент / reinforced concrete" Then
        a = 1
        Else
            a = 0.9
    End If
    If kn = "кратковременные нагрузки / short-term load" Then
        b = 1
        Else
            b = 0.9
    End If
    If klay = "слой бетонированния >1,5м / concreting layer >1,5m" Then
        c = 0.85
        Else
            c = 1
    End If
    Gb = a * b * c
End Function


