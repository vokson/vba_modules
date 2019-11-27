VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_SCAD_CSV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const VERSION = "1.1.2"

' (C_ImprovedCollection) col - Collection of Dictionary
' (string) fileName - полный путь к файлу CSV
' (boolean) result - успешна запись или нет

' Пишет описание загружений SCAD из Dictionary в файл CSV
' Формат словаря 
' (string) NAME - имя загружения
' (boolean) IS_NORMATIVE - загружение нормативное или нет (расчетное)
' (integer) TYPE - тип загружения согласно спецификации SCAD
' (integer) MODE - подтип загружения согласно спецификации SCAD
' (double) RELIABILITY_FACTOR - Yf
' (double) LONG_TIME_PART - доля длительности

Public Function writeLoadings(col As C_ImprovedCollection, fileName As String) As Boolean

    writeLoadings = False

    On Error GoTo Handler

    Open fileName For Output As #1

    Dim i As Long
    Dim dic As Dictionary
    Dim s As String
    For i = 1 To col.Count
        
        Set dic = col(i)

        s = CStr(dic.Item("NAME")) & ";" & _
            CStr(dic.Item("IS_NORMATIVE"))  & ";" & _
            CStr(dic.Item("TYPE"))  & ";" & _
            CStr(dic.Item("MODE"))  & ";" & _
            CStr(dic.Item("RELIABILITY_FACTOR"))  & ";" & _
            CStr(dic.Item("LONG_TIME_PART"))
        
        s = Replace(s, ",", ".")
        Print #1, s

    Next i

    writeLoadings = True

Handler:
    Close #1
End Function

' (C_ImprovedCollection) col - Collection of Dictionary
' (string) fileName - полный путь к файлу CSV
' (boolean) result - успешна запись или нет

' Пишет описание загружений SCAD из Dictionary в файл CSV
' Формат словаря 
' (long) NODE - номер узла
' (string) LOAD_NAME - имя загружения
' (double) FX
' (double) FY
' (double) FZ
' (double) MX
' (double) MY
' (double) MZ

Public Function writeNodeLoads(col As C_ImprovedCollection, fileName As String) As Boolean

    writeNodeLoads = False

    On Error GoTo Handler

    Open fileName For Output As #1

    Dim i As Long
    Dim dic As Dictionary
    Dim s As String
    For i = 1 To col.Count
        
        Set dic = col(i)
        s =  CStr(dic.Item("NODE")) & ";" & _
            CStr(dic.Item("LOAD_NAME")) & ";" & _
            CStr(dic.Item("FX"))  & ";" & _
            CStr(dic.Item("FY"))  & ";" & _
            CStr(dic.Item("FZ"))  & ";" & _
            CStr(dic.Item("MX"))  & ";" & _
            CStr(dic.Item("MY"))  & ";" & _
            CStr(dic.Item("MZ"))

        s = Replace(s, ",", ".")
        Print #1, s
           

    Next i

    writeNodeLoads = True

Handler:
    Close #1
End Function

' (C_ImprovedCollection) col - Collection of Dictionary
' (string) fileName - полный путь к файлу CSV
' (boolean) result - успешна запись или нет

' Пишет описание стержневых элементов SCAD из Dictionary в файл CSV
' Формат словаря 
' (string) TEXT - имя элемента
' (long) TYPE_ELEM
' (long) TYPE_RIGID
' (long) NUM_INSERT
' (long) NUM_SYS_COORD
' (long) NUM_SYS_COORD_EFFORS
' (long) NUM_BED
' (Array of double) LIST_OF_NODES_COORDINATES - массив координат узлов в формате X1, Y1, Z1, X2, Y2, Z2, X3..

Public Function writeElements(col As C_ImprovedCollection, fileName As String) As Boolean

    writeElements = False

    On Error GoTo Handler

    Open fileName For Output As #1

    Dim i As Long
    Dim dic As Dictionary
    Dim s As String
    For i = 1 To col.Count
        
        Set dic = col(i)
        s =  CStr(dic.Item("TEXT")) & ";" & _
            CStr(dic.Item("TYPE_ELEM")) & ";" & _
            CStr(dic.Item("TYPE_RIGID"))  & ";" & _
            CStr(dic.Item("NUM_INSERT"))  & ";" & _
            CStr(dic.Item("NUM_SYS_COORD"))  & ";" & _
            CStr(dic.Item("NUM_SYS_COORD_EFFORS"))  & ";" & _
            CStr(dic.Item("NUM_BED"))  & ";" & _
            Join(dic.Item("LIST_OF_NODES_COORDINATES"), ";")

        s = Replace(s, ",", ".")
        Print #1, s
           

    Next i

    writeElements = True

Handler:
    Close #1
End Function

' (C_ImprovedCollection) col - Collection of Dictionary
' (string) fileName - полный путь к файлу CSV
' (boolean) result - успешна запись или нет

' Пишет описание узлов SCAD из Dictionary в файл CSV
' Формат словаря 
' (long) NUM_NODE
' (string) TEXT - имя элемента
' (double) X
' (double) Y
' (double) Z

Public Function readNodes(col As C_ImprovedCollection, fileName As String) As Boolean

    readNodes = False

    On Error GoTo Handler

    Dim math As New C_Math
    Dim result As New C_ImprovedCollection
    Dim rowWithNames As String
    Dim arrWithNames As Variant

    Open fileName For Input As #1

    ' Читаем первую строку с наименованиями столбцов
    Line Input #1, rowWithNames
    arrWithNames = Split(rowWithNames, ";")

    ' Ищем индексы в массиве для нужных столбцов
    Dim NumNode_index As Long
    Dim Text_index As Long
    Dim X_index As Long
    Dim Y_index As Long
    Dim Z_index As Long

    NumNode_index = math.findIdexOfValueInArray("NumNode", arrWithNames)
    Text_index = math.findIdexOfValueInArray("Text", arrWithNames)
    X_index = math.findIdexOfValueInArray("x", arrWithNames)
    Y_index = math.findIdexOfValueInArray("y", arrWithNames)
    Z_index = math.findIdexOfValueInArray("z", arrWithNames)

    Dim text As String
    Dim arr As Variant
    Dim dic As Dictionary
    Dim i As Long


    Do While Not EOF(1)
        Line Input #1, text
        arr = Split(text, ";")

        Set dic = New Dictionary

        dic.Item("NUM_NODE") = CLng(arr(NumNode_index))
        dic.Item("TEXT") = CStr(arr(Text_index))
        dic.Item("X") = Val(arr(X_index))
        dic.Item("Y") = Val(arr(Y_index))
        dic.Item("Z") = Val(arr(Z_index))

        result.Add dic

        Set dic = Nothing
    Loop

    Set col = result
    
    Set result = Nothing
    readNodes = True

Handler:
    Close #1
End Function

' (C_ImprovedCollection) col - Collection of Dictionary
' (string) fileName - полный путь к файлу CSV
' (boolean) result - успешна запись или нет

' Пишет описание стержневых элементов SCAD из Dictionary в файл CSV
' Формат словаря 
' (long) NUM_ELEM
' (string) TEXT - имя элемента
' (long) TYPE_ELEM
' (long) TYPE_RIGID
' (long) NUM_INSERT
' (long) NUM_SYS_COORD
' (long) NUM_SYS_COORD_EFFORS
' (long) NUM_BED
' (Array of double) LIST_OF_NODES - массив номеров узлов, принадлежащих данному элементу

Public Function readElements(col As C_ImprovedCollection, fileName As String) As Boolean

    readElements = False

    On Error GoTo Handler

    Dim math As New C_Math
    Dim result As New C_ImprovedCollection
    Dim rowWithNames As String
    Dim arrWithNames As Variant

    Open fileName For Input As #1

    ' Читаем первую строку с наименованиями столбцов
    Line Input #1, rowWithNames
    arrWithNames = Split(rowWithNames, ";")

    ' Ищем индексы в массиве для нужных столбцов
    Dim NumElem_index As Long
    Dim Text_index As Long
    Dim TypeElem_index As Long
    Dim TypeRigid_index As Long
    Dim NumInsert_index As Long
    Dim NumSysCoord_index As Long
    Dim NumSysCoordEffors_index As Long
    Dim NumBed_index As Long
    Dim QuantityNode_index As Long
    Dim ListNode_index As Long

    NumElem_index = math.findIdexOfValueInArray("NumElem", arrWithNames)
    Text_index = math.findIdexOfValueInArray("Text", arrWithNames)
    TypeElem_index = math.findIdexOfValueInArray("TypeElem", arrWithNames)
    TypeRigid_index = math.findIdexOfValueInArray("TypeRigid", arrWithNames)
    NumInsert_index = math.findIdexOfValueInArray("NumInsert", arrWithNames)
    NumSysCoord_index = math.findIdexOfValueInArray("NumSysCoord", arrWithNames)
    NumSysCoordEffors_index = math.findIdexOfValueInArray("NumSysCoordEffors", arrWithNames)
    NumBed_index = math.findIdexOfValueInArray("NumBed", arrWithNames)
    QuantityNode_index = math.findIdexOfValueInArray("QuantityNode", arrWithNames)
    ListNode_index = math.findIdexOfValueInArray("ListNode", arrWithNames)

    Dim text As String
    Dim arr As Variant
    Dim dic As Dictionary
    Dim i As Long
    Dim arrOfNodes() As Long

    Do While Not EOF(1)
        Line Input #1, text
        arr = Split(text, ";")

        Set dic = New Dictionary

        dic.Item("NUM_ELEM") = CLng(arr(NumElem_index))
        dic.Item("TEXT") = CStr(arr(Text_index))
        dic.Item("TYPE_ELEM") = CLng(arr(TypeElem_index))
        dic.Item("TYPE_RIGID") = CLng(arr(TypeRigid_index))
        dic.Item("NUM_INSERT") = CLng(arr(NumInsert_index))
        dic.Item("NUM_SYS_COORD") = CLng(arr(NumSysCoord_index))
        dic.Item("NUM_SYS_COORD_EFFORS") = CLng(arr(NumSysCoordEffors_index))
        dic.Item("NUM_BED") = CLng(arr(NumBed_index))
        dic.Item("QUANTITY_NODE") = CLng(arr(QuantityNode_index))

        Dim countOfNodes As Long
        countOfNodes = CLng(arr(QuantityNode_index))
        ReDim arrOfNodes(0)

        For i = ListNode_index To ListNode_index + countOfNodes - 1
            arrOfNodes(UBound(arrOfNodes)) = CLng(arr(i))
            ReDim Preserve arrOfNodes(UBound(arrOfNodes) + 1)
        Next i

        ReDim Preserve arrOfNodes(UBound(arrOfNodes) - 1)

        dic.Item("LIST_OF_NODES") = arrOfNodes

        result.Add dic

        Set dic = Nothing
    Loop

    Set col = result
    
    Set result = Nothing
    readElements = True

Handler:
    Close #1
End Function

' (C_ImprovedCollection) col - Collection of Dictionary
' (string) fileName - полный путь к файлу CSV
' (boolean) result - успешна запись или нет

' Пишет описание стержневых элементов SCAD из Dictionary в файл CSV
' Формат словаря 
' (string) TEXT - имя элемента
' (long) TYPE_ELEM - тип элемента
' (double) LOCAL_Z_ANGLE - угол поворота локальной оси Z в градусах
' (string) RIGID_DESCRIPTION - описание типа жесткости в формате Языка Архивации SCAD
' (Array of double) LIST_OF_NODES_COORDINATES - массив координат узлов в формате X1, Y1, Z1, X2, Y2, Z2, X3..

Public Function writePedestals(col As C_ImprovedCollection, fileName As String) As Boolean

    writePedestals = False

    On Error GoTo Handler

    Open fileName For Output As #1

    Dim i As Long
    Dim dic As Dictionary
    Dim s As String
    For i = 1 To col.Count
        
        Set dic = col(i)
        s =  CStr(dic.Item("TEXT")) & ";" & _
            CStr(dic.Item("TYPE_ELEM")) & ";" & _
            CStr(dic.Item("LOCAL_Z_ANGLE"))  & ";" & _
            CStr(dic.Item("RIGID_DESCRIPTION"))  & ";" & _
            CStr(dic.Item("FOOTING_THICKNESS"))  & ";" & _
            Join(dic.Item("LIST_OF_NODES_COORDINATES"), ";")

        s = Replace(s, ",", ".")
        Print #1, s
           

    Next i

    writePedestals = True

Handler:
    Close #1
End Function

' (C_ImprovedCollection) col - Collection of Dictionary
' (string) fileName - полный путь к файлу CSV
' (boolean) result - успешна запись или нет

' Пишет описание стержневых элементов SCAD из Dictionary в файл CSV
' Формат словаря 
' (string) TEXT - имя элемента
' См. SCAD Plugin API функция GetArmElemRodPart

Public Function writePedestalsReinforcement(col As C_ImprovedCollection, fileName As String) As Boolean

    writePedestalsReinforcement = False

    On Error GoTo Handler

    Open fileName For Output As #1

    Dim i As Long
    Dim dic As Dictionary
    Dim s As String
    For i = 1 To col.Count

        Set dic = col(i)
        s =  CStr(dic.Item("TEXT")) & ";" & _
            CStr(dic.Item("dS1L1_1")) & ";" & _
            CStr(dic.Item("nS1L1_1")) & ";" & _
            CStr(dic.Item("dS2L1_1")) & ";" & _
            CStr(dic.Item("nS2L1_1")) & ";" & _
            CStr(dic.Item("dS1L1_2")) & ";" & _
            CStr(dic.Item("nS1L1_2")) & ";" & _
            CStr(dic.Item("dS2L1_2")) & ";" & _
            CStr(dic.Item("nS2L1_2")) & ";" & _
            CStr(dic.Item("dS3L1_1")) & ";" & _
            CStr(dic.Item("nS3L1_1")) & ";" & _
            CStr(dic.Item("dS4L1_1")) & ";" & _
            CStr(dic.Item("nS4L1_1")) & ";" & _
            CStr(dic.Item("dSw")) & ";" & _
            CStr(dic.Item("nSw")) & ";" & _
            CStr(dic.Item("StepSw")) & ";" & _
            CStr(dic.Item("dSw2")) & ";" & _
            CStr(dic.Item("nSw2")) & ";" & _
            CStr(dic.Item("StepSw2")) 

        s = Replace(s, ",", ".")
        Print #1, s
           

    Next i

    writePedestalsReinforcement = True

Handler:
    Close #1
End Function

' (C_ImprovedCollection) col - Collection of Dictionary
' (string) fileName - полный путь к файлу CSV
' (boolean) result - успешна запись или нет

' Пишет описание жесткостей SCAD из Dictionary в файл CSV
' Формат словаря 
' (long) NUM_RIGID
' (string) TEXT - имя элемента
' (long) QUANTITY_IDENT - кол-во описательных характеристик в DESCRIPTION
' (string) DESCRIPTION - описание типа жесткости
' (long) QUANTITY_ELEM - кол-во элементов, которым присовен данный тип жесткости
' (Array of double) LIST_ELEM - массив номеров элементов с данным типом жесткости

Public Function readRigities(col As C_ImprovedCollection, fileName As String) As Boolean

    readRigities = False

    On Error GoTo Handler

    Dim math As New C_Math
    Dim result As New C_ImprovedCollection
    Dim rowWithNames As String
    Dim arrWithNames As Variant

    Open fileName For Input As #1

    ' Читаем первую строку с наименованиями столбцов
    Line Input #1, rowWithNames
    arrWithNames = Split(rowWithNames, ";")

    ' Ищем индексы в массиве для нужных столбцов
    Dim NumRigid_index As Long
    Dim Text_index As Long
    Dim QuantityIdent_index As Long
    Dim Description_index As Long
    Dim QuantityElem_index As Long
    Dim ListElem_index As Long

    NumRigid_index = math.findIdexOfValueInArray("NumRigid", arrWithNames)
    Text_index = math.findIdexOfValueInArray("Text", arrWithNames)
    QuantityIdent_index = math.findIdexOfValueInArray("QuantityIdent", arrWithNames)
    Description_index = math.findIdexOfValueInArray("Description", arrWithNames)
    QuantityElem_index = math.findIdexOfValueInArray("QuantityElem", arrWithNames)
    ListElem_index = math.findIdexOfValueInArray("ListElem", arrWithNames)

    Dim text As String
    Dim arr As Variant
    Dim dic As Dictionary
    Dim i As Long
    Dim arrOfElems() As Long


    Do While Not EOF(1)
        Line Input #1, text
        arr = Split(text, ";")

        Set dic = New Dictionary

        dic.Item("NUM_RIGID") = CLng(arr(NumRigid_index))
        dic.Item("TEXT") = CStr(arr(Text_index))
        dic.Item("QUANTITY_IDENT") = CLng(arr(QuantityIdent_index))
        dic.Item("DESCRIPTION") = CStr(arr(Description_index))
        dic.Item("QUANTITY_ELEM") = CLng(arr(QuantityElem_index))

        Dim countOfElems As Long
        countOfElems = CLng(arr(QuantityElem_index))
        ReDim arrOfElems(0)

        For i = ListElem_index To ListElem_index + countOfElems - 1
            arrOfElems(UBound(arrOfElems)) = CLng(arr(i))
            ReDim Preserve arrOfElems(UBound(arrOfElems) + 1)
        Next i

        ReDim Preserve arrOfElems(UBound(arrOfElems) - 1)

        dic.Item("LIST_ELEM") = arrOfElems

        result.Add dic

        Set dic = Nothing
    Loop

    Set col = result
    
    Set result = Nothing
    readRigities = True

Handler:
    Close #1
End Function

' (C_ImprovedCollection) col - Collection of Dictionary
' (string) fileName - полный путь к файлу CSV
' (boolean) result - успешна запись или нет

' Пишет описание местных систем координат элементов SCAD из Dictionary в файл CSV
' Формат словаря 
' (long) NUM_ELEM - номер элемента
' (long) NUM_SYSTEM_COORD_ELEM - номер системы координат
' (string) TEXT - имя элемента
' (long) TYPE - тип системы координат
' (long) QUANTITY_DATA - кол-во элементов описания типа системы координат
' (Array of double) LIST_DATA - массив описаний

Public Function readSystemCoordElem(col As C_ImprovedCollection, fileName As String) As Boolean

    readSystemCoordElem = False

    On Error GoTo Handler

    Dim math As New C_Math
    Dim result As New C_ImprovedCollection
    Dim rowWithNames As String
    Dim arrWithNames As Variant

    Open fileName For Input As #1

    ' Читаем первую строку с наименованиями столбцов
    Line Input #1, rowWithNames
    arrWithNames = Split(rowWithNames, ";")

    ' Ищем индексы в массиве для нужных столбцов
    Dim NumElem_index As Long
    Dim NumSystemCoordElem_index As Long
    Dim Text_index As Long
    Dim Type_index As Long
    Dim QuantityData_index As Long
    Dim ListData_index As Long

    NumElem_index = math.findIdexOfValueInArray("NumElem", arrWithNames)
    NumSystemCoordElem_index = math.findIdexOfValueInArray("NumSystemCoordElem", arrWithNames)
    Text_index = math.findIdexOfValueInArray("Text", arrWithNames)
    Type_index = math.findIdexOfValueInArray("Type", arrWithNames)
    QuantityData_index = math.findIdexOfValueInArray("QuantityData", arrWithNames)
    ListData_index = math.findIdexOfValueInArray("ListData", arrWithNames)

    Dim text As String
    Dim arr As Variant
    Dim dic As Dictionary
    Dim i As Long
    Dim arrOfData() As Long


    Do While Not EOF(1)
        Line Input #1, text
        arr = Split(text, ";")

        Set dic = New Dictionary

        dic.Item("NUM_ELEM") = CLng(arr(NumElem_index))
        dic.Item("NUM_SYSTEM_COORD_ELEM") = CLng(arr(NumSystemCoordElem_index))
        dic.Item("TEXT") = CStr(arr(Text_index))
        dic.Item("TYPE") = CLng(arr(Type_index))
        dic.Item("QUANTITY_DATA") = CLng(arr(QuantityData_index))

        Dim countOfData As Long
        countOfData = CLng(arr(QuantityData_index))
        ReDim arrOfData(0)

        For i = ListData_index To ListData_index + countOfData - 1
            arrOfData(UBound(arrOfData)) = CLng(arr(i))
            ReDim Preserve arrOfData(UBound(arrOfData) + 1)
        Next i

        ReDim Preserve arrOfData(UBound(arrOfData) - 1)

        dic.Item("LIST_DATA") = arrOfData

        result.Add dic

        Set dic = Nothing
    Loop

    Set col = result
    
    Set result = Nothing
    readSystemCoordElem = True

Handler:
    Close #1
End Function

' (C_ImprovedCollection) col - Collection of Dictionary
' (string) fileName - полный путь к файлу CSV
' (boolean) result - успешна запись или нет

' Пишет усилия от комбинаций загружений SCAD из Dictionary в файл CSV
' Формат словаря 
' (long) NUM_ELEM - номер элемента
' (long) NUM_SYSTEM_COORD_ELEM - номер системы координат
' (string) TEXT - имя элемента
' (long) TYPE - тип системы координат
' (long) QUANTITY_DATA - кол-во элементов описания типа системы координат
' (Array of double) LIST_DATA - массив описаний

Public Function readCombinationEfforts(col As C_ImprovedCollection, fileName As String) As Boolean

    readCombinationEfforts = False

    On Error GoTo Handler

    Dim math As New C_Math
    Dim result As New C_ImprovedCollection
    Dim rowWithNames As String
    Dim arrWithNames As Variant

    Open fileName For Input As #1

    ' Читаем первую строку с наименованиями столбцов
    Line Input #1, rowWithNames
    arrWithNames = Split(rowWithNames, ";")

    ' Ищем индексы в массиве для нужных столбцов
    Dim Element_index As Long
    Dim Point_index As Long
    Dim Combination_index As Long
    Dim Efforts_index As Long

    Element_index = math.findIdexOfValueInArray("Element", arrWithNames)
    Point_index = math.findIdexOfValueInArray("Point", arrWithNames)
    Combination_index = math.findIdexOfValueInArray("Combination", arrWithNames)
    ' Вычисляем индекс, с которого начинаются усилия
    Efforts_index = CLng(math.max(math.max(CDbl(Element_index), CDbl(Point_index)),CDbl( Combination_index))) + 1

    Dim text As String
    Dim arr As Variant
    Dim dic As Dictionary
    Dim i As Long


    Do While Not EOF(1)
        Line Input #1, text
        arr = Split(text, ";")

        Set dic = New Dictionary

        dic.Item("ELEMENT") = CLng(arr(Element_index))
        dic.Item("POINT") = CLng(arr(Point_index))
        dic.Item("COMBINATION") = CLng(arr(Combination_index))

        For i = Efforts_index To UBound(arrWithNames)
            dic.Item(CStr(arrWithNames(i))) = Val(arr(i))
        Next i

        result.Add dic

        Set dic = Nothing
    Loop

    Set col = result
    
    Set result = Nothing
    readCombinationEfforts = True

Handler:
    Close #1
End Function