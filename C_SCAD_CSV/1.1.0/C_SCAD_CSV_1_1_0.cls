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

Const VERSION = "1.1.0"

' (C_ImprovedCollection) col - Collection of Dictionary
' (string) fileName - ������ ���� � ����� CSV
' (boolean) result - ������� ������ ��� ���

' ����� �������� ���������� SCAD �� Dictionary � ���� CSV
' ������ ������� 
' (string) NAME - ��� ����������
' (boolean) IS_NORMATIVE - ���������� ����������� ��� ��� (���������)
' (integer) TYPE - ��� ���������� �������� ������������ SCAD
' (integer) MODE - ������ ���������� �������� ������������ SCAD
' (double) RELIABILITY_FACTOR - Yf
' (double) LONG_TIME_PART - ���� ������������

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
' (string) fileName - ������ ���� � ����� CSV
' (boolean) result - ������� ������ ��� ���

' ����� �������� ���������� SCAD �� Dictionary � ���� CSV
' ������ ������� 
' (long) NODE - ����� ����
' (string) LOAD_NAME - ��� ����������
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
' (string) fileName - ������ ���� � ����� CSV
' (boolean) result - ������� ������ ��� ���

' ����� �������� ���������� ��������� SCAD �� Dictionary � ���� CSV
' ������ ������� 
' (string) TEXT - ��� ��������
' (long) TYPE_ELEM
' (long) TYPE_RIGID
' (long) NUM_INSERT
' (long) NUM_SYS_COORD
' (long) NUM_SYS_COORD_EFFORS
' (long) NUM_BED
' (Array of double) LIST_OF_NODES_COORDINATES - ������ ��������� ����� � ������� X1, Y1, Z1, X2, Y2, Z2, X3..

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
' (string) fileName - ������ ���� � ����� CSV
' (boolean) result - ������� ������ ��� ���

' ����� �������� ���������� ��������� SCAD �� Dictionary � ���� CSV
' ������ ������� 
' (long) NUM_ELEM
' (string) TEXT - ��� ��������
' (long) TYPE_ELEM
' (long) TYPE_RIGID
' (long) NUM_INSERT
' (long) NUM_SYS_COORD
' (long) NUM_SYS_COORD_EFFORS
' (long) NUM_BED
' (Array of double) LIST_OF_NODES - ������ ������� �����, ������������� ������� ��������

Public Function readElements(col As C_ImprovedCollection, fileName As String) As Boolean

    readElements = False

    On Error GoTo Handler

    Dim math As New C_Math
    Dim result As New C_ImprovedCollection
    Dim rowWithNames As String
    Dim arrWithNames As Variant

    Open fileName For Input As #1

    ' ������ ������ ������ � �������������� ��������
    Line Input #1, rowWithNames
    arrWithNames = Split(rowWithNames, ";")

    ' ���� ������� � ������� ��� ������ ��������
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
    Dim arrOfNodes() As Integer


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
        dic.Item("NUM_BED") = CLng(arr(QuantityNode_index))

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
' (string) fileName - ������ ���� � ����� CSV
' (boolean) result - ������� ������ ��� ���

' ����� �������� ���������� ��������� SCAD �� Dictionary � ���� CSV
' ������ ������� 
' (string) TEXT - ��� ��������
' (long) TYPE_ELEM - ��� ��������
' (double) LOCAL_Z_ANGLE - ���� �������� ��������� ��� Z � ��������
' (string) RIGID_DESCRIPTION - �������� ���� ��������� � ������� ����� ��������� SCAD
' (Array of double) LIST_OF_NODES_COORDINATES - ������ ��������� ����� � ������� X1, Y1, Z1, X2, Y2, Z2, X3..

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
' (string) fileName - ������ ���� � ����� CSV
' (boolean) result - ������� ������ ��� ���

' ����� �������� ���������� ��������� SCAD �� Dictionary � ���� CSV
' ������ ������� 
' (string) TEXT - ��� ��������
' ��. SCAD Plugin API ������� GetArmElemRodPart

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