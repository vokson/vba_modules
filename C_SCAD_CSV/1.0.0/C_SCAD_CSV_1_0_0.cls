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

Const VERSION = "1.0.0"

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
    For i = 1 To col.Count
        
        Set dic = col(i)
        Print #1, _
            CStr(dic.Item("NAME")) & ";" & _
            CStr(dic.Item("IS_NORMATIVE"))  & ";" & _
            CStr(dic.Item("TYPE"))  & ";" & _
            CStr(dic.Item("MODE"))  & ";" & _
            CStr(dic.Item("RELIABILITY_FACTOR"))  & ";" & _
            CStr(dic.Item("LONG_TIME_PART"))

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
    For i = 1 To col.Count
        
        Set dic = col(i)
        Print #1, _
            CStr(dic.Item("NODE")) & ";" & _
            CStr(dic.Item("LOAD_NAME")) & ";" & _
            CStr(dic.Item("FX"))  & ";" & _
            CStr(dic.Item("FY"))  & ";" & _
            CStr(dic.Item("FZ"))  & ";" & _
            CStr(dic.Item("MX"))  & ";" & _
            CStr(dic.Item("MY"))  & ";" & _
            CStr(dic.Item("MZ"))

    Next i

    writeNodeLoads = True

Handler:
    Close #1
End Function