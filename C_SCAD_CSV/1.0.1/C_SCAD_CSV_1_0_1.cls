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

Const VERSION = "1.0.1"

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