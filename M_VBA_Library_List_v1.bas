Attribute VB_Name = "M_VBA_Library_List_v1"
Option Explicit

Public Function getListOfRequiredModules() As Dictionary
    Dim dic As New Dictionary
    '
    dic.Item("C_SP63_13330_2012_Formulas") = "2_0_1"
    dic.Item("C_SP63_13330_2012_Tables") = "2_0_1"
    dic.Item("M_SP63_13330_2012_Test") = "2_0_1"
    
    Set getListOfRequiredModules = dic
End Function
