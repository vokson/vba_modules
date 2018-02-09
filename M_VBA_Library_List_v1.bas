Attribute VB_Name = "M_VBA_Library_List_v1"
Option Explicit

Public Function getListOfRequiredModules() As Dictionary
    Dim dic As New Dictionary
    
    dic.Item("C_Soil") = "2_1_0"
    dic.Item("C_Math") = "1_2_0"
    dic.Item("C_SP22_13330_2011_Tables") = "1_1_0"
    dic.Item("C_SP22_13330_2011_Formulas") = "1_0_0"
    dic.Item("C_SP22_13330_2011") = "1_1_0"
    dic.Item("M_SP22_13330_2011_Test") = "1_1_0"
    
    Set getListOfRequiredModules = dic
End Function
