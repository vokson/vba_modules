Attribute VB_Name = "M_VBA_Library_List_v1"
Option Explicit

Public Function getListOfRequiredModules() As Dictionary
    Dim dic As New Dictionary
    
    dic.Item("C_Soil") = "1_0_0"
    dic.Item("C_Excel_Worksheet") = "1_0_0"
    
    Set getListOfRequiredModules = dic
End Function
