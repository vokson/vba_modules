Attribute VB_Name = "M_VBA_Library_List_v1"
Option Explicit

Public Function getListOfRequiredModules() As Dictionary
    Dim dic As New Dictionary
    
    dic.Item("C_Soil_Database") = "1.1.0"
    
    Set getListOfRequiredModules = dic
End Function
