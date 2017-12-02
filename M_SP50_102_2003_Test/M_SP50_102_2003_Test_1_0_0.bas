Attribute VB_Name = "M_SP50_102_2003_Test"
Option Explicit

Public Function getTablesE2(pileCase As Integer, length As Double, parameter As String) As Double
    Dim sp : Set sp = New C_SP50_102_2003
    getTablesE2 = sp.Tables.E2(pileCase, length, parameter)
    Set sp = Nothing
End Function

Public Function getTablesE3(z As Double, parameter As String) As Double
    Dim sp : Set sp = New C_SP50_102_2003
    getTablesE3 = sp.Tables.E3(z, parameter)
    Set sp = Nothing
End Function