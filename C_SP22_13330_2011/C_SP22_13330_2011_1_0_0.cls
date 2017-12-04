VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_SP50_102_2003"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pTables As C_SP50_102_2003_Tables
Private pFormulas As C_SP50_102_2003_Formulas

Private Sub Class_Initialize()
    Set pTables = New C_SP50_102_2003_Tables
    Set pFormulas = New C_SP50_102_2003_Formulas
End Sub

Private Sub Class_Terminate()
    Set pTables = Nothing
    Set pFormulas = Nothing
End Sub

Public Property Get Tables() As C_SP50_102_2003_Tables
  Set Tables = pTables
End Property

Public Property Get Formulas() As C_SP50_102_2003_Formulas
  Set Formulas = pFormulas
End Property