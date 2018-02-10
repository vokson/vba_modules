VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_SP63_13330_2012"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Private pTables As C_SP63_13330_2012_Tables
Private pFormulas As C_SP63_13330_2012_Formulas
Private pClauses As C_SP63_13330_2012_Clauses

Private Sub Class_Initialize()
    'Set pTables = New C_SP63_13330_2012_Tables
    Set pFormulas = New C_SP63_13330_2012_Formulas
    Set pClauses = New C_SP63_13330_2012_Clauses
End Sub

Private Sub Class_Terminate()
    'Set pTables = Nothing
    Set pFormulas = Nothing
    Set pClauses = Nothing
End Sub

' Public Property Get Tables() As C_SP63_13330_2012_Tables
'   Set Tables = pTables
' End Property

Public Property Get Formulas() As C_SP63_13330_2012_Formulas
  Set Formulas = pFormulas
End Property

Public Property Get Clauses() As C_SP63_13330_2012_Clauses
  Set Clauses = pClauses
End Property