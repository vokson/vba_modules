Attribute VB_Name = "M_SP63_13330_2012_Test"
Option Explicit

Public Function getFormula8_55(Rb As Double, b As Double, h0 As Double, FIn As Double) As Double
    
    Dim sp As New C_SP63_13330_2012
        getFormula8_55 = sp.Formulas.f8_55(Rb, b, h0, FIn)
    Set sp = Nothing

End Function

Public Function getFormula8_57(Rbt As Double, b As Double, h0 As Double, C As Double, FIn As Double) As Double
    
    Dim sp As New C_SP63_13330_2012
        getFormula8_57 = sp.Formulas.f8_57(Rbt, b, h0, C, FIn)
    Set sp = Nothing

End Function

Public Function getFormula8_58(qsw As Double, C As Double) As Double
    
    Dim sp As New C_SP63_13330_2012
        getFormula8_58 = sp.Formulas.f8_58(qsw, C)
    Set sp = Nothing

End Function

Public Function getFormula8_59(Rsw As Double, Asw As Double, sw As Double) As Double
    
    Dim sp As New C_SP63_13330_2012
        getFormula8_59 = sp.Formulas.f8_59(Rsw, Asw, sw)
    Set sp = Nothing

End Function

Public Function getClause8_1_34(stress As Double, Rb As Double, Rbt As Double) As Double
    
    Dim sp As New C_SP63_13330_2012
        getClause8_1_34 = sp.Clauses.c8_1_34(stress, Rb, Rbt)
    Set sp = Nothing

End Function