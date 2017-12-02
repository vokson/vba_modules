Attribute VB_Name = "M_SP50_102_2003_Test"
Option Explicit

Public Function getTablesE1( _
        isDrivenPile As Boolean, _
        typeOfSoil As String, _
        subtypeOfSoil As String, _
        densityType As String, _
        IL As Double, _
        e As Double _
    ) As Double 'ÍÕ/Ï4

    Dim sp As New C_SP50_102_2003
    Dim soil As New C_Soil

    soil.TypeBySize = typeOfSoil
    soil.SubtypeBySize = subtypeOfSoil
    soil.TypeByDensity = densityType
    soil.LiquidityIndex = IL
    soil.VoidRatio = e

    getTablesE1 = sp.Tables.E1(isDrivenPile, soil)

    Set sp = Nothing
    Set soil = Nothing
End Function

Public Function getTablesE2(pileCase As Integer, length As Double, parameter As String) As Double
   Dim sp As New C_SP50_102_2003
    getTablesE2 = sp.Tables.E2(pileCase, length, parameter)
    Set sp = Nothing
End Function

Public Function getTablesE3(z As Double, parameter As String) As Double
    Dim sp As New C_SP50_102_2003
    getTablesE3 = sp.Tables.E3(z, parameter)
    Set sp = Nothing
End Function

Public Function getFormulasE8(K As Double, bp As Double, E As Double, I As Double) As Double
    Dim sp As New C_SP50_102_2003
    getFormulasE8 = sp.Formulas.E8(K, bp, E, I)
    Set sp = Nothing
End Function

Public Function getFormulasE12(H0 As Double, M0 As Double, eHH As Double, eHM As Double) As Double
    Dim sp As New C_SP50_102_2003
    getFormulasE12 = sp.Formulas.E12(H0, M0, eHH, eHM)
    Set sp = Nothing
End Function

Public Function getFormulasE13(H0 As Double, M0 As Double, eMH As Double, eMM As Double) As Double
   Dim sp As New C_SP50_102_2003
    getFormulasE13= sp.Formulas.E13(H0, M0, eMH, eMM)
    Set sp = Nothing
End Function

Public Function getFormulasE14(A0 As Double, alpha_e As Double, E As Double, I As Double) As Double
    Dim sp As New C_SP50_102_2003
    getFormulasE14= sp.Formulas.E14(A0, alpha_e, E, I)
    Set sp = Nothing
End Function

Public Function getFormulasE15(B0 As Double, alpha_e As Double, E As Double, I As Double) As Double
    Dim sp As New C_SP50_102_2003
    getFormulasE15= sp.Formulas.E15(B0, alpha_e, E, I)
    Set sp = Nothing
End Function

Public Function getFormulasE16(C0 As Double, alpha_e As Double, E As Double, I As Double) As Double
    Dim sp As New C_SP50_102_2003
    getFormulasE16= sp.Formulas.E16(C0, alpha_e, E, I)
    Set sp = Nothing
End Function

Public Function getFormulasE17( _
        z As Double, _
        nu1 As Double, _
        nu2 As Double, _
        psi As Double, _
        Y1 As Double, _
        FI1 As Double, _
        C1 As Double _
    ) As Double

    Dim sp As New C_SP50_102_2003
    getFormulasE17 = sp.Formulas.E17(z, nu1, nu2, psi, Y1, FI1, C1)
    Set sp = Nothing

End Function

Public Function getFormulasE19( _
        K As Double, _
        alpha_e As Double, _ 
        z As Double, _
        E As Double, _
        I As Double, _
        U0 As Double, _
        W0 As Double, _
        M0 As Double, _
        H0 As Double, _
        A1 As Double, _
        B1 As Double, _
        C1 As Double, _
        D1 As Double _
    ) As Double ' Íœ‡

    Dim sp As New C_SP50_102_2003
    getFormulasE19 = sp.Formulas.E19(K, alpha_e, z, E, I, U0, W0, M0, H0, A1, B1, C1, D1)
    Set sp = Nothing
End Function