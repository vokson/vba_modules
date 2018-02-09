Attribute VB_Name = "M_SP22_13330_2011_Test"
Option Explicit

Public Function getTables5_10(typeOfSoil As String, IL As Double) As Double

    Dim sp As New C_SP22_13330_2011
    Dim soil As New C_Soil

    soil.ClassOfSoil = "ƒ»—œ≈–—Õ€…"
    soil.TypeBySize = typeOfSoil
    soil.LiquidityIndex = IL

    getTables5_10 = sp.Tables.t5_10(soil)

    Set sp = Nothing
    Set soil = Nothing
End Function

Public Function getTables5_4_Yc1(typeOfSoil As String, subtypeOfSoil As String, typeDensity As String, Saturation As String, IL As Double) As Double

    Dim sp As New C_SP22_13330_2011
    Dim soil As New C_Soil

    soil.ClassOfSoil = "ƒ»—œ≈–—Õ€…"
    soil.TypeBySize = typeOfSoil
    soil.SubtypeBySize = subtypeOfSoil
    soil.TypeByDensity = typeDensity
    soil.TypeByDegreeOfSaturation = Saturation
    soil.LiquidityIndex = IL

    getTables5_4_Yc1 = sp.Tables.t5_4_Yc1(soil)

    Set sp = Nothing
    Set soil = Nothing
End Function

Public Function getTables5_4_Yc2(typeOfSoil As String, subtypeOfSoil As String, typeDensity As String, Saturation As String, IL As Double, L_H As Double, ModelFlexible As Boolean) As Double

    Dim sp As New C_SP22_13330_2011
    Dim soil As New C_Soil

    soil.ClassOfSoil = "ƒ»—œ≈–—Õ€…"
    soil.TypeBySize = typeOfSoil
    soil.SubtypeBySize = subtypeOfSoil
    soil.TypeByDensity = typeDensity
    soil.TypeByDegreeOfSaturation = Saturation
    soil.LiquidityIndex = IL
    

    getTables5_4_Yc2 = sp.Tables.t5_4_Yc2(soil, L_H, ModelFlexible)

    Set sp = Nothing
    Set soil = Nothing
End Function

Public Function getTables5_5_My(InternalFrictionAngle_2 As Double)
    Dim sp As New C_SP22_13330_2011

    getTables5_5_My = sp.Tables.t5_5("My", InternalFrictionAngle_2)
    
    Set sp = Nothing
End Function

Public Function getTables5_5_Mq(InternalFrictionAngle_2 As Double)
    Dim sp As New C_SP22_13330_2011
    
    getTables5_5_Mq = sp.Tables.t5_5("Mq", InternalFrictionAngle_2)
    
    Set sp = Nothing
End Function

Public Function getTables5_5_Mc(InternalFrictionAngle_2 As Double)
    Dim sp As New C_SP22_13330_2011
    
    getTables5_5_Mc = sp.Tables.t5_5("Mc", InternalFrictionAngle_2)

    Set sp = Nothing
End Function

Public Function getTables5_8(ksi As Double, L_B As Double, typeOfFoundation As String)
    Dim sp As New C_SP22_13330_2011

    getTables5_8 = sp.Tables.t5_8(ksi, L_B, typeOfFoundation)

    Set sp = Nothing
End Function

Public Function getTables5_9(L_B As Double, typeOfFoundation As String)
    Dim sp As New C_SP22_13330_2011

    getTables5_9 = sp.Tables.t5_9(L_B, typeOfFoundation)

    Set sp = Nothing
End Function

Public Function getTables5_12y(InternalFrictionAngle_1 As Double, delta As Double)
    Dim sp As New C_SP22_13330_2011

    getTables5_12y = sp.Tables.t5_12("Ny", InternalFrictionAngle_1, delta)

    Set sp = Nothing
End Function

Public Function getTables5_12q(InternalFrictionAngle_1 As Double, delta As Double)
    Dim sp As New C_SP22_13330_2011

    getTables5_12q = sp.Tables.t5_12("Nq", InternalFrictionAngle_1, delta)

    Set sp = Nothing
End Function

Public Function getTables5_12c(InternalFrictionAngle_1 As Double, delta As Double)
    Dim sp As New C_SP22_13330_2011

    getTables5_12c = sp.Tables.t5_12("Nc", InternalFrictionAngle_1, delta)

    Set sp = Nothing
End Function

Public Function getFormulas5_7(Yc1 As Double, Yc2 As Double, k As Double, My As Double, Mq As Double, _
                        Mc As Double, b As Double, Y2 As Double, Y2_ As Double, C2 As Double, _
                        d1 As Double, db As Double) As Double
    Dim sp As New C_SP22_13330_2011
        getFormulas5_7 = sp.Formulas.f5_7(Yc1, Yc2, k, My, Mq, Mc, b, Y2, Y2_, C2, d1, db)
    Set sp = Nothing
End Function

Public Function getFormulas5_8(hs As Double, hcf As Double, Ycf As Double, Y2_ As Double)
    Dim sp As New C_SP22_13330_2011
        getFormulas5_8 = sp.Formulas.f5_8(hs, hcf, Ycf, Y2_)
    Set sp = Nothing
End Function

Public Function getFormulas5_24(d As Double, ke As Double, N As Double, e As Double, a As Double)
    Dim sp As New C_SP22_13330_2011
        getFormulas5_24 = sp.Formulas.f5_24(d, ke, N, e, a)
    Set sp = Nothing
End Function

Public Function getFormulas5_25(nu As Double, e As Double)
    Dim sp As New C_SP22_13330_2011
        getFormulas5_25 = sp.Formulas.f5_25(nu, e)
    Set sp = Nothing
End Function

Public Function getFormulas5_29(value As Double, e As Double)
    Dim sp As New C_SP22_13330_2011
        getFormulas5_29 = sp.Formulas.f5_29(value, e)
    Set sp = Nothing
End Function

Public Function getFormulas5_32(B_ As Double, L_ As Double, Ny As Double, Nq As Double, _
                                Nc As Double, ksi_y As Double, ksi_q As Double, ksi_c As Double, _
                                Y1 As Double, Y1_ As Double, C1 As Double, d As Double)
    Dim sp As New C_SP22_13330_2011
        getFormulas5_32 = sp.Formulas.f5_32(B_, L_, Ny, Nq, Nc, ksi_y, ksi_q, ksi_c, Y1, Y1_, C1, d)
    Set sp = Nothing
End Function

Public Function getFormulas5_33y(L_B As Double)
    Dim sp As New C_SP22_13330_2011
        getFormulas5_33y = sp.Formulas.f5_33y(L_B)
    Set sp = Nothing
End Function

Public Function getFormulas5_33q(L_B As Double)
    Dim sp As New C_SP22_13330_2011
        getFormulas5_33q = sp.Formulas.f5_33q(L_B)
    Set sp = Nothing
End Function

Public Function getFormulas5_33c(L_B As Double)
    Dim sp As New C_SP22_13330_2011
        getFormulas5_33c = sp.Formulas.f5_33c(L_B)
    Set sp = Nothing
End Function

Public Function getFormulas6_37(angle As Double, seismicity As Double)
    Dim sp As New C_SP22_13330_2011
        getFormulas6_37 = sp.Formulas.f6_37(angle, seismicity)
    Set sp = Nothing
End Function

Public Function getFormulas6_38(F1 As Double, ksi_q As Double, ksi_c As Double, _
                                Y1_ As Double, C1 As Double, Fi1 As Double, d As Double)
    Dim sp As New C_SP22_13330_2011
        getFormulas6_38 = sp.Formulas.f6_38(F1, ksi_q, ksi_c, Y1_, C1, Fi1, d)
    Set sp = Nothing
End Function

Public Function getFormulas6_39(F2 As Double, F3 As Double, ksi_y As Double, _
                                Y1 As Double, b As Double, seismicity As Double, p0 As Double)
    Dim sp As New C_SP22_13330_2011
        getFormulas6_39 = sp.Formulas.f6_39(F2, F3, ksi_y, Y1, b, seismicity, p0)
    Set sp = Nothing
End Function

Public Function getFormulas6_41(pb As Double, p0 As Double, b As Double)
    Dim sp As New C_SP22_13330_2011
        getFormulas6_41 = sp.Formulas.f6_41(pb, p0, b)
    Set sp = Nothing
End Function

Public Function getFormulas6_42(pb As Double, p0 As Double, b As Double, L As Double)
    Dim sp As New C_SP22_13330_2011
        getFormulas6_42 = sp.Formulas.f6_42(pb, p0, b, L)
    Set sp = Nothing
End Function

Public Function getFormulas6_43(pb As Double, b As Double, L As Double, ea As Double)
    Dim sp As New C_SP22_13330_2011
        getFormulas6_43 = sp.Formulas.f6_43(pb, b, L, ea)
    Set sp = Nothing
End Function
