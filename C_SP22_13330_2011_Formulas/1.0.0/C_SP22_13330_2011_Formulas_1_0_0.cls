VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_SP22_13330_2011_Formulas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private math As C_Math

Private Sub Class_Initialize()
Set math = New C_Math
End Sub

Private Sub Class_Terminate()
    Set math = Nothing
End Sub

' СП 22.13330.2011  - Формула 5.7
Function f5_7(Yc1 As Double, Yc2 As Double, k As Double, My As Double, _
                            Mq As Double, Mc As Double, b As Double, Y2 As Double, _
                            Y2_ As Double, C2 As Double, d1 As Double, db As Double) As Double
    
    Dim kz As Double

    kz = 8 / b + 0.2
        If b < 10 Then kz = 1
        If db > 2 Then db = 2
    f5_7 = Yc1 * Yc2 / k * (My * kz * b * Y2 + Mq * d1 * Y2_ + (Mq - 1) * db * Y2_ + Mc * C2)

End Function

' СП 22.13330.2011 - Формула 5.8
Function f5_8(hs As Double, hcf As Double, Ycf As Double, Y2_ As Double)
  f5_8 = hs + hcf * Ycf / Y2_
End Function

' СП 22.13330.2011 - Формула 5.24
Function f5_24(d As Double, ke As Double, N As Double, e As Double, a As Double)
  f5_24 = d * ke * N * e / (a / 2) ^ 3
End Function

' СП 22.13330.2011 - Формула 5.25
Function f5_25(nu As Double, e As Double)
  f5_25 = (1 - nu ^ 2) / e
End Function

' СП 22.13330.2011 - Формула 5.29
Function f5_29(value As Double, e As Double)
  f5_29 = value - 2 * e
End Function

' СП 22.13330.2011 - Формула 5.32
Function f5_32(B_ As Double, L_ As Double, Ny As Double, Nq As Double, _
                                Nc As Double, ksi_y As Double, ksi_q As Double, ksi_c As Double, _
                                Y1 As Double, Y1_ As Double, C1 As Double, d As Double)
  f5_32 = B_ * L_ * (Ny * ksi_y * B_ * Y1 + Nq * ksi_q * Y1_ * d + Nc * ksi_c * C1)
End Function

' СП 22.13330.2011 - Формула 5.33 ksi_y
Function f5_33y(L_B As Double)
  If L_B < 1 Then L_B = 1
  f5_33y = 1 - 0.25 / L_B
End Function

' СП 22.13330.2011 - Формула 5.33 ksi_q
Function f5_33q(L_B As Double)
  If L_B < 1 Then L_B = 1
  f5_33q = 1 + 1.5 / L_B
End Function

' СП 22.13330.2011 - Формула 5.33 ksi_c
Function f5_33c(L_B As Double)
  If L_B < 1 Then L_B = 1
  f5_33c = 1 + 0.3 / L_B
End Function

' СП 22.13330.2011 - Формула 6.37
Function f6_37(angle As Double, seismicity As Double)
  Dim delta As Double
  delta = 0
  Select Case seismicity
    Case "7":
      delta = 2
    Case "8":
      delta = 4
    Case "9":
      delta = 7
  End Select
  f6_37 = angle - delta
End Function

' СП 22.13330.2011 - Формула 6.38
Function f6_38(F1 As Double, ksi_q As Double, ksi_c As Double, _
              Y1_ As Double, C1 As Double, Fi1 As Double, d As Double)
    Dim GradToRad As Double
    GradToRad = Fi1 * math.pi / 180
    f6_38 = ksi_q * F1 * Y1_ * d + ksi_c * (F1 - 1) * C1 / Tan(GradToRad)
End Function

' СП 22.13330.2011 - Формула 6.39
Function f6_39(F2 As Double, F3 As Double, ksi_y As Double, _
                Y1 As Double, b As Double, seismicity As Double, p0 As Double)
    Dim k_eq As Double
    k_eq = 0
        Select Case seismicity
            Case "7":
                k_eq = 0.1
            Case "8":
                k_eq = 0.2
            Case "9":
                k_eq = 0.4
    End Select
        
    If F2 < k_eq * F3 Then
        f6_39 = p0
    Else
        f6_39 = p0 + ksi_y * Y1 * b * (F2 - k_eq * F3)
    End If
End Function

' СП 22.13330.2011 - Формула 6.41
Function f6_41(pb As Double, p0 As Double, b As Double)
    f6_41 = b * (pb - p0) / (6 * (pb + p0))
End Function

' СП 22.13330.2011 - Формула 6.42
Function f6_42(pb As Double, p0 As Double, b As Double, L As Double)
    f6_42 = 0.5 * b * L * (pb + p0)
End Function

' СП 22.13330.2011 - Формула 6.43
Function f6_43(pb As Double, b As Double, L As Double, ea As Double)
    f6_43 = b * L * pb / (1 + 6 * ea / b)
End Function

