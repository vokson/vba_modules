VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_SP22_13330_2011_Tables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Private math As C_Math

' Private Sub Class_Initialize()
'     Set math = New C_Math
' End Sub

' Private Sub Class_Terminate()
'     Set math = Nothing
' End Sub

' “‡·ÎËˆ‡ 5.10
' C_Soil v.2.0.0
Function t5_10(soil As C_Soil) As Double

    Dim result As Double

    Select Case soil.TypeBySize
        Case " –”œÕŒŒ¡ÀŒÃŒ◊Õ€…":
            result = 0.27

        Case "œ≈—Œ ", "—”œ≈—‹":
            result = (0.3 + 0.35) / 2

        Case "—”√À»ÕŒ ":
            result = (0.37 + 0.35) / 2

        Case "√À»Õ¿":
            If soil.LiquidityIndex <= 0 Then result = (0.2 + 0.3) / 2

            If soil.LiquidityIndex > 0 And soil.LiquidityIndex <= 0.25 Then _
               result = 0.3 + (0.38 - 0.3) * soil.LiquidityIndex / 0.25

            If soil.LiquidityIndex > 0.25 And soil.LiquidityIndex <= 1# Then _
               result = 0.38 + (0.45 - 0.38) * (soil.LiquidityIndex - 0.25) / 0.75

            If soil.LiquidityIndex > 1# Then result = 0.45
    End Select

    t5_10 = result

End Function