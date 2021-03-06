VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_SP50_102_2003_Tables"
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

' ������� �.1
' isDrivenPile - (true - �������� ����, false - ��������, �������, ����-��������, ����-�����) As Boolean
' soil As C_Soil - ������ ������
Public Function E1(isDrivenPile As Boolean, soil As C_Soil) As Double
  Dim values As New Dictionary

  values.Item("CLAY_IL") = Array(0, 0.5, 0.75, 1.0)
  values.Item("CLAY_DRIVEN") = Array(8000, 5000, 2500, 650)
  values.Item("CLAY_NONDRIVEN") = Array(6000, 4000, 2000, 500)

  values.Item("CLAY_SANDY_IL") = Array(0, 1.0)
  values.Item("CLAY_SANDY_DRIVEN") = Array(5000, 2500)
  values.Item("CLAY_SANDY_NONDRIVEN") = Array(4000, 2000)

  values.Item("SAND_FINE_e") = Array(0.6, 0.8)
  values.Item("SAND_FINE_DRIVEN") = Array(5000, 2500)
  values.Item("SAND_FINE_NONDRIVEN") = Array(4000, 2000)

  values.Item("SAND_SMALL_e") = Array(0.6, 0.75)
  values.Item("SAND_SMALL_DRIVEN") = Array(8000, 5000)
  values.Item("SAND_SMALL_NONDRIVEN") = Array(6000, 4000)

  values.Item("SAND_MIDDLE_e") = Array(0.55, 0.70)
  values.Item("SAND_MIDDLE_DRIVEN") = values.Item("SAND_SMALL_DRIVEN")
  values.Item("SAND_MIDDLE_NONDRIVEN") = values.Item("SAND_SMALL_NONDRIVEN")

  values.Item("SAND_COARSE_e") = Array(0.55, 0.70)
  values.Item("SAND_COARSE_DRIVEN") = Array(13000, 8000)
  values.Item("SAND_COARSE_NONDRIVEN") = Array(10000, 6000)

  values.Item("SAND_GRAVELOUS_e") = Array(0.55, 0.70)
  values.Item("SAND_GRAVELOUS_DRIVEN") = Array(0, 0)
  values.Item("SAND_GRAVELOUS_NONDRIVEN") = Array(20000, 10000)

  values.Item("MACROFRAGMENTAL_GRAVEL_DRIVEN") = 0
  values.Item("MACROFRAGMENTAL_GRAVEL_NONDRIVEN") = 20000
  values.Item("MACROFRAGMENTAL_COBBLES_DRIVEN") = 0
  values.Item("MACROFRAGMENTAL_COBBLES_NONDRIVEN") = 20000

  E1 = 0#

  Dim drivenText As String
  Dim soilText As String

  if isDrivenPile Then drivenText = "_DRIVEN" Else drivenText = "_NONDRIVEN"
  soilText = ""

  
  Select Case soil.TypeBySize
      Case "����������������":

        soilText = "MACROFRAGMENTAL_"

        Select Case soil.SubtypeBySize
          Case "��������": soilText = soilText & "COBBLES"
          Case "���������": soilText = soilText & "GRAVEL"
        End Select

        E1 = values.Item(soilText & drivenText)
        
      Case "�����":

        soilText = "SAND_"

        Select Case soil.SubtypeBySize
          Case "���������": soilText = soilText & "FINE"
          Case "������": soilText = soilText & "SMALL"
          Case "������� ���������": soilText = soilText & "MIDDLE"
          Case "�������": soilText = soilText & "COARSE"
          Case "�����������": soilText = soilText & "GRAVELOUS"
        End Select

        E1 = CDbl(math.interpolateOneDimensionalArray( _ 
                soil.VoidRatio, _
                values.Item(soilText & "_e"), _
                values.Item(soilText & drivenText)))

        If soil.TypeByDensity = "�������" Then E1 = E1 * 1.3

      Case "�����", "��������", "������":
        soilText = "CLAY"
        if soil.TypeBySize = "������" Then soilText = "CLAY_SANDY"

        E1 = CDbl(math.interpolateOneDimensionalArray( _ 
                soil.LiquidityIndex, _
                values.Item(soilText & "_IL"), _
                values.Item(soilText & drivenText)))

  End Select
  
End Function

' ������� �.2
' pileCase As Integer
'    0 - �������� ���� �� ���������� �����
'    1 - �������� ���� �� �������� �����
'    2 - ������� ���� � �������� �����
' length As Double
' parameter As String - "A0", "B0", "C0" (����.)
Public Function E2(pileCase As Integer, length As Double, parameter As String) As Double
  Dim values As New Dictionary

  If Not (pileCase = 0 Or pileCase = 1 Or pileCase = 2) Then
    MsgBox "SP50.102.2003 Table E.2 - pileCase - ERROR"
    Exit Function
  End If

  If Not (parameter = "A0" Or parameter = "B0" Or parameter = "C0") Then
    MsgBox "SP50.102.2003 Table E.2 - parameter - ERROR"
    Exit Function
  End If

  values.Item("length") = Array(0.5, 0.6, 0.7, 0.8, 0.9, 1.0, 1.1, 1.2, 1.3, 1.4, 1.5, 1.6, 1.7, 1.8, 1.9, 2.0, 2.2, 2.4, 2.6, 2.8, 3.0, 3.5, 4.0)
  values.Item("case0_A0") = Array(72.004, 50.007, 36.745, 28.14, 22.244, 18.03, 14.916, 12.552, 10.717, 9.266, 8.101, 7.154, 6.375, 5.73, 5.19, 4.737, 4.032, 3.526, 3.163, 2.905, 2.727, 2.502, 2.441)
  values.Item("case0_B0") = Array(192.026, 111.149, 70.023, 46.943, 33.008, 24.106, 18.16, 14.041, 11.103, 8.954, 7.349, 6.129, 5.189, 4.456, 3.878, 3.418, 2.756, 2.327, 2.048, 1.869, 1.758, 1.641, 1.621)
  values.Item("case0_C0") = Array(576.243,278.069,150.278,88.279,55.307,36.486,25.123,17.944,13.235,10.05,7.838,6.268,5.133,4.299,3.679,3.213,2.591,2.227,2.013,1.889,1.818,1.757,1.751)
  values.Item("case1_A0") = Array(48.006,33.344,24.507,18.775,14.851,12.049,9.983,8.418,7.208,6.257,5.498,4.887,4.391,3.985,3.653,3.381,2.977,2.713,2.548,2.453,2.406,2.394,2.419)
  values.Item("case1_B0") = Array(96.037,55.609,35.059,23.533,16.582,12.149,9.196,7.159,5.713,4.664,3.889,3.308,2.868,2.533,2.277,2.081,1.819,1.673,1.6,1.572,1.568,1.597,1.618)
  values.Item("case1_C0") = Array(192.291,92.942,50.387,29.763,18.814,12.582,8.836,6.485,4.957,3.937,3.24,2.758,2.419,2.181,2.012,1.894,1.758,1.701,1.687,1.693,1.707,1.739,1.75)
  values.Item("case2_A0") = Array(0.042,0.072,0.114,0.17,0.241,0.329,0.434,0.556,0.695,0.849,1.014,1.186,1.361,1.532,1.693,1.841,2.08,2.24,2.33,2.371,2.385,2.389,2.401)
  values.Item("case2_B0") = Array(0.125,0.18,0.244,0.319,0.402,0.494,0.593,0.698,0.807,0.918,1.028,1.134,1.232,1.321,1.397,1.46,1.545,1.586,1.596,1.593,1.586,1.584,1.6)
  values.Item("case2_C0") = Array(0.5,0.6,0.699,0.798,0.896,0.992,1.086,1.176,1.262,1.342,1.415,1.48,1.535,1.581,1.617,1.644,1.675,1.685,1.687,1.687,1.691,1.711,1.732)

  E2 = CDbl(math.interpolateOneDimensionalArray(length, values.Item("length"), _
     values.Item("case" & CStr(pileCase) & "_" & parameter)))
    
End Function

' ������� �.3
' z As Double
' parameter As String - "A1", "B1", "C1", "D1", "A3", "B3", "C3", "D3", "A4", "B4", "C4", "D4" (����.)
Public Function E3(z As Double, parameter As String) As Double
  Dim values As New Dictionary

  If  parameter <> "A1" And parameter <> "B1" And parameter <> "C1" And parameter <> "D1" And _
      parameter <> "A3" And parameter <> "B3" And parameter <> "C3" And parameter <> "D3" And _
      parameter <> "A4" And parameter <> "B4" And parameter <> "C4" And parameter <> "D4" Then

    MsgBox "SP50.102.2003 Table E.3 - parameter - ERROR"
    Exit Function
  End If

  values.Item("z") = Array(0.0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1.0, 1.1, 1.2, 1.3, 1.4, 1.5, 1.6, 1.7, 1.8, 1.9, 2.0, 2.2, 2.4, 2.6, 2.8, 3.0, 3.5, 4.0)
  values.Item("A1") = Array(1.000, 1.000, 1.000, 1.000, 1.000, 1.000, 0.999, 0.999, 0.997, 0.995, 0.992, 0.987, 0.979, 0.969, 0.955, 0.937, 0.913, 0.882, 0.843, 0.795, 0.735, 0.575, 0.347, 0.033, -0.385, -0.928, -2.928, -5.853)
  values.Item("B1") = Array(0.000, 0.100, 0.200, 0.300, 0.400, 0.500, 0.600, 0.700, 0.799, 0.899, 0.997, 1.095, 1.192, 1.287, 1.379, 1.468, 1.553, 1.633, 1.706, 1.770, 1.823, 1.887, 1.874, 1.755, 1.490, 1.037, -1.272, -5.941)
  values.Item("C1") = Array(0.000, 0.005, 0.020, 0.045, 0.080, 0.125, 0.180, 0.245, 0.320, 0.405, 0.499, 0.604, 0.718, 0.841, 0.974, 1.115, 1.264, 1.421, 1.584, 1.752, 1.924, 2.272, 2.609, 2.907, 3.128, 3.225, 2.463, -0.927)
  values.Item("D1") = Array(0.000, 0.000, 0.001, 0.005, 0.011, 0.021, 0.036, 0.057, 0.085, 0.121, 0.167, 0.222, 0.288, 0.365, 0.456, 0.560, 0.678, 0.812, 0.961, 1.126, 1.308, 1.720, 2.195, 2.724, 3.288, 3.858, 4.980, 4.548)
  values.Item("A3") = Array(0.000, 0.000, -0.001, -0.005, -0.011, -0.021, -0.036, -0.057, -0.085, -0.121, -0.167, -0.222, -0.287, -0.365, -0.455, -0.559, -0.676, -0.808, -0.956, -1.118, -1.295, -1.693, -2.141, -2.621, -3.103, -3.541, -3.919, -1.614)
  values.Item("B3") = Array(0.000, 0.000, 0.000, -0.001, -0.002, -0.005, -0.011, -0.020, -0.034, -0.055, -0.083, -0.122, -0.173, -0.238, -0.319, -0.420, -0.543, -0.691, -0.867, -1.074, -1.314, -1.906, -2.663, -3.600, -4.718, -6.000, -9.544, -11.731)
  values.Item("C3") = Array(1.000, 1.000, 1.000, 1.000, 1.000, 0.999, 0.998, 0.996, 0.992, 0.985, 0.975, 0.960, 0.938, 0.907, 0.866, 0.811, 0.739, 0.646, 0.530, 0.385, 0.207, -0.271, -0.949, -1.877, -3.108, -4.688, -10.340, -17.919)
  values.Item("D3") = Array(0.000, 0.100, 0.200, 0.300, 0.400, 0.500, 0.600, 0.699, 0.799, 0.897, 0.994, 1.090, 1.183, 1.273, 1.358, 1.437, 1.507, 1.566, 1.612, 1.640, 1.646, 1.575, 1.352, 0.917, 0.197, -0.891, -5.854, -15.076)
  values.Item("A4") = Array(0.000, -0.005, -0.020, -0.045, -0.080, -0.125, -0.180, -0.245, -0.320, -0.404, -0.499, -0.603, -0.716, -0.838, -0.967, -1.105, -1.248, -1.396, -1.547, -1.699, -1.848, -2.125, -2.339, -2.437, -2.346, -1.969, 1.074, 9.244)
  values.Item("B4") = Array(0.000, 0.000, -0.003, -0.009, -0.021, -0.042, -0.072, -0.114, -0.171, -0.243, -0.333, -0.443, -0.575, -0.730, -0.910, -1.116, -1.350, -1.613, -1.906, -2.227, -2.578, -3.360, -4.228, -5.140, -6.023, -6.765, -6.789, -0.358)
  values.Item("C4") = Array(0.000, 0.000, 0.000, -0.001, -0.003, -0.008, -0.016, -0.030, -0.051, -0.082, -0.125, -0.183, -0.259, -0.356, -0.479, -0.630, -0.815, -1.036, -1.299, -1.608, -1.966, -2.849, -3.973, -5.355, -6.990, -8.840, -13.692, -15.611)
  values.Item("D4") = Array(1.000, 1.000, 1.000, 1.000, 1.000, 0.999, 0.997, 0.994, 0.989, 0.980, 0.967, 0.946, 0.917, 0.876, 0.821, 0.747, 0.652, 0.529, 0.374, 0.181, -0.057, -0.692, -1.592, -2.821, -4.445, -6.520, -13.826, -23.140)

  E3 = CDbl(math.interpolateOneDimensionalArray(z, values.Item("z"), values.Item(parameter)))
    
End Function

