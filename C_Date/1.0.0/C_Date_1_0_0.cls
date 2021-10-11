VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Date"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function DateToUnixTimestamp(dt As Date) As Long
    DateToUnixTimestamp = DateDiff("s", "1/1/1970 00:00:00", dt)
End Function

Public Function UnixTimestampToDate(ts As Long) As Date
    UnixTimestampToDate = DateAdd("s", ts, "1/1/1970 00:00:00")
End Function
