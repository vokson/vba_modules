VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_NIIK_TCM_Transmittal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function make(ByRef wordDoc As Word.Document, dic As Dictionary) As String
    
    wordDoc.Tables(1).Columns(4).Cells.item(1).Range.text = dic.item("TRANSMITTAL")
    wordDoc.Tables(1).Columns(4).Cells.item(2).Range.text = dic.item("DATE")
    wordDoc.Tables(1).Columns(4).Cells.item(3).Range.text = dic.item("PURPOSE")
    wordDoc.Tables(1).Columns(4).Cells.item(4).Range.text = dic.item("SUMMARY")
    wordDoc.Tables(1).Columns(4).Cells.item(5).Range.text = dic.item("IS_REPLY_REQUESTED")
    wordDoc.Tables(1).Columns(4).Cells.item(6).Range.text = dic.item("REPLY_BY")
    wordDoc.Tables(1).Columns(4).Cells.item(7).Range.text = dic.item("COUNT")

    Dim docs
    docs = dic.Item("DOCS")

    Dim i As Integer
    For i = LBound(docs) To UBound(docs)

        if wordDoc.Tables(2).Rows.Count < (i+2) Then wordDoc.Tables(2).Rows.Add

        wordDoc.Tables(2).Rows(i + 2).Cells.item(1).Range.text = docs(i).item("NIPIGAZ_CODE")
        wordDoc.Tables(2).Rows(i + 2).Cells.item(2).Range.text = docs(i).item("REVISION")
        wordDoc.Tables(2).Rows(i + 2).Cells.item(3).Range.text = docs(i).item("TCM_CODE")
        wordDoc.Tables(2).Rows(i + 2).Cells.item(4).Range.text = docs(i).item("CLASS")
        wordDoc.Tables(2).Rows(i + 2).Cells.item(5).Range.text = docs(i).item("REASON")
        wordDoc.Tables(2).Rows(i + 2).Cells.item(6).Range.text = docs(i).item("ENGLISH_TITLE")
        wordDoc.Tables(2).Rows(i + 2).Cells.item(7).Range.text = docs(i).item("RUSSIAN_TITLE")

    Next i

End Function