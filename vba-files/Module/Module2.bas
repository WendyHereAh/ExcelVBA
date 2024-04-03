Attribute VB_Name = "Module2"
Sub Clear()
    Range("A3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    ActiveSheet.ListObjects("Table1").Resize Range("A1:B3")
    Range("A3").Select
End Sub


