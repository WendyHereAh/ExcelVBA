Sub Clear_first_300()
    Range("A4:B302").Select
    Selection.Delete
    MsgBox "Deleted!"
End Sub
