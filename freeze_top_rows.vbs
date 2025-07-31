Sub FreezeTopRowAllSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        ws.Rows("2:2").Select
        ActiveWindow.FreezePanes = True
    Next ws
End Sub
