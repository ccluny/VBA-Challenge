Attribute VB_Name = "Module1"
Sub solution():
    ' MsgBox "test"
    Dim RowCount As Integer
    Dim ws As Worksheet
    Dim total As Double
    Dim i As Integer
    Dim j As Integer

    For Each ws In Worksheets
        ' MsgBox (ws.Range("A1").Value)
        ws.Range("I1").Value = "Ticker"
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row
        ws.Range("I2").Value = RowCount
    
        For i = 2 To 3
            MsgBox (ws.Cells(i + 1, 1).Value)
        Next i
        
        MsgBox (ws.Cells(2, 7).Value)
    Next ws
    
    
    
End Sub
