Attribute VB_Name = "Module1"
 Sub Cell_Looping()
 
 
 Dim ws As Worksheet
    Dim i As Long
    Dim r As Long
    Dim RowCount As Long
    Dim value As String
    
    For Each ws In Worksheets
        
        FirstOpen = 0
        LastClose = 0
        TotalStockVolume = 0
        ' select I1
        ' set the value to "Ticker"
        ws.Range("I1").value = "Ticker"
        ' select J1
        ' set the value to "Yearly Change"
        ws.Range("J1").value = "Yearly Change"
        ' select K1
        ' set the value to "Percent Change"
        ws.Range("K1").value = "Percent Change"
        ' select L1
        ' set the value to "Total Stock Volume"
        ws.Range("L1").value = "Total Stock Volume"
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row
        i = 2
        ' for debugging
        RowCount = 92500
        For r = 2 To RowCount
            ' select G2
            ' ws.Cells(r, 7)
            ' add G2 to TotalStockVolume
            TotalStockVolume = ws.Cells(r, 7).value + TotalStockVolume
            ' select A263 and compare to A264
            ' ws.Cells(r, 1) and ws.Cells(r + 1, 1)
            If ws.Cells(r, 1).value <> ws.Cells(r + 1, 1).value And FirstOpen <> 0 Then
                ' select F263 (r, 5)
                ' ws.Cells(r, 5)
                ' set the value for LastClose
                LastClose = ws.Cells(r, 5).value
                ' select J2
                ' i - 1 is the row
                ' 10 is the column J
                ' ws.Cells(i - 1, 10)
                ' put LastClose - FirstOpen into J2
                ws.Cells(i - 1, 10).value = LastClose - FirstOpen
                ' select K2
                ' ws.Cells(i - 1, 11)
                ' put (LastClose - FirstOpen) / FirstOpen into K2
                ws.Cells(i - 1, 11).value = (LastClose - FirstOpen) / FirstOpen
                ' select L2
                ' ws.Cells(i - 1, 12)
                ' put TotalStockVolume into L2
                ws.Cells(i - 1, 12).value = TotalStockVolume
                ' reset TotalStockVolume to 0
                TotalStockVolume = 0
            End If
            ' next row A3
            ' see if the value changed (compare to A2)
            If value = ws.Cells(r, 1).value Then
                ' no change
                ' repeat approx 200 times
                ' ...
            Else
                ' see if the value changed (compare A263 to A262)
                ' value changed
                ' copy the value at A263
                ' select I3
                ' paste the value from A263
                ' starting at A2
                ' ws.Cells(r, 1)
                ' copy the value at A2
                value = ws.Cells(r, 1).value
                ' select I2
                ' paste the value from A2
                ' ws.Cells(i, 9)
                ws.Cells(i, 9).value = value
                ' select C2
                ' ws.Cells(r, 3)
                ' set the value for FirstOpen
                FirstOpen = ws.Cells(r, 3).value
                ' go down one cell
                i = i + 1
            End If
        Next r
    Next ws
            
            
            
            
           End Sub





 
