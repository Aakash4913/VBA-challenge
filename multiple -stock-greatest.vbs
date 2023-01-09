Sub greatest()


'
Dim greatest_increase_title As String
Dim greatest_decrease_title As String
Dim greatest_volume_title As String

Dim greatest_increase_stock As String
Dim greatest_decrease_stock As String
Dim greatest_volume_stock As String

Dim percent_change1 As Double
Dim percent_change2 As Double
Dim stock_volume1 As Double


Dim greatest_percent_increase As Double
Dim greatest_percent_decrease As Double
Dim greatest_stock_volume As Double


For Each ws In Worksheets



lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
greatest_increase_title = "Greatest Percent Increase"
greatest_decrease_title = "Greatest Percent Decrease"
greatest_volume_title = "Greatest Total Volume"

greatest_percent_increase = ws.Cells(2, 11).Value
greatest_percent_decrease = ws.Cells(2, 11).Value
greatest_stock_volume = ws.Cells(2, 12).Value



For i = 2 To lastrow
    
    percent_change1 = ws.Cells(i, 11).Value
    percent_change2 = ws.Cells(i, 11).Value
    stock_volume1 = ws.Cells(i, 12).Value
    
    
        If percent_change1 >= greatest_percent_increase Then
            greatest_percent_increase = ws.Cells(i, 11).Value
            greatest_increase_stock = ws.Cells(i, 9).Value
            
        End If
        
        If percent_change2 <= greatest_percent_decrease Then
            greatest_percent_decrease = ws.Cells(i, 11).Value
            greatest_decrease_stock = ws.Cells(i, 9).Value
        
        End If
        
        If stock_volume1 >= greatest_stock_volume Then
            greatest_stock_volume = ws.Cells(i, 12).Value
            greatest_volume_stock = ws.Cells(i, 9).Value
            
        End If
        
        
    Next i

ws.Cells(4, 16).Value = greatest_percent_increase
ws.Cells(5, 16).Value = greatest_percent_decrease
ws.Cells(6, 16).Value = greatest_stock_volume


ws.Cells(4, 15).Value = greatest_increase_stock
ws.Cells(5, 15).Value = greatest_decrease_stock
ws.Cells(6, 15).Value = greatest_volume_stock

ws.Cells(4, 14).Value = greatest_increase_title
ws.Cells(5, 14).Value = greatest_decrease_title
ws.Cells(6, 14).Value = greatest_volume_title

ws.Cells(4, 16).NumberFormat = "0.00%"
ws.Cells(5, 16).NumberFormat = "0.00%"

Next ws

End Sub
