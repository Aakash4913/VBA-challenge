Sub yearly_change()

Dim i As Long

Dim stock1 As String
Dim stock2 As String

Dim open1 As Double
Dim close1 As Double
Dim yearly_change As Double
Dim percent_change As Double

Dim lastrow As Long
Dim j As Long

Dim close_final As Double
Dim yearly_final As Double
Dim percent_final As Double
Dim total_stock_volume As Double
Dim stock_volume As Double

Dim unique_stock_title As String
Dim yearly_title As String
Dim percent_title As String
Dim total_stock_title As String
Dim stock_title1 As String
Dim stock_title2 As String



For Each ws In Worksheets

'setting titles and getting the lastrow
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
yearly_title = "Yearly Change"
percent_title = "Percent Change"
total_stock_title = "Total Stock Volume"
unique_stock_title = "Ticker"
ws.Cells(1, 10).Value = yearly_title
ws.Cells(1, 11).Value = percent_title
ws.Cells(1, 12).Value = total_stock_title

 
 stock_title1 = ws.Cells(2, 1)
 
stock2 = ws.Cells(2, 1).Value
open1 = ws.Cells(2, 3).Value
stock1 = ws.Cells(2, 1).Value
j = 1
total_stock_volume = 0

For i = 2 To lastrow

    stock_volume = ws.Cells(i, 7).Value
    total_stock_volume = total_stock_volume + stock_volume

    
    stock1 = ws.Cells(i, 1).Value
  
  If stock1 <> stock2 Then
   
   stock_title2 = ws.Cells(i, 1).Value
   
   close1 = ws.Cells(i - 1, 6).Value

    yearly_change = close1 - open1
    percent_change = ((close1 - open1) / open1)
    j = j + 1
    
    ws.Cells(j, 10).Value = yearly_change
    ws.Cells(j, 11).Value = percent_change
    ws.Cells(j, 12).Value = (total_stock_volume - stock_volume)
    ws.Cells(j + 1, 9).Value = stock_title2
    
    open1 = ws.Cells(i, 3).Value
    stock2 = ws.Cells(i, 1).Value
    total_stock_volume = 0
    
    ws.Cells(j, 11).NumberFormat = "0.00%"
    
If yearly_change < 0 Then
    ws.Cells(j, 10).Interior.ColorIndex = 3

Else
    ws.Cells(j, 10).Interior.ColorIndex = 4

End If

End If
    
Next i
    
    
    
    
close_final = ws.Cells(i - 1, 6).Value
yearly_final = close_final - open1
percent_final = ((close_final - open1) / open1)
  
  ws.Cells(j + 1, 10) = yearly_final
  ws.Cells(j + 1, 11) = percent_final
  ws.Cells(j + 1, 12) = total_stock_volume
  ws.Cells(2, 9) = stock_title1
  ws.Cells(1, 9) = "Ticker"
  
  ws.Cells(j + 1, 11).NumberFormat = "0.00%"

If yearly_final < 0 Then
    ws.Cells(j + 1, 10).Interior.ColorIndex = 3

Else
    ws.Cells(j + 1, 10).Interior.ColorIndex = 4

End If

Next ws


End Sub
