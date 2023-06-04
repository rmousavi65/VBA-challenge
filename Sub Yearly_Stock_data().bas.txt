Sub Yearly_Stock_data()

'Declare a worksheet
Dim worksheetCount As Integer

'Set variables and rows for worksheet

Dim I As Integer
Dim j As Long
Dim num_Ticker As Integer
Dim yearly_Change As Double
Dim percentage As Double
Dim opening_Price As Double
Dim closing_Price As Double
Dim stock_Volume As Double
         
'Activeworkbook

worksheetCount = ActiveWorkbook.Worksheets.Count
    
    For I = 1 To worksheetCount
    ActiveWorkbook.Worksheets(I).Activate
    ticker_row = 2
    yearly_Change = 0
    percentage = 0
    opening_Price = 0
    num_Ticker = 0
    stock_Volume = 0
 
 'Create column headings
 
    Cells(1, "I").Value = "Ticker"
    Cells(1, "J").Value = "Yearly Change"
    Cells(1, "K").Value = "Percentage Change"
    Cells(1, "L").Value = "Stock Volume"
                
    'Loop for lastrow
    
    For j = 2 To ActiveWorkbook.Worksheets(I).UsedRange.Rows.Count
      
      If opening_Price = 0 Then
         opening_Price = Cells(j, 3).Value
      End If
                        
      stock_Volume = stock_Volume + Cells(j, 7).Value
                
       If Cells(j + 1, 1).Value <> Cells(j, 1).Value Then
           Ticker = Cells(j, 1).Value
           Cells(ticker_row, "I").Value = Ticker
           num_Ticker = num_Ticker + 1
           Cells(num_Ticker + 1, 9) = Cells(j, 1).Value
           closing_Price = Cells(j, 6)
           yearly_Change = closing_Price - opening_Price
           Cells(num_Ticker + 1, 10).Value = yearly_Change
                        
       If yearly_Change > 0 Then
          Cells(num_Ticker + 1, 10).Interior.ColorIndex = 4
       ElseIf yearly_Change < 0 Then
          Cells(num_Ticker + 1, 10).Interior.ColorIndex = 3
       Else
          Cells(num_Ticker + 1, 10).Interior.ColorIndex = 6
       End If
                        
       If opening_Price = 0 Then
             percentage = 0
       Else
          percentage = (yearly_Change / opening_Price) * 100
          
       End If
         Cells(num_Ticker + 1, 11).Value = Format(percentage, "Percent")
         Cells(num_Ticker + 1, 12).Value = stock_Volume
        End If
                                
        If percentage > 0 Then
         Cells(num_Ticker + 1, 11).Interior.ColorIndex = 4
        ElseIf percentage < 0 Then
         Cells(num_Ticker + 1, 11).Interior.ColorIndex = 3
        Else
         Cells(num_Ticker + 1, 11).Interior.ColorIndex = 6
        End If


                Next j


         Next I
End Sub
