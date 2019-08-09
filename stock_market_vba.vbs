Sub Stock_Market()

    ' Declare Variables
    Dim Stock_Total As Double
    Dim Ticker_Name As String
    Dim Table_Row As Integer
    Dim Last_Row As Long
    Dim j As Integer
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    
    j = 1
    
    Do While j <= Worksheets.Count
        
        ' Choose worksheet corresponding to 'j' value
        Worksheets(j).Select
        
        ' Define Last_Row
        Last_Row = Range("A" & Rows.Count).End(xlUp).Row
        
        ' Add Headers
        Range("I1") = "Ticker"
        Range("J1") = "Yearly Change"
        Range("K1") = "Percent Change"
        Range("L1") = "Total Volume"
        Range("I1:L1").Font.Bold = True
        Columns("J").AutoFit
        Columns("K").AutoFit
        Columns("L").AutoFit
        
        ' Set initial Stock_Total to zero
        Stock_Total = 0
        
        ' Set initial location for ticker name and totals in table
        Table_Row = 2
        
        ' Set initial Open_Price
        Open_Price = Cells(2, 3).Value
        
        ' Start loop through all of the rows
        For i = 2 To Last_Row
        
            ' Check if a change in Ticker_Name occurs
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
                ' Store the Ticker_Name
                Ticker_Name = Cells(i, 1).Value
                
                ' Add to the to the ticker's Stock_Total
                Stock_Total = Stock_Total + Cells(i, 7).Value
                
                ' Print Ticker_Name in table
                Range("I" & Table_Row).Value = Ticker_Name
                
                ' Print Stock_Total in table
                Range("L" & Table_Row).Value = Stock_Total
                
                ' Store the closing price
                Close_Price = Cells(i, 6)
                
                ' Calculate Yearly_Change
                Yearly_Change = Close_Price - Open_Price
                
                ' Print Yearly_Change in table
                Range("J" & Table_Row).Value = Yearly_Change
                    
                ' Format if stock price increases
                If Yearly_Change > 0 Then
                    
                    Range("J" & Table_Row).Interior.ColorIndex = 4
                   
                'Format if stock price decreases
                ElseIf Yearly_Change < 0 Then
                    
                    Range("J" & Table_Row).Interior.ColorIndex = 3
                    
                End If
                
                If Close_Price <> 0 Then
                
                    ' Calculate Percent_Change
                    Percent_Change = (1 - Open_Price / Close_Price) * 100
                    
                    'Print Percent_Change in table
                    Range("K" & Table_Row).Value = Percent_Change & "%"
                
                Else
                    Range("K" & Table_Row).Value = "BANKRUPT"
                
                End If
                
                ' Increase Table_Row
                Table_Row = Table_Row + 1
                
                ' Reset Stock_Total
                Stock_Total = 0
                
                ' Set Open_Price
                Open_Price = Cells(i + 1, 3).Value
                
            ' If the TickerName stays then same, do the following
            Else
            
                ' Add to the StockTotal
                Stock_Total = Stock_Total + Cells(i, 7).Value
                
                
            End If
            
        Next i
        
        j = j + 1
    Loop

End Sub

