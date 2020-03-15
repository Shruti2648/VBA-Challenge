Sub Stock_Analysis():

' Add columns to display ticker symbol, yearly change, percent change, and volume:

Cells(1, 9).EntireColumn.Insert
Cells(1, 9).Value = "Ticker Symbol"
    
Cells(1, 10).EntireColumn.Insert
Cells(1, 10).Value = "Yearly Change"
    
Cells(1, 11).EntireColumn.Insert
Cells(1, 11).Value = "Percent Change"
    
Cells(1, 12).EntireColumn.Insert
Cells(1, 12).Value = "Total Stock Volume"

' Create variables to store ticker symbol, yearly change, percent change, and volume:

Dim ticker_symbol As String
Dim yearly_change As Double
Dim percent_change As Double
Dim stock_volume As Double

' Loop through each row:

For i = 2 To 100

' Initially set yearly_change, percent_change, and stock_volume equal to 0:

    yearly_change = 0
    percent_change = 0
    stock_volume = 0
    
' Retrieve ticker_symbol and enter in cell:

    ticker_symbol = Cells(i, 1).Value
    Cells(i, 9).Value = ticker_symbol
        
' Calculate yearly_change and enter in cell:

    yearly_change = (Cells(i, 3).Value) - (Cells(i, 6).Value)
    Cells(i, 10).Value = yearly_change
        
' Calculate percent_change and enter in cell:

    percent_change = yearly_change * 100
    Cells(i, 11).Value = percent_change
    
' Calculate stock_volume and enter in cell:

    stock_volume = Cells(i, 1).Value
    Cells(i, 12).Value = stock_volume
    

    
    
Next i

End Sub
