Sub Analyze_Dataset():

' Run Stock_Analysis() for all sheets in the dataset:

    Dim sheet As Worksheet
    Application.ScreenUpdating = False
    For Each sheet In Worksheets
        sheet.Select
        Call Stock_Analysis
    Next
    Application.ScreenUpdating = True

End Sub


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

' Create variables to store ticker symbol, yearly change, percent change, and stock volume:

Dim ticker_symbol As String
Dim yearly_change As Double
Dim percent_change As Double
Dim stock_volume As Double

' Create variables to store total yearly change, average yearly change, and total stock volume:

Dim total_yearly_change As Double
total_yearly_change = 0

Dim average_yearly_change As Double
average_yearly_change = 0

Dim total_stock_volume As Double
total_stock_volume = 0

' Create a variable to refer to the cells that will display ticker symbol, total yearly change, percent change, and total stock volume:

Dim r As Integer
r = 2

' Create a variable to count the number of rows for each ticker symbol:

Dim c As Long
c = 0

' Determine the last row of the dataset:

Dim last_row As Long
last_row = Cells(Rows.Count, "A").End(xlUp).Row

' Loop through each row:

For i = 2 To last_row

' Initially set yearly_change, percent_change, and stock_volume equal to 0:

    yearly_change = 0
    percent_change = 0
    stock_volume = 0
    
' Retrieve ticker_symbol, yearly_change, and stock_volume:

    ticker_symbol = Cells(i, 1).Value
    yearly_change = (Cells(i, 3).Value) - (Cells(i, 6).Value)
    stock_volume = Cells(i, 7).Value
      c = c + 1
    
' Calculate total_yearly_change and total_stock_volume:
    
    If Cells(i, 1) = Cells(i + 1, 1) Then
    total_yearly_change = total_yearly_change + yearly_change
    total_stock_volume = total_stock_volume + stock_volume
  
    
' Calculate average_yearly_change, percent_change. Enter all values into the designated cells:
    
    Else
    Cells(r, 9).Value = ticker_symbol
    
    average_yearly_change = total_yearly_change / c
    Cells(r, 10).Value = average_yearly_change
    
    percent_change = average_yearly_change * 100 / Cells(i, 3).Value
    Cells(r, 11).Value = percent_change
    
    Cells(r, 12).Value = total_stock_volume
    
' Conditional formatting for cell interior color for yearly change:

    If average_yearly_change < 0 Then
    Cells(r, 10).Interior.ColorIndex = 3
    
    Else
    Cells(r, 10).Interior.ColorIndex = 4
    
    End If
    
' Reset variables to 0:
    
    r = r + 1
    total_yearly_change = 0
    average_yearly_change = 0
    total_stock_volume = 0
    c = 0
    
    End If

    
Next i

End Sub
