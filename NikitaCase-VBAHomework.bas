Sub StockSummaries():

Dim ws As Worksheet

Dim ticker As String
Dim ticker_vol As Double

Dim sum_row As Double
Dim last_row As Double
Dim last_sum_row As Double
  
Dim stock_open As Double
Dim stock_close As Double

Dim yearly_change As Double
Dim percentchange As Double

Dim gr_ticker As String
Dim greatest As Double
Dim l_ticker As String
Dim least As Double
Dim gr_vol As String
Dim greatest_total_volume As Double

'''''''''''''''''''''''''''''''''''''''''''''''''


For Each ws In Worksheets
ws.Activate


    sum_row = 2
    last_row = Range("A" & Rows.Count).End(xlUp).Row
    ticker_vol = 0
    stock_open = Cells(2, 3).Value


    For r = 2 To last_row
        
        If Cells(r + 1, 1).Value <> Cells(r, 1).Value Then
            
            ' assign value to variables
            ticker = Cells(r, 1).Value
            ticker_vol = ticker_vol + Cells(r, 7).Value
            stock_close = Cells(r, 6).Value
            yearly_change = stock_close - stock_open
            
            
            ' stop diving by 0 so the code can work
            If stock_open = 0 Then
            Else
                percentchange = (yearly_change / stock_open)
            End If
        
            ' print variables
            Range("J" & sum_row).Value = ticker
            Range("K" & sum_row).Value = yearly_change
            Range("L" & sum_row).Value = FormatPercent(percentchange, [2])
            Range("M" & sum_row).Value = ticker_vol
            
            
            ' apply conditional formatting to yearly change cell
            If Range("K" & sum_row).Value > 0 Then
                Range("K" & sum_row).Interior.ColorIndex = 4
            Else
                Range("K" & sum_row).Interior.ColorIndex = 3
            End If
                            

            ' reset and increment
            sum_row = sum_row + 1
            ticker_vol = 0
            stock_open = Cells(r + 1, 3).Value
                
        Else
            ' add stock volume until stock changes
            ticker_vol = ticker_vol + Cells(r, 7).Value
        
        End If
    Next r
    
    ' print table headings
    
    Range("J1").Value = "Ticker"
    Range("K1").Value = "Yearly Change"
    Range("L1").Value = "Percent Change"
    Range("M1").Value = "Total Stock Volume"
            

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    ' find max and min percent change
        
    last_sum_row = Range("J" & Rows.Count).End(xlUp).Row
    
    greatest = 0
    least = 0
    
    For n = 2 To last_sum_row
        If Range("L" & n).Value > greatest Then
            greatest = Range("L" & n).Value
            gr_ticker = Range("J" & n).Value
        
        ElseIf Range("L" & n).Value < least Then
            least = Range("L" & n).Value
            l_ticker = Range("J" & n).Value
        End If
    Next n
    
    
    ' find highest ticker volume
    
    greatest_total_volume = 0
    
    For v = 2 To last_sum_row
        If Range("M" & v).Value > greatest_total_volume Then
            greatest_total_volume = Range("M" & v).Value
            gr_vol = Range("J" & v).Value
    
        End If
    Next v
    
        
    ' print summary table
    
    
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    Range("O2").Value = "Greatest % Increase"
    Range("P2").Value = gr_ticker
    Range("Q2").Value = FormatPercent(greatest, [2])
    
    Range("O3").Value = "Greatest % Decrease"
    Range("P3").Value = l_ticker
    Range("Q3").Value = FormatPercent(least, [2])
    
    Range("O4").Value = "Greatest Total Volume"
    Range("P4").Value = gr_vol
    Range("Q4").Value = greatest_total_volume
    
Next ws


End Sub