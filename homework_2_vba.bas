Attribute VB_Name = "Module1"
Sub stock_count()
    
    Dim ws As Worksheet

    For Each ws In Worksheets
        ws.Activate
       
        'assign headers
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
        'Assign bonus headers
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        
        Cells.Columns("O").AutoFit
        
        'Declare variables
        
        Dim ticker As String
        Dim total_volume As Double
        Dim percent_change As Double
        Dim yearly_change As Double
        
        Dim i As Long
        Dim last_row As Long
        Dim unique_ticker_index As Long
        Dim open_index As Long
        Dim open_index2 As Long
        
        'declare bonus variables
        Dim greatest_increase As Double
        Dim greatest_decrease As Double
        Dim greastest_volume As Long
        Dim greatest_increase_row As Double
        Dim greatest_decrease_row As Double
        Dim greatest_volume_row As Double
        
        'assign variables
        
        total_volume = 0
        unique_ticker_index = 2
        last_row = Cells(Rows.Count, 1).End(xlUp).row
        open_index = 2
        
        
        'cycle through rows
        For i = 2 To last_row
        
        'if the ticker matches, then calculate total volume and move on to next i row
        'if ticker does not match, find yearly change, percent change, ticker name, and populate the columns
            
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            
                    ticker = Cells(i, 1).Value
                    
                    total_volume = total_volume + Cells(i, 7).Value
                    
                
                    If total_volume = 0 Then
                        Cells(unique_ticker_index, 9).Value = Cells(i, 1).Value
                        Cells(unique_ticker_index, 10).Value = 0
                        Cells(unique_ticker_index, 11).Value = 0
                        Cells(unique_ticker_index, 12).Value = 0
                    
                    'if total volume is not 0, you will calculate all and populate i-l
                    Else
                        
                        If Cells(open_index, 3).Value = 0 Then
                        
                            'loop through open_index to find the first non-0 open price for ticker
                            For open_index2 = open_index To i
                                
                                If Cells(open_index2, 3).Value <> 0 Then
                                    open_index = open_index2
                                    Exit For
                                
                                End If
                            Next open_index2
                        
                        
                        End If
                                    
                    
                        'find change between first open price and last close price
                        yearly_change = Cells(i, 6).Value - Cells(open_index, 3).Value
                        
                        percent_change = Round((yearly_change / Cells(open_index, 3).Value) * 100, 2)
                        
                        Cells(unique_ticker_index, 9).Value = ticker
                        Cells(unique_ticker_index, 10).Value = yearly_change
                        Cells(unique_ticker_index, 11).Value = percent_change & "%"
                        Cells(unique_ticker_index, 12).Value = total_volume
                        
                        'Add conditoinal formatting to color change cells
                                        
                        If Cells(unique_ticker_index, 10).Value > 0 Then
                            Cells(unique_ticker_index, 10).Interior.ColorIndex = 4
                        
                         ElseIf Cells(unique_ticker_index, 10).Value < 0 Then
                            Cells(unique_ticker_index, 10).Interior.ColorIndex = 3
                        
                        Else
                        
                        End If
                        
                          If Cells(unique_ticker_index, 11).Value > 0 Then
                            Cells(unique_ticker_index, 11).Interior.ColorIndex = 4
                        
                         ElseIf Cells(unique_ticker_index, 11).Value < 0 Then
                            Cells(unique_ticker_index, 11).Interior.ColorIndex = 3
                        
                        Else
                        
                        End If
                        
                    
                    End If
                
                    unique_ticker_index = unique_ticker_index + 1
                    
                    open_index = i + 1
                    
                    total_volume = 0
                    
        
            'if ticker is equal to next i, sum total volume
            Else
                total_volume = total_volume + Cells(i, 7).Value
                
                
            End If
        
        Next i
    
    'find max from column k
    greatest_increase = WorksheetFunction.Max(Range("K2:K" & last_row)) * 100
    Cells(2, "Q").Value = greatest_increase & "%"
    
    
    'use match method to pull row - 1 of max
    greatest_increase_row = WorksheetFunction.Match((Cells(2, "Q").Value), Range("K2:K" & last_row), 0)
    Cells(2, "P").Value = Cells(greatest_increase_row + 1, "I").Value
    
    'find min from column k
    greatest_decrease = WorksheetFunction.Min(Range("K2:K" & last_row)) * 100
    Cells(3, "Q").Value = greatest_decrease & "%"
    
    
    'use match method to pull row - 1 of min
    greatest_decrease_row = WorksheetFunction.Match((Cells(3, "Q").Value), Range("K2:K" & last_row), 0)
    Cells(3, "P").Value = Cells(greatest_decrease_row + 1, "I").Value

    'find max from column l
    greatest_volume = WorksheetFunction.Max(Range("L2:L" & last_row))
    Cells(4, "Q").Value = greatest_volume
    
    
    'use match method to pull row-1 of max volume
    greatest_volume_row = WorksheetFunction.Match((Cells(4, "Q").Value), Range("L2:L" & last_row), 0)
    Cells(4, "P").Value = Cells(greatest_volume_row + 1, "I").Value
    

    Next ws

    MsgBox ("Stock analysis complete")
    
End Sub
