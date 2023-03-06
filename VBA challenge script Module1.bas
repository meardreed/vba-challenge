Attribute VB_Name = "Module1"
Sub stock_analysis()


    Dim ws As Worksheet

    'loop through all worksheets
        For Each ws In Worksheets
        
            ' declare variables
            Dim ticker As String
            Dim total_vol As Double
            Dim sum_row As Integer
            
            Dim start_ticker As Long
            Dim yearly_change As Double
            Dim open_price As Double
            Dim end_price As Double
            
            Dim percent_change As Double
            
            Dim last_row As Long
            
            
            Dim max_increase As Double
            Dim max_incticker As String
            
            Dim max_decrease As Double
            Dim max_decticker As String
            
            Dim max_volticker As String
            Dim max_vol As Double
            
            
        
        
            ' initialize variables
            total_vol = 0
            last_row = ws.Cells(Rows.Count, "A").End(xlUp).Row
            sum_row = 2
            start_ticker = 2
            yearly_change = 0
            percent_change = 0
            open_price = 0
            end_price = 0
            
            max_increase = 0
            max_decrease = 0
            max_vol = 0
            
            
            
            
            'Create Headers
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volume"
            
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"
            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatest Total Volume"
            
            
            
            ' Initial value for the first ticker
            open_price = ws.Cells(2, 3).Value
            
            
            
            'Iterate all rows
             For i = 2 To last_row
            
                'When we get to a Different ticker
                If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                
                        ' Add volume for same ticker
                         total_vol = total_vol + ws.Cells(i, 7).Value
                    
                        
                        
                        ' get ticker in summary table
                        ws.Range("I" & sum_row).Value = ws.Cells(i, 1).Value
                        
                        
                        'Define last row for same ticker
                         end_price = ws.Cells(i, 6).Value
                         
                        
                         ' calculate yearly change
                        yearly_change = end_price - open_price
                        
                        'get yearly change in table
                        ws.Range("J" & sum_row).Value = yearly_change
                        
                        
                            
                            ' color formatting for yearly change
        
                                If (yearly_change >= 0) Then
                                
                                ws.Range("J" & sum_row).Interior.ColorIndex = 4
                                
                                ElseIf (yearly_change < 0) Then
                                
                                ws.Range("J" & sum_row).Interior.ColorIndex = 3
                                
                                End If
                                
                        
                        
                            If open_price <> 0 Then
                        
                            ' calculate percent change
                            percent_change = (yearly_change / open_price)
                            
                            ' get percent change in summary table
                            
                            ws.Range("K" & sum_row).Value = FormatPercent(percent_change)
                            
                            
                            End If
                            
                            
                           
                        
                        ' get volume in summary table
                        ws.Range("L" & sum_row).Value = total_vol
                        
                        
                        
                        ' reset variable
                        sum_row = sum_row + 1
                        
                        
                        ' open price for next ticker
                        open_price = ws.Cells(i + 1, 3).Value
                        
                        
                             
                            
                             ' add calculations for max increase/decrease/vol
                            
                            If (percent_change > max_increase) Then
                                max_increase = percent_change
                                max_incticker = ticker
                                
                            ElseIf (percent_change < max_decrease) Then
                                max_decrease = percent_change
                                max_decticker = ticker
                            End If
                            
                            If (total_vol > max_vol) Then
                                max_vol = total_vol
                                max_volticker = ticker
                            End If
                            
                            
                            ' reset variable
                            percent_change = 0
                            total_vol = 0
                    
                        
                'if it is the Same ticker
                Else
                
                         
                         ' Add volume for same ticker
                         total_vol = total_vol + ws.Cells(i, 7).Value
                          
                                            
                     
                 End If
        
        
               Next i
            
                             ' add locations for max increase/decrease/vol
                            
                            ws.Range("Q2").Value = FormatPercent(max_increase)
                            ws.Range("P2").Value = max_incticker
                            ws.Range("Q3").Value = FormatPercent(max_decrease)
                            ws.Range("P3").Value = max_decticker
                            ws.Range("Q4").Value = max_vol
                            ws.Range("P4").Value = max_volticker
                            
            
        
            
            ' Formatting
            
            Columns("I:Q").Select
            Columns("I:Q").EntireColumn.AutoFit
            

    
     Next ws
    



End Sub

