Attribute VB_Name = "Module1"
'need to add % sign at the end of the percentage_change output
'figure out the overflow error --> how to reduce the memory usage
'figure out how to hit other worksheets in one runSub stock_summary()

Sub stock_summary()

Dim ticker As String
Dim open_price As Double
    
Dim close_price As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim stock_volume As Double
    stock_volume = 0
    
Dim summary_table_row As Double
    summary_table_row = 2
    
Dim last_row As Double
last_row = Cells(Rows.Count, 1).End(xlUp).Row

Dim ws As Worksheet

Dim max_increase As Double
Dim max_decrease As Double
Dim max_volume As Double
Dim row_number As Integer
Dim max_ticker As String




            For Each ws In Worksheets
                
            'ws.Range("K1") = "Open Price"
            'ws.Range("L1") = "Close Price"
            ws.Range("M1") = "Ticker"
            ws.Range("N1") = "Yearly Change"
            ws.Range("O1") = "Percentage Change"
            ws.Range("P1") = "Stock Volume Total"
            open_price = ws.Range("C2").Value
            
                       For i = 2 To last_row
                                                 
                                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                                    ticker = ws.Cells(i, 1).Value
                                    close_price = ws.Cells(i, 6).Value
                                                 
                                        If open_price = 0 Then
                                           close_price = yearly_change
                                           percentage_change = "100"
                                        Else
                                           yearly_change = close_price - open_price
                                           percent_change = yearly_change / open_price
                                                                                                                         
                                        End If
                                        
                                    stock_volume = stock_volume + ws.Cells(i, 7).Value
                                    ws.Range("M" & summary_table_row).Value = ticker
                                    ws.Range("N" & summary_table_row).Value = yearly_change
                                    ws.Range("O" & summary_table_row).Value = Format(percent_change, "Percent")
                                    ws.Range("P" & summary_table_row).Value = stock_volume
                                    'ws.Range("K" & summary_table_row).Value = open_price
                                    'ws.Range("L" & summary_table_row).Value = close_price
                                    
                                              
                                        If yearly_change < 0 Then
                                            ws.Range("N" & summary_table_row).Interior.Color = vbRed
                                        ElseIf yearly_change > 0 Then
                                            ws.Range("N" & summary_table_row).Interior.Color = vbGreen
                                        End If
                                            
                                    open_price = ws.Cells(i + 1, 3).Value
                                    summary_table_row = summary_table_row + 1
                                    stock_volume = 0
                                    closing_cost = 0
                                                   
                                                 
                                Else
                                    stock_volume = stock_volume + ws.Cells(i, 7).Value
                                                     
                                End If
                    
                    
                        Next i
        
            summary_table_row = 2
           
                      
            ws.Range("T2").Value = "Greatest Increase"
            ws.Range("T3").Value = "Greatest Decrease"
            ws.Range("T4").Value = "Greatest Total Volume"
            ws.Range("U1").Value = "Ticker"
            ws.Range("V1").Value = "Value"
            
            
            Set Rng = ws.Range("O:O")
            max_increase = WorksheetFunction.Max(Rng)
            row_number = WorksheetFunction.Match(max_increase, Rng, 0) + Rng.Row - 1
            max_ticker = ws.Cells(row_number, 13).Value
            
            ws.Range("U2").Value = max_ticker
            ws.Range("V2").Value = max_increase
            ws.Range("V2").Value = Format(max_increase, "Percent")
            
            Set Rng = ws.Range("O:O")
            max_decrease = WorksheetFunction.Min(Rng)
            row_number = WorksheetFunction.Match(max_decrease, Rng, 0) + Rng.Row - 1
            max_ticker = ws.Cells(row_number, 13).Value
            
            ws.Range("U3").Value = max_ticker
            ws.Range("V3").Value = max_decrease
            ws.Range("V3").Value = Format(max_decrease, "Percent")
            
            Set Rng = ws.Range("P:P")
            max_volume = WorksheetFunction.Max(Rng)
            row_number = WorksheetFunction.Match(max_volume, Rng, 0) + Rng.Row - 1
            max_ticker = ws.Cells(row_number, 13).Value
            
            ws.Range("U4").Value = max_ticker
            ws.Range("V4").Value = max_volume
            
            'max_increase = 0
            'max_decrease = 0
            'max_volume = 0
                    
                    
Next ws
    
   
End Sub

