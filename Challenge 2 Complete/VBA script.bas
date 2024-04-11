Attribute VB_Name = "Module1"
Sub StockTickerData()
    
    Dim lastRow As Double
    Dim ws As Worksheet
    Dim i As Double
    Dim sumtablerow As Double
    Dim total As Double
    Dim startprice As Double
    Dim endprice As Double
    Dim ticker As String
    Dim MaxInc As Double
    Dim MaxDec As Double
    Dim MaxTotal As Double
    
    
    
    sumtablerow = 2
    lastTablerow = Cells(Rows.Count, 10).End(xlUp).Row

'You must run it twice to get the color formatting and summary table results

    For Each ws In Worksheets
     
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Volume"
        ws.Range("P2") = "Greatest % Increase"
        ws.Range("R2:R3").NumberFormat = "0.00%"
        ws.Range("P3") = "Greatest % Decrease"
        ws.Range("P4") = "Greatest Total Volume"
        ws.Range("Q1") = "Ticker"
        ws.Range("R1") = "Value"
     
     
     
        
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        lastTablerow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
      
    
        
        
        For i = 2 To lastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ticker = ws.Cells(i, 1).Value
                total = total + ws.Cells(i, 7).Value
                endprice = ws.Cells(i, 6).Value
                
                ws.Range("I" & sumtablerow).Value = ticker
                ws.Range("J" & sumtablerow).Value = endprice - startprice
                ws.Range("K" & sumtablerow).Value = (endprice - startprice) / startprice
                ws.Range("K" & sumtablerow).NumberFormat = "0.00%"
                ws.Range("L" & sumtablerow).Value = total
                
                sumtablerow = sumtablerow + 1
                
                total = 0
                
            Else
                total = total + ws.Cells(i, 7).Value
                
                If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                startprice = ws.Cells(i, 3).Value
                
                End If
                
                
            End If
              
        Next i
        
                
    
        For i = 2 To lastTablerow
        
                ws.Range("R2") = MaxInc

                ws.Range("R2").NumberFormat = "0.00%"

                MaxInc = WorksheetFunction.Max(ws.Range("K2:K" & lastTablerow))
                
                ws.Range("R3") = MaxDec

                ws.Range("R3").NumberFormat = "0.00%"

                MaxDec = WorksheetFunction.Min(ws.Range("K2:K" & lastTablerow))
                
                ws.Range("R4") = MaxTotal

                MaxTotal = WorksheetFunction.Max(ws.Range("L2:L" & lastTablerow))
                
            
            If (ws.Cells(i, 10).Value >= 0) Then
           
                ws.Cells(i, 10).Interior.ColorIndex = 4
                
                
        
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
                
            
        
            End If
            
        Next i
        
        For i = 2 To lastTablerow
        
            If (ws.Cells(i, 11).Value = MaxInc) Then
                ws.Range("Q2") = ws.Cells(i, 9).Value
            
            ElseIf (ws.Cells(i, 11).Value = MaxDec) Then
                ws.Range("Q3") = ws.Cells(i, 9).Value
            
            ElseIf (ws.Cells(i, 12).Value = MaxTotal) Then
                ws.Range("Q4") = ws.Cells(i, 9).Value
                
            Else
            
            End If
            
                
         Next i
            
        
            
        sumtablerow = 2
    
    Next ws


End Sub

