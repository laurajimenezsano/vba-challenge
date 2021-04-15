
' ## Setting values

Option Explicit

Sub Ticker()

'Setting loop for sheets

Dim ws As Worksheet
For Each ws In Worksheets

        Dim nextrow, i, change_oc, max, mostinc, min, mostdec, most, mostvolume As Long
        Dim summary_table, open_price, close_price, total, tag, newtag, mosttag As String
        Dim difference, percent_change As Double
    
        nextrow = 2
        total = 0
        
        'insert column names for ticker, yearly change, percent change and total stock volume, greatest % increase, greatest % decrease and greatest total volume
        ws.Range("M1,S1").Value = "Ticker"
        ws.Range("N1").Value = "Yearly Change"
        ws.Range("T1").Value = "Value"
        ws.Range("O1").Value = "Percent Change"
        ws.Range("P1").Value = "Total Stock Volume"
        ws.Cells(2, 18).Value = "Greatest % Increase"
        ws.Cells(3, 18).Value = "Greatest % Decrease"
        ws.Cells(4, 18).Value = "Greatest Total Volume"
        
        
        'find the unique tickers and add them into the summary table
        
        
            ws.Cells(2, 13).Value = ws.Cells(2, 1).Value
            ws.Cells(2, 11).Value = ws.Cells(2, 3).Value
        
               
            For i = 2 To ws.Range("A2").End(xlDown).Row
              
              If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
                total = ws.Cells(i, 7).Value + total
              
              Else
                summary_table = ws.Cells(i, 1).Value
                total = ws.Cells(i, 7).Value + total
                open_price = ws.Cells(i + 1, 3).Value
                close_price = ws.Cells(i, 6).Value
                ws.Range("M" & nextrow) = summary_table
                ws.Range("K" & nextrow + 1) = open_price
                ws.Range("L" & nextrow) = close_price
                ws.Range("P" & nextrow) = total
                nextrow = nextrow + 1
                total = 0
              End If
                
            Next i
            
        'values for difference and percent change for each year by ticker
        
            nextrow = 2
        
            For change_oc = 2 To ws.Range("M2").End(xlDown).Row
                  difference = ws.Cells(change_oc, 12).Value - ws.Cells(change_oc, 11).Value
                  If ws.Cells(change_oc, 12).Value = 0 Then
                    percent_change = 0
                  Else
                    percent_change = difference / ws.Cells(change_oc, 12).Value
                End If
                  ws.Range("N" & nextrow) = difference
                  ws.Range("O" & nextrow) = percent_change
                  ws.Range("O" & nextrow) = Format(ws.Range("O" & nextrow), "Percent")
                  
        'conditonal formatting
        
                  If difference > 0 Then
                    ws.Range("N" & nextrow).Interior.ColorIndex = 4
                    ws.Range("O" & nextrow).Interior.ColorIndex = 4
                    ElseIf difference < 0 Then
                    ws.Range("N" & nextrow).Interior.ColorIndex = 3
                    ws.Range("O" & nextrow).Interior.ColorIndex = 3
                    Else
                    End If
                    
                  nextrow = nextrow + 1
                  
            Next change_oc
        
        'clear calculations
        
        ws.Columns(11).ClearContents
        ws.Columns(12).ClearContents
        
        'bonus: greatest % increase, greatest % decrease, greatest total volume traded
        
        max = 0
        For mostinc = 2 To ws.Range("O2").End(xlDown).Row
            If ws.Cells(mostinc, 15).Value > max Then
            max = ws.Cells(mostinc, 15).Value
            tag = ws.Cells(mostinc, 13).Value
            
            End If
        Next mostinc
        
        min = 0
        For mostdec = 2 To ws.Range("O2").End(xlDown).Row
            If ws.Cells(mostdec, 15).Value < min Then
            min = ws.Cells(mostdec, 15).Value
            newtag = ws.Cells(mostdec, 13).Value
            
            End If
        Next mostdec
        
        most = 0
        For mostvolume = 2 To ws.Range("P2").End(xlDown).Row
            If ws.Cells(mostvolume, 16).Value > most Then
            most = ws.Cells(mostvolume, 16).Value
            mosttag = ws.Cells(mostvolume, 13).Value
            
            End If
        Next mostvolume
        
        ws.Range("S2").Value = tag
        ws.Range("S3").Value = newtag
        ws.Range("S4").Value = mosttag
        ws.Range("T2").Value = max
        ws.Range("T3").Value = min
        ws.Range("T4").Value = most
        ws.Range("T2") = Format(ws.Range("T2"), "Percent")
        ws.Range("T3") = Format(ws.Range("T3"), "Percent")
        ws.Columns("M:T").AutoFit
Next ws

End Sub

