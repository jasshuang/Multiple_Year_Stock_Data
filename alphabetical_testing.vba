Sub ticker_vol()


Dim ws As Worksheet


For Each ws In Worksheets

    ' dim the variables
    
    Dim currentrow As Long
    Dim closeproce As Double
    Dim change As Double
    Dim lastrow As Long
    Dim openprice As Double
    Dim percentchange As Double
    Dim totalvolume As Double
    
    
    ' assign the variables
    lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    currentrow = 2
    totalvolume = 0
    change = 0
    
    'copy the ticker and remove duplicate
    ws.Range("A2:A" & lastrow).Copy Destination:=ws.Range("I1")
    ws.Range("I1:I" & lastrow).RemoveDuplicates Columns:=1, Header:=xlYes
    
    'loop through the rows to get the total
    For i = 2 To lastrow
        
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            totalvolume = totalvolume + ws.Range("G" & i).Value
            ws.Range("L" & currentrow).Value = totalvolume
            
            openprice = ws.Cells(currentrow, 3).Value
            closeprice = ws.Cells(i, 6).Value
            change = closeprice - openprice
            percentchange = change / openprice
            
            ws.Range("J" & currentrow).Value = change
            ws.Range("K" & currentrow).Value = percentchange
                
                'conditional formatting the cells
                If percentchange > 0 Then
                ws.Range("K" & currentrow).Interior.ColorIndex = 4
                Else
                ws.Range("K" & currentrow).Interior.ColorIndex = 3
                End If
    
            
            totalvolume = 0
            currentrow = currentrow + 1
            
            
            Else
            totalvolume = totalvolume + ws.Range("G" & i).Value
        
            End If
            
     
        
    Next i
    
    'Title of the columns
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    

'working on the bonus
    Dim newlastrow As Long
    Dim increase As Long
    Dim decrease As Long
    Dim greatestvolume As Double
    
    newlastrow = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    'finding greatest % increase value
    ws.Range("Q2").Value = WorksheetFunction.Max(Range("K2" & ":" & "K" & newlastrow))
    increase = ws.Range("Q2").Value
    
    'finding greatest % increase ticker
    For i = 2 To newlastrow
    
    If ws.Cells(i, "K") = ws.Range("Q2").Value Then
    ws.Cells(2, "P").Value = ws.Cells(i, "I").Value
    End If
    
    Next i
    
    'finding greatest % decrease value
    ws.Range("Q3").Value = WorksheetFunction.Min(Range("K2" & ":" & "K" & newlastrow))
    decreae = ws.Range("Q3").Value
    
    'finding greatest % decrease ticker
    For i = 2 To newlastrow
    
    If ws.Cells(i, "K") = ws.Range("Q3").Value Then
    ws.Cells(3, "P").Value = ws.Cells(i, "I").Value
    End If
    
    Next i
    
    'finding greatest volume value
    ws.Range("Q4").Value = WorksheetFunction.Max(Range("L2" & ":" & "L" & newlastrow))
    greatestvolume = ws.Range("Q4").Value
    
    'finding greatest volume ticker
    For i = 2 To newlastrow
    
    If ws.Cells(i, "L") = ws.Range("Q4").Value Then
    ws.Cells(4, "P").Value = ws.Cells(i, "I").Value
    End If
    
    Next i
    
    'add Title to the table
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"

Next ws



End Sub
