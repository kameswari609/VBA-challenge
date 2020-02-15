Sub summarystats()

    Dim totalstockvolume As LongLong
    
    Dim pricebegining As Double
    Dim pricending  As Double
    Dim ticketsymbol As String
    Dim ws As Worksheet
    
    Dim minpercentage As Double
    Dim minticketvalue As String
    Dim maxpercentage As Double
    Dim maxticketvalue As String
    
          Dim greateststockvolume As LongLong
          Dim ticketvaluestock As String
          
    
    For Each ws In Worksheets
      Dim Summary_Table_Row As Integer
      Summary_Table_Row = 2
        
      minpercentage = 0
      minticketvalue = "A"
      totalstockvolume = 0
      pricebegining = 0
      pricending = 0
      maxpercentage = 0
      maxticketvalue = "A"
      greateststockvolume = 0
      ticketvaluestock = "erv"
      

      
         
      ' Determine the Last Row
      LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
      ws.Range("I1").Value = "TICKET SYMBOL"
       
      ws.Range("J1").Value = "YEAR CHANGE"
      ws.Range("K1").Value = "PERCENTAGE YEAR CHANGE"
      ws.Range("L1").Value = "TOTAL STOCK VOLUME "
    
      pricebegining = ws.Cells(2, 3).Value
     
      
      For i = 2 To LastRow
    
    
        ' Check if we are still within a ticket symbol, if it is not...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
          ticketsymbol = ws.Cells(i, 1).Value
          ws.Range("I" & Summary_Table_Row).Value = ticketsymbol
            
          priceending = ws.Cells(i, 6).Value
       
          ws.Range("J" & Summary_Table_Row).Value = (priceending - pricebegining)
         
          If pricebegining > 0 Then
         
            ws.Range("K" & Summary_Table_Row).Value = ((priceending / pricebegining) - 1)
        
          End If
          
          totalstockvolume = (totalstockvolume + ws.Cells(i, 7).Value)
        
          ws.Range("L" & Summary_Table_Row).Value = totalstockvolume
           
          pricebegining = ws.Cells(i + 1, 3).Value
          
          If ws.Range("K" & Summary_Table_Row).Value < minpercentage Then
            minpercentage = ws.Range("K" & Summary_Table_Row).Value
            minticketvalue = ticketsymbol
          End If
          
          
           If ws.Range("K" & Summary_Table_Row).Value > maxpercentage Then
            maxpercentage = ws.Range("K" & Summary_Table_Row).Value
            maxticketvalue = ticketsymbol
          End If
          
          If ws.Range("L" & Summary_Table_Row).Value > greateststockvolume Then
            greateststockvolume = ws.Range("L" & Summary_Table_Row).Value
            ticketvaluestock = ticketsymbol
          End If
    
          ' Add one to the summary table row
          Summary_Table_Row = Summary_Table_Row + 1
        
          totalstockvolume = 0
     
          priceending = 0
     
          
        ElseIf ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
      
          totalstockvolume = (totalstockvolume + ws.Cells(i, 7).Value)
    
        End If
    
      Next i
     
      ws.Columns("K:K").NumberFormat = "0.00%"
      ws.Range("K" & Summary_Table_Row).Style = "Percent"
      Dim rng As Range
      Dim condition1 As FormatCondition, condition2 As FormatCondition
    
      'Fixing/Setting the range on which conditional formatting is to be desired
      Set rng = ws.Range("J:J")
    
      
    
      'Defining and setting the criteria for each conditional format
      Set condition1 = rng.FormatConditions.Add(xlCellValue, xlGreater, "=0")
      Set condition2 = rng.FormatConditions.Add(xlCellValue, xlLess, "=0")
    
      'Defining and setting the format to be applied for each condition
      With condition1
       
        .Interior.ColorIndex = 4
       
      End With
    
      With condition2
        .Interior.ColorIndex = 3
         
      End With
      
      
      ws.Range("O3").Value = "greatest %decrease"
       ws.Range("O4").Value = "greatest % increase"
         ws.Range("O5").Value = "greatest total stock volume"
     ws.Range("P2").Value = "Ticket"
     ws.Range("Q2").Value = "volume"
     ws.Range("P3").Value = minticketvalue
     ws.Range("Q3").Value = minpercentage
     ws.Range("Q3").Style = "Percent"
     ws.Range("Q4").Style = "Percent"
      ws.Range("P4").Value = maxticketvalue
     ws.Range("Q4").Value = maxpercentage
     ws.Range("P5").Value = ticketvaluestock
     ws.Range("Q5").Value = greateststockvolume
      
      ws.Range("O:O").Column.AutoFit
      
      
    
    Next ws
     
     
     
End Sub

