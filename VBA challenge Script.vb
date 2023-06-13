Sub VBA_Challenge():

'Step One: Create column titles for new columns'
    For Each ws In Worksheets
    
    Dim worksheetname As String
    
'added cell titles'
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Value"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest Percent Increase"
        ws.Cells(3, 15).Value = "Greatest Percent Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
    
        
        'declare last row
        
        LastRow = ws.Cells(Rows.count, 1).End(xlUp).Row
        


'declare variables here'
        
        Dim yearlychange As Double
        Dim percentchange As Double
        Dim TSV As LongLong
        Dim ticker As String
        Dim yearopen As Double
        Dim yearclose As Double
        Dim tickerlist As Integer
        Dim column As Integer
        Dim year As Long
        Dim i As Long
        Dim count As Long
        Dim count_row As Integer
        Dim percent_list As Integer
        Dim StoVol As Long
        Dim GVT As String
        Dim tickerup As String
        Dim tickerdown As String
        Dim greatestincrease As Double
        Dim greatestdecrease As Double
        Dim great_vol As LongLong
        
        
        
        
        
        
        'Defined varibles'
        
        tickerlist = 2
        column = 1
        year = 2
        count_row = 2
        count_total = 0
        percent_list = 2
        TSV = 0
        StoVol = 2
        greatestincrease = 0
        greatestdecrease = 0
        
       
       
      

    
'Calculated how to find the Ticker symbol'

            For i = 2 To LastRow
                
                
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
                    ticker = ws.Cells(i, 1).Value
                    ws.Range("I" & tickerlist).Value = ticker
                    
                    
                
                tickerlist = tickerlist + 1
                
                
                End If
                
'this is the forumla for finding the toal stock volume,
            
            Next i
            
            
            For i = 2 To LastRow
                
                TSV = TSV + ws.Cells(i, 7).Value
                
                 If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
                    TSV = TSV + ws.Cells(i, 7).Value
                    ws.Range("L" & StoVol).Value = TSV
                    StoVol = StoVol + 1
                    TSV = 0
                    
                End If
                
                If TSV > great_vol Then
                great_vol = TSV
                GVT = ws.Cells(i, 1).Value
                ws.Cells(4, 16).Value = GVT
                ws.Cells(4, 17).Value = great_vol
                
                End If
                
            Next i
           
            

  'Below if the for loop that contains the caculations for determing year open, year close, yearly change, and percent change varibles.
        
            
For i = 2 To LastRow

    yearopen = ws.Cells(i, 3).Value
    
    'set conditionals
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
      
        yearclose = ws.Cells(i, 6).Value
        
        'yearlyclose calculation
        yearlyclose = yearclose - yearopen
        
        ' Write the yearly close value to column J
        ws.Range("J" & year).Value = yearlyclose
        year = year + 1
        
        'add in calculation for % change'
        
        
            percentchange = Round((yearlyclose / yearopen) * 100, 2)
            ws.Range("K" & percent_list).Value = percentchange & "%"
            percent_list = percent_list + 1
        
    ' these are the conditionals to determinegreatest percent increase and greatest percent decrease.
    
            If percentchange < greatestdecrease Then
                greatestdecrease = percentchange
                tickerdown = ws.Cells(i, 1).Value
                ws.Cells(3, 16).Value = tickerdown
                ws.Cells(3, 17).Value = greatestdecrease
            End If
           
           
            If percentchange > greatestincrease Then
               greatestincrease = percentchange
               tickerup = ws.Cells(i, 1).Value
               ws.Cells(2, 16).Value = tickerup
               ws.Cells(2, 17).Value = greatestincrease
           End If
   
    End If
    
        
        
        
' This is the formula for changing the color based on the value of the percent change. Negative values are supposed to be red, and positive green.
        If yearlyclose >= 0 Then
            ws.Cells(year, 10).Interior.ColorIndex = 4 ' Green
        ElseIf yearlyclose <= -0.01 Then
            ws.Cells(year, 10).Interior.ColorIndex = 3 ' Red
        End If
  
       
Next i


 'now that the calculations are done, it loops to next work sheet


    Next ws


End Sub

