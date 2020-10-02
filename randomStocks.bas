Attribute VB_Name = "Module11"
Sub randomStocks()

'Range("O11").Value = Time
'Dim ws As Worksheet
'Set ws = Worksheets("2014")

    For Each ws In Worksheets
        Dim tickers() As String
        Dim openDate() As Date
        Dim closeDate() As Date
        Dim openPrice() As Double
        Dim closePrice() As Double
        Dim volume() As Double
        Dim yearlyChange() As Double
        Dim pctChange() As Double
        
        Dim tickerDate As Date
    
        Dim lastRow As Long
        Dim upperBound As Long
        Dim index As Long
        
        Dim indIncr As Long
        Dim indDecr As Long
        Dim indVol As Long

        upperBound = -1
        
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To lastRow
           
            tickerDate = CDate(Format(ws.Cells(i, 2).Value, "0000/00/00"))
            
            'initialize arrays with first record
            'check if current ticker already exists in the array
            'exit the loop when ticker index is found
            'ticker is found when x is not -1
            If (i = 2) Then
                index = -1
            ElseIf (ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value) Then
                
                index = -1
                For x = 0 To upperBound
                
                    If tickers(x) = ws.Cells(i, 1).Value Then
                        index = x
                        Exit For
                    End If
            
                Next x
            End If
            
            'if ticker is not found, add ticker to the arrays
            'reset index
            'resize array to accomodate the new ticker
            If index = -1 Then
            
                upperBound = upperBound + 1
                
                ReDim Preserve tickers(upperBound)
                ReDim Preserve openDate(upperBound)
                ReDim Preserve openPrice(upperBound)
                ReDim Preserve closeDate(upperBound)
                ReDim Preserve closePrice(upperBound)
                ReDim Preserve volume(upperBound)
                
                tickers(upperBound) = ws.Cells(i, 1).Value
                openDate(upperBound) = tickerDate
                openPrice(upperBound) = ws.Cells(i, 3).Value
                closeDate(upperBound) = tickerDate
                closePrice(upperBound) = ws.Cells(i, 6).Value
                volume(upperBound) = ws.Cells(i, 7).Value
                
                index = upperBound
                
            Else
                'if ticker exists in the array, check if ticker date is the open date or the close date
                'set the open price or close price based on the ticker date
                If (tickerDate < openDate(index)) Then
                
                    openDate(index) = tickerDate
                    openPrice(index) = ws.Cells(i, 3).Value
                    
                End If
                
                If (tickerDate > closeDate(index)) Then
                
                    closeDate(index) = tickerDate
                    closePrice(index) = ws.Cells(i, 6).Value
                    
                End If
                
                'accumulate the total volume for each ticker
                volume(index) = volume(index) + ws.Cells(i, 7).Value
            
            End If
          
        Next i
        
        'calculate yearly change based on open price at start of year and
        'close price at end of year
        ReDim Preserve yearlyChange(upperBound)
        ReDim Preserve pctChange(upperBound)
        
        For k = 0 To upperBound
            
            yearlyChange(k) = closePrice(k) - openPrice(k)
            
            If (openPrice(k) > 0) Then
                pctChange(k) = Round((yearlyChange(k) / openPrice(k)), 6)
            Else
                pctChange(k) = 0
            End If
        
        Next k
    
        'calculate for
        'Greatest % Increase
        'Greatest % Decrease
        'Greates Total Volume
        indIncr = 0
        indDecr = 0
        indVol = 0
        
        'Print Summary Column Headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'List all ticker data
        For irow = 0 To upperBound
        
            ws.Range("I" & irow + 2).Value = tickers(irow)
            ws.Range("J" & irow + 2).Value = yearlyChange(irow)
            ws.Range("K" & irow + 2).Value = pctChange(irow)
            ws.Range("K" & irow + 2).Style = "Percent"
            ws.Range("K" & irow + 2).NumberFormat = "###0.00%"
            ws.Range("L" & irow + 2).Value = volume(irow)
            
            If (pctChange(irow) > 0) Then
                ws.Range("J" & irow + 2).Interior.ColorIndex = 4
            Else
                ws.Range("J" & irow + 2).Interior.ColorIndex = 3
            End If
            
            'keep track of the index with the greatest % increase, greatest % decrease
            'and greatest total volume
            'get the ticker index with the greatest % increase
            If (pctChange(irow) > pctChange(indIncr)) Then
                indIncr = irow
            End If
            'get the ticker index with the greatest % decrease
            If (pctChange(irow) < pctChange(indDecr)) Then
                indDecr = irow
            End If
            
            'get the ticker index with the greatest volume
            If (volume(irow) > volume(indVol)) Then
                indVol = irow
            End If
            
        Next irow
        
        'print analysis data 
        ws.Range("O" & 2).Value = "Greatest % Increase"
        ws.Range("O" & 3).Value = "Greatest % Decrease"
        ws.Range("O" & 4).Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        ws.Range("P2").Value = tickers(indIncr)
        ws.Range("P3").Value = tickers(indDecr)
        ws.Range("P4").Value = tickers(indVol)
        
        ws.Range("Q2").Value = pctChange(indIncr)
        ws.Range("Q3").Value = pctChange(indDecr)
        ws.Range("Q4").Value = volume(indVol)
        
        'format percent column
        ws.Range("Q2").Style = "Percent"
        ws.Range("Q2").NumberFormat = "###0.00%"
        
        ws.Range("Q3").Style = "Percent"
        ws.Range("Q3").NumberFormat = "###0.00%"
        
        ws.Columns("J:Q").AutoFit
    
    Next ws
    'Range("O12").Value = Time
End Sub






