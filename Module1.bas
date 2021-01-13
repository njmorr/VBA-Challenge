Attribute VB_Name = "Module1"
Sub TickerTapeMuliSheets()
    
    For Each ws In Worksheets
    
        Dim ticker As String
        Dim volume As Double
        volume = 0
         
        'Define headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greastest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'Define color selection
        colorRed = 22
        colorGreen = 42
     
        
        'Table locations defined
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        'Define last row
        Dim LastRow As Long
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'MsgBox (LastRow)
    
        For i = 2 To LastRow
            ticker = ws.Cells(i, 1).Value
            nextTicker = ws.Cells(i + 1, 1).Value
            previousTicker = ws.Cells(i - 1, 1).Value
            openingPrice = ws.Cells(i, 3).Value
            marketDate = ws.Cells(i, 2).Value
            closingPrice = ws.Cells(i, 6).Value
                   
                 
            'new loop for finding opening values on first day of the year
            If (ticker <> previousTicker) Then
               savedStartValue = openingPrice
                'ws.Range("J" & Summary_Table_Row).Value = SavedStartValue
            Else: openingPrice = savedStartValue
                'ws.Range("J" & Summary_Table_Row).Value = OpeningPrice
            End If
                    
            
            ' New loop for when ticker and nexticker do not match
            If ticker <> nextTicker Then
                yearlyChange = closingPrice - openingPrice
                    
                    If openingPrice > 0 Then
                        percentChange = (closingPrice - openingPrice) / openingPrice
                        
                    Else: percentChange = 0
                    
                    End If
                
                'Ticker Names inserted into new table
                ws.Range("I" & Summary_Table_Row).Value = ticker
                            
                'Volume totals totalized per ticker symbol and inserted into chart
                volume = volume + ws.Cells(i, 7).Value
                ws.Range("L" & Summary_Table_Row).Value = volume
                
                'finding closing values on last day year
                'Range("K" & Summary_Table_Row).Value = ClosingPrice
                
                'printing yearly change and % changes in new columns
                ws.Range("J" & Summary_Table_Row).Value = yearlyChange
                ws.Range("K" & Summary_Table_Row).Value = percentChange
                ws.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
                ' https://www.mrexcel.com/board/threads/vba-change-number-of-decimal-places-of-a-percentage.521221/
                ' Accessed 1/9/2021
                
                    ' if/then statement for conditional formatting
                    If yearlyChange >= 0 Then
                        ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = colorGreen
                    
                    Else:
                        ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = colorRed
                    
                    End If
                
                'Start fresh for next loop amd clear totalizers
                volume = 0
                Summary_Table_Row = Summary_Table_Row + 1
            Else
                volume = volume + ws.Cells(i, 7).Value
            
           End If
                    
        Next i

        lastRowSummary = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        Dim currentTicker As String
        Dim increaseTicker As String
        Dim decreaseTicker As String
        Dim volumeTicker As String
        
        Dim currentPercent As Double
        Dim currentVolume As Double
                
        Dim greatestPercentIncrease As Double
        greatestPercentIncrease = 0
        
        Dim greatestPercentDecrease As Double
        greatestPercentDecrease = 0
        
        Dim greatestVolume As Double
        greatestVolume = 0
 
        
        For j = 2 To lastRowSummary
        
            ' Define Relationships
            currentTicker = ws.Cells(j, 9).Value
            currentPercent = ws.Cells(j, 11).Value
            currentVolume = ws.Cells(j, 12).Value
            
            ' If/Then for greatest increase in %
            If currentPercent > greatestPercentIncrease Then
                greatestPercentIncrease = currentPercent
                increaseTicker = currentTicker
                ws.Range("P2").Value = increaseTicker
                ws.Range("Q2").Value = greatestPercentIncrease
                ws.Range("Q2").NumberFormat = "0.00%"
                ' Debug.Print ("current increase ticker is " + increaseTicker)
            End If

            ' If/Then for greatest decreae in %
            If currentPercent < greatestPercentDecrease Then
                greatestPercentDecrease = currentPercent
                decreaseTicker = currentTicker
                ws.Range("P3").Value = decreaseTicker
                ws.Range("Q3").Value = greatestPercentDecrease
                ws.Range("Q3").NumberFormat = "0.00%"
                ' Debug.Print ("current Decrease Ticker is " + decreaseTicker)
            End If
            
            ' If/Then for greatest increase in volume
            If currentVolume > greatestVolume Then
                greatestVolume = currentVolume
                volumeTicker = currentTicker
                ws.Range("P4").Value = volumeTicker
                ws.Range("Q4").Value = greatestVolume
                ws.Cells.SpecialCells(xlCellTypeVisible).EntireColumn.AutoFit
                ' https://www.thespreadsheetguru.com/the-code-vault/2014/3/25/vba-code-to-autofit-columns
                ' Accessed 12 Jan 2021
                ' Debug.Print ("current greatest volume is " + volumeTicker)
            End If
            
        Next j
        
    Next ws

End Sub

