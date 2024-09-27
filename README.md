# VBA-challenge
Sub stock_analysis():

    Dim total As Double
    Dim x As Long
    Dim y As Long
    Dim change As Double
    Dim start As Long
    Dim percentChange As Double
    Dim maxIncr As Long
    Dim maxDecr As Long
    Dim maxVol As Long
    

        'Set Title Row
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Quarterly Change"
    Range("K1").Value = "Total Stock Volume"

    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"

    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"

        'Set Initial Values
    y = 0
    total = 0
    change = 0
    start = 2

        'Get the row number of the last row with data
    row_Count = Cells(Rows.Count, "A").End(x1up).Row

    For x = 2 To row_Count
            'If ticker changes then print results
        If Cells(x + 1, 1).Value <> Cells(x, 1) Then
    
         'Store results in variable
            total = total + Cells(x, 7).Value
        
            'Handle Total 0 Volume
            If total = 0 Then
                'Print the results
                Range("I" & 2 + y).Value = Cells(x, 1).Value
                Range("J" & 2 + y).Value = 0
                Range("K" & 2 + y).Value = 0 & "%"
                Range("L" & 2 + y).Value = 0
            Else
                    'Find First non zero starting value
                If Cells(start, 3) = 0 Then
                    For found = start To x
                        If Cells(found, 3).Value <> 0 Then
                            start = found
                            Exit For
                        End If
                    Next found
                End If
            
                    'Calculate Change
                change = (Cells(x, 6).Value) - (Cells(start, 3).Value)
                percentChange = change / (Cells(start, 3).Value)
       
                    'Start of the next stock ticker
                start = x + 1
        
                    'Print the results
                Range("I" & 2 + y).Value = Cells(x, 1).Value
                Range("J" & 2 + y).Value = change
                Range("J" & 2 + y).NumberFormat = "0.00"
                Range("K" & 2 + y).Value = percentChange
                Range("K" & 2 + y).NumberFormat = "0.00%"
                
                    'Color positive and negative changes
                Select Case change
                    Case Is > 0
                        Range("J" & 2 + y).Interior.ColorIndex = 4
                    Case Is < 0
                        Range("J" & 2 + y).Interior.ColorIndex = 3
                    Case Else
                        Range("J" & 2 + y).Interior.ColorIndex = 0
                    End Select
            
            End If
                
                ' Reset Variables for new stock ticker
            change = 0
            y = y + 1
            total = 0
        
            ' If ticker is still the same, add results
        Else
            total = total + Cells(x, 7).Value
            
        End If
        
    Next x
    
    Range("Q2").Value = WorksheetFunction.Max(Range("K2:K" & row_Count))
    Range("Q2").NumberFormat = "0.00%"
    Range("Q3").Value = WorksheetFunction.Min(Range("K2:K" & row_Count))
    Range("Q3").NumberFormat = "0.00%"
    Range("Q4").Value = WorksheetFunction.Max(Range("L2:L" & row_Count))
       
    maxIncr = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & row_Count)), Range("K2:K" & row_Count), 0)
    maxDecr = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & row_Count)), Range("K2:K" & row_Count), 0)
    maxVol = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & row_Count)), Range("L2:L" & row_Count), 0)
    
    Range("P2").Value = Cells(maxIncr + 1, 9)
    Range("P3").Value = Cells(maxDecr + 1, 9)
    Range("P4").Value = Cells(maxVol + 1, 9)


End Sub

