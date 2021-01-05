Sub headers():
    
    'Set Headers
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"
    
    Range("P1") = "Ticker"
    Range("Q1") = "Value"
    
    Range("O2") = "Greatest % Increase"
    Range("O3") = "Greatest % Decrease"
    Range("O4") = "Greatest Total Volume"

    
    'Set Variables
    Dim i As Long
    Dim j As Integer
    Dim change As Double
    Dim percentchange As Double
    Dim total As Double
    Dim start As Long
    Dim rowCount As Long
    
    'Dim days As Integer
    'Dim dailyChange As Double
    'Dim averageChange As Double
    
    
    ' Set Values
    j = 0
    total = 0
    change = 0
    start = 2
    
    
    'Last row with data is saved
    rowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To rowCount
    
        'If Ticker doesn't match, then print results from above
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            'Data is stored in the variables
            total = total + Cells(i, 7).Value
            
            'Handle volume
            If total = 0 Then
                'print the results
                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = 0
                Range("K" & 2 + j).Value = 0
                Range("L" & 2 + j).Value = 0
    

            Else
                'Find First value that isn't 0
                If Cells(start, 3) = 0 Then
                    For find_value = start To 1
                        If Cells(find_value, 3).Value <> 0 Then
                            start = find_value
                            Exit For
                        End If
                    Next find_value
                End If
                
                'Find the Change from the Close to the Open
                change = (Cells(i, 6) - Cells(start, 3))
                percentchange = Round((change / Cells(start, 3) * 100), 2)
                
                        
                'start next ticker
                start = i + 1
                
                'print the data
                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = Round(change, 2)
                Range("K" & 2 + j).Value = "%" & percentchange
                Range("L" & 2 + j).Value = total
                
                'Green for positive changes and Red for negative changes
                Select Case change
                    Case Is > 0
                        Range("J" & 2 + j).Interior.ColorIndex = 4
                    Case Is < 0
                        Range("J" & 2 + j).Interior.ColorIndex = 3
                    Case Else
                        Range("J" & 2 + j).Interior.ColorIndex = 0
                End Select
            End If
            
            
            'reset values to start the new ticker
            total = 0
            change = 0
            j = j + 1
            Days = 0
            
            
        'If the tickers match, add the results
        Else
            total = total + Cells(i, 7).Value
            
        End If
    Next i
                        

End Sub




