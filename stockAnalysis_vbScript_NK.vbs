Sub stockAnalysis()
    
    For Each current In Worksheets
        
       'Variable declaration
        Dim i As Long    'current row index
        Dim j As Long   'start row index of ticker block
        Dim k As Long   'row index of the new ticker col
        Dim lastRowColA As Long 'last row of col A
        Dim lastRowColTicker As Long 'last row of col Ticker
        Dim yearlyChange As Double 'yearly change
        Dim perChange As Double 'percent change
        Dim totalVol As Double 'total stock vol
        Dim greIn As Double 'greatest increase cal
        Dim greDe As Double 'greatest decrease cal
        Dim greTotalVol As Double 'greatest total vol
        
        'Create column headers
        current.Range("I1").Value = "Ticker"
        current.Range("J1").Value = "Yearly Change"
        current.Range("K1").Value = "Percent Change"
        current.Range("L1").Value = "Total Stock Volume"
        
        'Create names of the rows and columns of the table
        current.Range("P1").Value = "Ticker"
        current.Range("Q1").Value = "Value"
        current.Range("O2").Value = "Greatest % Increase"
        current.Range("O3").Value = "Greatest % Decrease"
        current.Range("O4").Value = "Greatest Total Volume"
        
        
        'Setting initial values of variables
        j = 2
        k = 2
        lastRowA = current.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To lastRowA
            If current.Cells(i + 1, 1).Value <> current.Cells(i, 1).Value Then
                current.Cells(k, 9).Value = current.Cells(i, 1).Value
                current.Cells(k, 10).Value = current.Cells(i, 6).Value - current.Cells(j, 3).Value
            
                    If current.Cells(k, 10).Value < 0 Then
                        current.Cells(k, 10).Interior.ColorIndex = 3
                    Else
                        current.Cells(k, 10).Interior.ColorIndex = 4
                    End If
                    
                    
                    If current.Cells(j, 3).Value <> 0 Then
                        perChange = ((current.Cells(i, 6).Value - current.Cells(j, 3).Value) / current.Cells(j, 3).Value)
                        current.Cells(k, 11).Value = Format(perChange, "Percent")
                    Else
                        current.Cells(k, 11).Value = Format(0, "Percent")
                    End If
                
                    current.Cells(k, 12).Value = WorksheetFunction.Sum(Range(current.Cells(j, 7), current.Cells(i, 7)))
                    k = k + 1
                    j = i + 1
            
            End If
        Next i
        
    'Table part
        lastRowColTicker = current.Cells(Rows.Count, 9).End(xlUp).Row
        greIn = current.Cells(2, 11).Value
        greDe = current.Cells(2, 11).Value
        greTotalVol = current.Cells(2, 12).Value
    
        For i = 2 To lastRowColTicker
            If current.Cells(i, 11).Value > greIn Then
                greIn = current.Cells(i, 11).Value
                current.Cells(2, 16).Value = current.Cells(i, 9).Value
            Else
                greIn = greIn
            End If
            
            If current.Cells(i, 11).Value < greDe Then
                greDe = current.Cells(i, 11).Value
                current.Cells(3, 16).Value = current.Cells(i, 9).Value
            Else
                greDe = greDe
            End If
        
            current.Range("Q2").Value = Format(greIn, "Percent")
            current.Range("Q3").Value = Format(greDe, "Percent")
            current.Range("Q4").Value = Format(greTotalVol, "Scientific")
        
        Next i
    
    
    
    Next current
    
End Sub


