Sub Multiple_year_stock()


For Each ws In Worksheets
   ws.Activate
        
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
        
        
    Dim Ticker As String
    Dim PriceChange As Long
    Dim percentagechange As Double
    Dim totalstockVolume As Double
    Dim Rownumber As Double
    Dim lastRow As Long
    Dim x As Long
    Dim i As Long
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
       
        Rownumber = 2
        totalstockVolume = 0
        For i = 2 To lastRow
        
            If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
                Ticker = Cells(i, 1).Value
                Range("I" & Rownumber).Value = Ticker
                Opening = Cells(i, 3).Value
                
                
            ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                Closing = Cells(i, 6).Value
                
            
            
                YC = Closing - Opening
                Range("J" & Rownumber).Value = YC
            
            
                Percentchange = YC / Opening
                Range("K" & Rownumber).Value = Percentchange
                Rownumber = Rownumber + 1
                totalstockVolume = 0
            Else
            End If
        Next i
    Rownumber = 2
        For x = 2 To lastRow
            If Cells(x + 1, 1).Value <> Cells(x, 1).Value Then
            totalstockVolume = totalstockVolume + Cells(x, 7).Value
            Range("L" & Rownumber).Value = totalstockVolume
            Rownumber = Rownumber + 1
            totalstockVolume = 0
        Else
            totalstockVolume = totalstockVolume + Cells(x, 7).Value
        End If
        Next x
        
    
            Range("K2:K" & i).NumberFormat = "0.00%"
            
            
                Range("I" & Rownumber).Value = Ticker
                Range("L" & Rownumber).Value = totalstockVolume + Cells(i, 7).Value
                
                Rownumber = Rownumber + 1
       
                
                totalstockVolume = totalstockVolume + Cells(i, "G").Value
    Dim Color As Long
    For Color = 2 To lastRow
        If Cells(Color, 10).Value < 0 Then
        Cells(Color, 10).Interior.ColorIndex = 3
        Else
        Cells(Color, 10).Interior.ColorIndex = 4
        End If
    Next Color
    
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
    
 
    Dim Value As Double
    Dim MaxPercent As Double
    Dim MinPercent As Double
    Dim GreatestTV As Double
    Dim lastRow2 As Double
    Dim Ticker2 As String
    
    
    
   lastRow2 = Cells(Rows.Count, 11).End(xlUp).Row
   For q = 2 To lastRow2
   
MaxPercent = WorksheetFunction.Max(Range("K2:K" & q))
Range("Q2") = MaxPercent


MinPercent = WorksheetFunction.Min(Range("K2:K" & q))
Range("Q3") = MinPercent

GreatestTV = WorksheetFunction.Max(Range("L2:L" & q))
Range("Q4") = GreatestTV


    Next q

    For P = 2 To lastRow2
    
    If MaxPercent = Cells(P, 11).Value Then
    Ticker2 = Cells(P, 9).Value
    Range("P2").Value = Ticker2
    
    ElseIf MinPercent = Cells(P, 11).Value Then
    Ticker2 = Cells(P, 9).Value
    Range("P3").Value = Ticker2
    
    ElseIf GreatestTV = Cells(P, 12).Value Then
    Ticker2 = Cells(P, 9).Value
    Range("P4").Value = Ticker2
    

    Else
    End If
    Next P
    
    Range("Q2:Q3").NumberFormat = "0.00%"


Next ws
    
        
End Sub
