'Rafael Brennand - August 25th, 2020


Sub Stocks()
    
    Dim Ticker As String
    
    Dim YrChange As Double
    YrChange = 0
    
    
    Dim PctChange As Double
    PctChange = 0
    
    Dim cl As Double
    cl = 0
    
    Dim TV As Double
    TV = 0
    
    Dim TickerSummary As Integer
    TickerSummary = 2
    
    For Each ws In Worksheets
    
    
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        Dim op As Double
        op = ws.Cells(2, 3).Value
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"
       
        For i = 2 To LastRow
        
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                Ticker = ws.Cells(i, 1).Value
                
                cl = ws.Cells(i, 6).Value
                
                YrChange = cl - op
                    
                    If YrChange = 0 Then
                     PctChange = 0
                    ElseIf op = 0 Then
                     PctChange = 0
                     Else
                    PctChange = YrChange / op
                    End If
                
                TV = TV + ws.Cells(i, 7).Value
                
                ws.Range("I" & TickerSummary).Value = Ticker
                
                ws.Range("J" & TickerSummary).Value = YrChange
                
                ws.Range("K" & TickerSummary).Value = FormatPercent(PctChange)
                
                ws.Range("L" & TickerSummary).Value = TV
                
                    If YrChange >= 0 Then
                        ws.Range("J" & TickerSummary).Interior.ColorIndex = 4
                    Else
                        ws.Range("J" & TickerSummary).Interior.ColorIndex = 3
                    End If
                                        
                
                TickerSummary = TickerSummary + 1
                
                YrChange = 0
                
                op = ws.Cells(i + 1, 3)
                
                TV = 0
                
                Else
                
                TV = TV + ws.Cells(i, 7).Value
            
                
            End If
            
        Next i
        
    YrChange = 0
    PctChange = 0
    cl = 0
    TV = 0
    TickerSummary = 2
    
    Dim MinPctChange As Double
    MinYrChange = 0
        
    Dim MaxPctChange As Double
    MaxPctChange = 0
        
    Dim MaxTV As Double
    MaxTV = 0
        
    Dim TickerMaxPC As String
        
    Dim TickerMinPC As String
        
    Dim TickerMTV As String
    
    
    
        For j = 2 To LastRow
        
           
            If ws.Cells(j, 11).Value >= MaxPctChange Then
                MaxPctChange = ws.Cells(j, 11).Value
                TickerMaxPC = ws.Cells(j, 9).Value
            End If
            '--------------------------
            If ws.Cells(j, 11).Value <= MinPctChange Then
                MinPctChange = ws.Cells(j, 11).Value
                TickerMinPC = ws.Cells(j, 9).Value
            End If
            '--------------------------
            If ws.Cells(j, 12).Value >= MaxTV Then
                MaxTV = ws.Cells(j, 12).Value
                TickerMTV = ws.Cells(j, 9).Value
            End If
            
        Next j
        
        ws.Range("O" & 2).Value = "Greatest % Increase"
        ws.Range("O" & 3).Value = "Greatest % Decrease"
        ws.Range("O" & 4).Value = "Greatest Total Volume"
        ws.Range("P" & 1).Value = "Ticker"
        ws.Range("P" & 2).Value = TickerMaxPC
        ws.Range("P" & 3).Value = TickerMinPC
        ws.Range("P" & 4).Value = TickerMTV
        ws.Range("Q" & 1).Value = "Value"
        ws.Range("Q" & 2).Value = FormatPercent(MaxPctChange)
        ws.Range("Q" & 3).Value = FormatPercent(MinPctChange)
        ws.Range("Q" & 4).Value = MaxTV
        
    Next ws
    
    
End Sub


