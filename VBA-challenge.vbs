Sub stock_trade()

'Variable definition

    Dim ticker, acts, greati, greatd, gvol As String
    Dim total_volume, firstp, lastp, y_change, pcent_change, max, maxvol, min As Double
    Dim year_resume, lrow, lrowtable As Integer
    
    year_resume = 2
    max = 0
    min = 0
    acts = ActiveSheet.Name
    lrow = Sheets(acts).Cells(Rows.Count, 1).End(xlUp).Row
    lrowtable = Sheets(acts).Cells(Rows.Count, 9).End(xlUp).Row
    firstp = Sheets(acts).Cells(year_resume, 6).Value
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Year Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
   
   'Resume Table
   
    For I = year_resume To lrow

        If Cells(I, 1).Value <> Cells(I + 1, 1).Value Then
        
            ticker = Cells(I, 1).Value
            total_volume = total_volume + Cells(I, 7).Value
            lastp = Cells(I, 6).Value
            y_change = (lastp - firstp)
            
            If firstp = 0 Then
            pcent_change = 0
            Else
            pcent_change = y_change / firstp
            End If
            
            Cells(year_resume, 9).Value = ticker
            Cells(year_resume, 10).Value = y_change
            Cells(year_resume, 11).Value = pcent_change
            Cells(year_resume, 12).Value = total_volume
            
            firstp = Cells(I + 1, 6).Value
            year_resume = year_resume + 1
            total_volume = 0
               
        Else
            
            total_volume = total_volume + Cells(I, 7).Value
        
        End If
        
    Next I
    
'Column Format
Range("K:K").NumberFormat = "0.00%"
Range("L:L").NumberFormat = "#,##0"
    
'Max and min, greatest volume tickers and values

For I = 2 To lrowtable

    If Sheets(acts).Cells(I, 11).Value > max Then
    max = Cells(I, 11).Value
    Sheets(acts).Cells(2, 17).Value = max
    Range("Q2").NumberFormat = "0.00%"
    
    ElseIf Cells(I, 12).Value > maxvol Then
    maxvol = Cells(I, 12).Value
    Cells(4, 17).Value = maxvol
    Range("Q4").NumberFormat = "#,##0"
    
    ElseIf Cells(I, 11).Value < min Then
    min = Cells(I, 11).Value
    Cells(3, 17).Value = min
    Range("Q3").NumberFormat = "0.00%"
    
    End If
Next I

For n = 2 To lrowtable

    If Cells(n, 11).Value = Cells(2, 17).Value Then
    greati = Cells(n, 9).Value
    Cells(2, 16).Value = greati
    
    ElseIf Cells(n, 11).Value = Cells(3, 17).Value Then
    greatd = Cells(n, 9).Value
    Cells(3, 16).Value = greatd
    
    ElseIf Cells(n, 12).Value = Cells(4, 17).Value Then
    gvol = Cells(n, 9).Value
    Cells(4, 16).Value = gvol
    
    End If
    
Next n

'Conditional Format
Range("J2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = -16532384
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .Color = -16383753
        .TintAndShade = 0
    End With

End Sub

