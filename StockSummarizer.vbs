Sub ResetAllSheets():
    
    For Each Sheet In ThisWorkbook.Worksheets
        If Sheet.Name <> "Control Sheet" Then
            Worksheets(Sheet.Name).Range("I:Q").Clear
        End If
    Next Sheet
    
End Sub
Sub ResetCurrentSheet():
    Range("I:Q").Clear
End Sub
Sub SummarizeAllStockSheets()

    For Each Sheet In ThisWorkbook.Worksheets
        If Sheet.Name <> "Control Sheet" Then
            MsgBox ("Processing sheet - " + Sheet.Name)
            Worksheets(Sheet.Name).Activate
            Call SummarizeStocks
        End If
    Next Sheet
    MsgBox ("Stock summary complete. All sheets processed.")

End Sub
Sub SummarizeStocks()

    Dim currentTicker, nextTicker As String
    Dim totalStockVolume As Double
    Dim lastRow, summaryRow As Long
    Dim currentSheet, summarySheet As Worksheet
    Dim tickerOpeningPrice, tickerClosingPrice As Double
    Dim yearlyChange, percentChange As Double
    
    Call populateColumnHeader
    
    'Set initial values
    tickerOpeningPrice = Cells(2, 3).Value
    totalStockVolume = 0
    summaryRow = 2
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
    
        currentTicker = Cells(i, 1).Value
        totalStockVolume = totalStockVolume + Cells(i, 7).Value
        
        nextTicker = Cells(i + 1, 1).Value
        
        If currentTicker <> nextTicker Then
        
            'Before processing next ticker, load summary row of current ticker
            tickerClosingPrice = Cells(i, 6).Value
            yearlyChange = tickerClosingPrice - tickerOpeningPrice
            If tickerOpeningPrice <> 0 Then
                percentChange = Round((yearlyChange / tickerOpeningPrice) * 100, 2)
            End If
            
            Cells(summaryRow, 9).Value = currentTicker
            Cells(summaryRow, 10).Value = yearlyChange
            Cells(summaryRow, 11).Value = Str(percentChange) + "%"
            Cells(summaryRow, 12).Value = totalStockVolume
            
            'Conditional formatting for Yearly change column
            If yearlyChange >= 0 Then
               Cells(summaryRow, 10).Interior.Color = RGB(0, 255, 0)
            Else
               Cells(summaryRow, 10).Interior.Color = RGB(255, 0, 0)
            End If
            
            'Set values to process next ticker
            summaryRow = summaryRow + 1
            tickerOpeningPrice = Cells(i + 1, 3).Value
            tickerClosingPrice = 0
            stockVolumeTotal = 0
            
        End If
    
    Next i
    
    Call populateMaxKPISummary
    Columns("I:Q").AutoFit

End Sub
Sub populateColumnHeader()
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Range("L:L").NumberFormat = "0"
    
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(4, 15).NumberFormat = "0"

End Sub
Sub populateMaxKPISummary()

    summaryRow = Cells(Rows.Count, "I").End(xlUp).Row
    Call subMaxIncrease(summaryRow)
    Call subMinIncrease(summaryRow)
    Call subMaxStockVolume(summaryRow)

End Sub
Sub subMaxIncrease(summaryRow):
   
    Dim MaxTicker As String
    Dim MaxValue, Value As Variant
    
    MaxTicker = Cells(2, 9).Value
    MaxValue = Cells(2, 11).Value
    
    For i = 2 To summaryRow
        If Cells(i, 11).Value > MaxValue Then
            MaxTicker = Cells(i, 9).Value
            MaxValue = Cells(i, 11).Value
            Cells(i, 11).Copy
        End If
    Next i
        
    Cells(2, 16).Value = MaxTicker
    Cells(2, 17).Value = Str(MaxValue)
    Cells(2, 17).PasteSpecial Paste:=xlPasteFormats

End Sub
Sub subMinIncrease(summaryRow):
   
    Dim MinTicker As String
    Dim MinValue, Value As Variant
    
    MinTicker = Cells(2, 9).Value
    MinValue = Cells(2, 11).Value
    
    For i = 2 To summaryRow
        If Cells(i, 11).Value < MinValue Then
            MinTicker = Cells(i, 9).Value
            MinValue = Cells(i, 11).Value
            Cells(i, 11).Copy
        End If
    Next i
    
    Cells(3, 16).Value = MinTicker
    Cells(3, 17).Value = Str(MinValue)
    Cells(3, 17).PasteSpecial Paste:=xlPasteFormats
End Sub
Sub subMaxStockVolume(summaryRow):
   
    Dim MaxTicker As String
    Dim MaxValue As Variant
    
    MaxTicker = Cells(2, 9).Value
    MaxValue = Cells(2, 12).Value
    
    For i = 2 To summaryRow
        If Cells(i, 12).Value > MaxValue Then
            MaxTicker = Cells(i, 9).Value
            MaxValue = Cells(i, 12).Value
            Cells(i, 12).Copy
        End If
    Next i
    
    Cells(4, 16).Value = MaxTicker
    Cells(4, 17).Value = Str(MaxValue)
    Cells(4, 17).PasteSpecial Paste:=xlPasteFormats

End Sub
