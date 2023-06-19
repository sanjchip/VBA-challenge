Sub StockChallenge()

        'ADJUSTMENTS TO RUN THE SCRIPT ON EVERY WORKSHE

Dim ws As Long ' to store count value for the loop
Dim shtCount As Long  ' for count of the sheets in this workbook
shtCount = Sheets.Count
For ws = 1 To shtCount


        'SCRIPT THAT LOOPS THROUGH ALL THE STOCKS
' Arrays
Dim Ticker As String
Dim Volumes As LongPtr
Dim Yearly_Chg As Double
Dim Percent_Chg As Double
Dim OpenPrices As Single, ClosePrices As Single

Volumes = 0

Dim Summary_Table_Row As Integer

Summary_Table_Row = 2

Dim i As Long
Dim j As Integer
Dim GreatIncrease As Single
Dim GreatDecrease As Single
Dim GreatVolume As LongPtr
Dim LastRow As Long


    For i = 2 To Sheets(ws).Cells(Rows.Count, 1).End(xlUp).Row
    
        If Sheets(ws).Cells(i - 1, 1).Value <> Sheets(ws).Cells(i, 1).Value Then
        OpenPrices = Sheets(ws).Cells(i, 3).Value ' the opening price at the beginning of a given year
        End If

        If Sheets(ws).Cells(i + 1, 1).Value <> Sheets(ws).Cells(i, 1).Value Then
        ClosePrices = Sheets(ws).Cells(i, 6).Value ' the closing price at the end of that year
        Ticker = Sheets(ws).Cells(i, 1).Value
        Volume = Volume + Sheets(ws).Cells(i, 7).Value
        Yearly_Chg = ClosePrices - OpenPrices
        Percent_Chg = Yearly_Chg / OpenPrices
        Sheets(ws).Range("I" & Summary_Table_Row).Value = Ticker
        Sheets(ws).Range("J" & Summary_Table_Row).Value = Yearly_Chg
        Sheets(ws).Range("J" & Summary_Table_Row).NumberFormat = "0.00" 'Formatting the value
        Sheets(ws).Range("K" & Summary_Table_Row).Value = Percent_Chg
        Sheets(ws).Range("K" & Summary_Table_Row).NumberFormat = "0.00%" 'Formatting the value
        Sheets(ws).Range("L" & Summary_Table_Row).Value = Volume ' the total stock volume for the ticker
        Sheets(ws).Range("L" & Summary_Table_Row).NumberFormat = "0"
        Summary_Table_Row = Summary_Table_Row + 1
        Volume = 0 ' to reset the volume for the next loop
        Else
        Volume = Volume + Sheets(ws).Cells(i, 7).Value 'to add the volume into the next loop
        
        End If
        
    Next i
    
    
    
    
    For j = 2 To Sheets(ws).Cells(Rows.Count, 10).End(xlUp).Row
    
        'formatting the color
        
        If Sheets(ws).Cells(j, 10).Value > 0 Then
        Sheets(ws).Cells(j, 10).Interior.ColorIndex = 4
        Else
        Sheets(ws).Cells(j, 10).Interior.ColorIndex = 3
        End If
        
        'Greatest Increase, decrease and total volume
     

	LastRow = Sheets(ws).Range("K" & Rows.Count).End(xlUp).Row   
      
    If Sheets(ws).Cells(j, 11).Value = WorksheetFunction.Max(Sheets(ws).Range("K2:K" & LastRow)) Then
        GreatIncrease = WorksheetFunction.Max(Sheets(ws).Range("K2:K" & LastRow))
        Sheets(ws).Range("P3") = GreatIncrease
        Ticker = Sheets(ws).Cells(j, 9)
        Sheets(ws).Range("O3") = Ticker
        End If
        
        If Sheets(ws).Cells(j, 11).Value = WorksheetFunction.Min(Sheets(ws).Range("K2:K" & LastRow)) Then
        GreatDecrease = WorksheetFunction.Min(Sheets(ws).Range("K2:K" & LastRow))
        Sheets(ws).Range("P4") = GreatDecrease
        Ticker = Sheets(ws).Cells(j, 9)
        Sheets(ws).Range("O4") = Ticker
        End If
        
        If Sheets(ws).Cells(j, 12).Value = WorksheetFunction.Max(Sheets(ws).Range("L2:L" & LastRow)) Then
        GreatVolume = WorksheetFunction.Max(Sheets(ws).Range("L2:L" & LastRow))
        Sheets(ws).Range("P5") = GreatVolume
        Ticker = Sheets(ws).Cells(j, 9)
        Sheets(ws).Range("O5") = Ticker
        End If
    Next j

'adding headers
Sheets(ws).Range("I1").Value = "Ticker"
Sheets(ws).Range("J1").Value = "Yearly Change"
Sheets(ws).Range("K1").Value = "Percent Change"
Sheets(ws).Range("L1").Value = "Total Stock Volume"
Sheets(ws).Range("O1").Value = "Ticker"
Sheets(ws).Range("P1").Value = "Value"

Sheets(ws).Range("N3").Value = "Greatest % increase"
Sheets(ws).Range("N4").Value = "Greatest % increase"
Sheets(ws).Range("N5").Value = "Greatest volume"

'formatting
Sheets(ws).Range("P3").NumberFormat = "0.00%"
Sheets(ws).Range("P4").NumberFormat = "0.00%"
     
       
Next ws

End Sub






