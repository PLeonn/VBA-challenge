Attribute VB_Name = "Module1"
Sub MarketData()

'Declare Variables
Dim ws As Worksheet
Dim Ticker As String
Dim SummaryTableRow As Integer
Dim Volume As Double
Volume = 0
Dim openPrice As Double
Dim closePrice As Double
Dim yearChange As Double
Dim ConForm As Range
Dim lastRow As Long
Dim percentChange As Double

    'Loop through all sheets
    For Each ws In ThisWorkbook.Worksheets

    'Headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Chng"
        ws.Cells(1, 11).Value = "% Chng"
        ws.Cells(1, 12).Value = "Total Stock Volume"

    'Last Row
        lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
        SummaryTableRow = 2
        
        'Nested loop
        For i = 2 To lastRow
        
        'IMPORTANT I can't figure out why, but my code is just changing the I2 value to the ticker, not looping to other rows. Not sure why.
        
            'Conditional
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            
                'Populates ticker column
                Ticker = Cells(i, 1).Value
                Range("I" & SummaryTableRow).Value = Ticker
                
                'Populates volume column
                Volume = Volume + Cells(i, 7).Value
                Range("L" & SummaryTableRow).Value = Volume
                
                'Calculates year change and percent change
                openPrice = Cells(i, 3).Value
                closePrice = Cells(i, 6).Value
                
                yearChange = closePrice - openPrice
                percentChange = yearChange / closePrice
                
                'Populates summary table
                Range("J" & SummaryTableRow).Value = yearChange
                Range("K" & SummaryTableRow).Value = percentChange
                
                'Resets volume to 0 and increments SummaryTableRow
                Volume = 0
                SummaryTableRow = SummaryTableRow + 1
                
                Else
                    Volume = Volume + Cells(i, 7).Value
            End If
            
            'Conditional Formatting
            Set ConForm = ws.Range("K1:K" & SummaryTableRow)
            
            For Each Cell In ConForm

                If Cell.Value > 0 Then
                    Cell.Interior.ColorIndex = 4
                
                ElseIf Cell.Value < 0 Then
                    Cell.Interior.ColorIndex = 3
                
                Else
                    Cell.Interior.ColorIndex = xlNone
                    
                End If
                
            Next
            
        Next i
        
    Next
        

End Sub
