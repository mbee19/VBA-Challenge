'Solution for VBA-Challenge
'Morgan Bee

Sub FinancialAnalysis()

' SettingDimensions

Dim QuarterlyChange As Double
Dim PercentChange As Double
Dim TotalVolume As LongLong
Dim LastRow As Long
Dim Summary_Table_Row As Long
Dim i As Long
Dim ws As Worksheet
Dim StockQuarterOpen As Double
Dim StockQuarterClose As Double
Dim Ticker As String
Dim CurrentTicker As String
Dim GreatestPercentIncrease As Double
Dim GreatestPercentDecrease As Double
Dim GreatestTotalVolume As LongLong
Dim SummaryLastRow As Integer
Dim CurrentPercentIncrease As Double
Dim CurrentPercentDecrease As Double
Dim CurrentTotalVolume As LongLong

' Loop for all sheets
For Each ws In ThisWorkbook.Sheets

' setting titles
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Quarterly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"
ws.Range("O1").Value = "Ticker"
ws.Range("P1").Value = "Value"

' find last row of data
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

' set values
QuarterlyChange = 0
PercentChange = 0
TotalVolume = 0
StockQuarterClose = 0
Summary_Table_Row = 2
StockQuarterOpen = ws.Cells(2, 3).Value

'create loop
For i = 2 To LastRow

        ' check if the ticker is different from the row above
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        ' add to total volume
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        
        ' place ticker into Summary Table
        ws.Range("I" & Summary_Table_Row).Value = ws.Cells(i, 1).Value
        
        ' place total stock volume into Summary Table
        ws.Range("L" & Summary_Table_Row).Value = TotalVolume
        
        ' Calculate Quarterly Change
        StockQuarterClose = ws.Cells(i, 6).Value
        QuarterlyChange = StockQuarterClose - StockQuarterOpen
        
        ' place quarterly change into Summary Table
        ws.Range("J" & Summary_Table_Row).Value = QuarterlyChange
        
        ' add conditional formatting to make positive Quarterly Change Green
                If QuarterlyChange > 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                
                ElseIf QuarterlyChange < 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                
                ElseIf QuarterlyChange = 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 0
                
                End If
        
        ' calculate % change
        PercentChange = ((StockQuarterClose - StockQuarterOpen) / StockQuarterOpen)
        
        ' Place% change into Summary Table, format as Percentage
        ws.Range("K" & Summary_Table_Row).Value = PercentChange
        ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
        
        ' add conditional formatting to make positive percentage change Green
                If PercentChange > 0 Then
                ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
                
                ElseIf PercentChange < 0 Then
                ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
                
                ElseIf PercentChange = 0 Then
                ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 0
                
                End If
        
        ' add row to Summary Table
        Summary_Table_Row = Summary_Table_Row + 1
        
        ' reset the Total Volume to 0
        TotalVolume = 0
        
        ' reset Quarter Open to next i value, reset StockQuarterClose, reset QuarterlyChange
        StockQuarterOpen = ws.Cells(i + 1, 3).Value
        StockQuarterClose = 0
        QuarterlyChange = 0
    
        
        ' if ticker is the same from the row above then
        Else
        
        ' add to Total Volume
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        
        
        End If
        
Next i



' calculate last row for Summary Table
' Set values for Summary Table
SummaryLastRow = Cells(Rows.Count, 9).End(xlUp).Row
GreatestPercentIncrease = 0
GreatestPercentDecrease = 0
GreatestTotalVolume = 0
Ticker = " "

        ' loop through summary table
        For i = 2 To SummaryLastRow
        CurrentPercentIncrease = ws.Cells(i, 11).Value
        CurrentTicker = ws.Cells(i, 9).Value
        
        ' create conditional statement to check for Maximum Value (Greatest Percent Increase)
                If CurrentPercentIncrease > GreatestPercentIncrease Then
                GreatestPercentIncrease = CurrentPercentIncrease
                
                Ticker = CurrentTicker
                
                End If
        
        ' place Ticker and Greatest Percentage Value into the appropriate cell
               ws.Range("O2").Value = Ticker
               ws.Range("P2").Value = GreatestPercentIncrease
               ws.Range("P2").NumberFormat = "0.00%"
               
        Next i
        
        ' Loop again and
        ' Create conditional statement to check for Minimum Value (Greatest Percent Decrease)
        For i = 2 To SummaryLastRow
        CurrentPercentDecrease = ws.Cells(i, 11).Value
        CurrentTicker = ws.Cells(i, 9).Value
                
                If CurrentPercentDecrease < GreatestPercentDecrease Then
                GreatestPercentDecrease = CurrentPercentDecrease
                
                Ticker = CurrentTicker
                    
                End If
        
        ' place Ticker and Value into the appropriate cell
                ws.Range("O3").Value = Ticker
                ws.Range("P3").Value = GreatestPercentDecrease
                ws.Range("P3").NumberFormat = "0.00%"
                
        Next i
        
        ' Loop again and
        ' Create conditional statement to check for Maximum Value (Total Volume)
        For i = 2 To SummaryLastRow
        CurrentTotalVolume = ws.Cells(i, 12).Value
        CurrentTicker = ws.Cells(i, 9).Value
        
                If CurrentTotalVolume > GreatestTotalVolume Then
                GreatestTotalVolume = CurrentTotalVolume
                
                Ticker = CurrentTicker
                
               End If
        
        Next i
        
        ' place Ticker and Value into the appropriate cell
                ws.Range("O4").Value = Ticker
                ws.Range("P4").Value = GreatestTotalVolume

Next ws

End Sub




