Attribute VB_Name = "Module1"
Sub Stocks()

Dim Ticker_Name As String
    Dim No_Of_Rows As Long
    Dim Summary_Table_Row As Integer
    Dim Stock_Total As Double
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Quartely_Change As Double
    Dim Percent_Change As Double
    Dim cell As Range
    Dim Start_Row As Long
    Dim Percent_Increase As Double
    Dim Percent_Decrease As Double
    Dim Total_Volume As Double
    
    'Set range to determine searches
        Find_Percent = Range("k:k")
        Find_Volume = Range("L:L")
    'Use max/min to find the values
        Percent_Increase = Application.WorksheetFunction.Max(Find_Percent)
        Cells(2, 17).Value = Percent_Increase
        Percent_Decrease = Application.WorksheetFunction.Min(Find_Percent)
        Cells(3, 17).Value = Percent_Decrease
        Total_Volume = Application.WorksheetFunction.Max(Find_Volume)
        Cells(4, 17).Value = Total_Volume
    
    ' Set title row
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Quartely Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
    
    ' Initialize
        Stock_Total = 0
        Summary_Table_Row = 2
        No_Of_Rows = Cells(Rows.Count, 1).End(xlUp).Row
        Start_Row = 2
    
    ' Loop through all rows
        For i = 2 To No_Of_Rows
        ' Check to see if stock ticker changes or it's the last row
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ' Set the Ticker name
            Ticker_Name = Cells(i, 1).Value
            ' Add to the Stock Total
            Stock_Total = Stock_Total + Cells(i, 7).Value
            ' Set open price from the start of the data for this ticker
            Open_Price = Cells(Start_Row, 3).Value
            ' Closing price at the end of the ticker
            Close_Price = Cells(i, 6).Value
            ' Calculate Yearly Change & Percent Change
            Quartely_Change = Close_Price - Open_Price
            If Open_Price <> 0 Then
                Percent_Change = Quartely_Change / Open_Price
            Else
                Percent_Change = 0
            End If
            ' Format & Print the results in the Summary Table:
            Range("I" & Summary_Table_Row).Value = Ticker_Name
            Range("J" & Summary_Table_Row).Value = Quartely_Change
            Range("J" & Summary_Table_Row).NumberFormat = "0.00"
            Range("K" & Summary_Table_Row).Value = Percent_Change
            Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            Range("L" & Summary_Table_Row).Value = Stock_Total
            
            ' Color code based on change
            If Quartely_Change > 0 Then
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 4  ' Green
            ElseIf Quartely_Change < 0 Then
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 3  ' Red
            Else
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 0  ' No color
            
            End If
            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            ' Reset the Total and update the start row for the next ticker

        Stock_Total = 0
        Start_Row = i + 1
        Else
            ' Add to the Stock Total
            Stock_Total = Stock_Total + Cells(i, 7).Value
        End If
        
Next i

End Sub

