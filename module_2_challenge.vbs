Sub Yearly_Stock()

'Variables
Dim ws As Worksheet
Dim ticker As String
Dim vol As Double
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim current_ticker As Double
Dim Total_Stock_Volume As Double


For Each ws In Worksheets
    
'Column Headers name
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
'setting starting ticker
    current_ticker = 2
    Total_Stock_Volume = 0


'Last row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

year_open = ws.Cells(2, 3).Value

For i = 2 To LastRow

Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

'Looking for ticker name change
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

    ticker = ws.Cells(i, 1).Value

    year_close = ws.Cells(i, 6).Value
    

' Calculations
    yearly_change = year_close - year_open
    percent_change = (year_close - year_open) / year_open
    


'insert values into summary table
    ws.Cells(current_ticker, 9).Value = ticker
    ws.Cells(current_ticker, 10).Value = yearly_change
    ws.Cells(current_ticker, 11).Value = percent_change
    ws.Cells(current_ticker, 12).Value = Total_Stock_Volume

'Color formating
    If ws.Cells(current_ticker, 10).Value < 0 Then
    ws.Cells(current_ticker, 10).Interior.ColorIndex = 3
    Else
    ws.Cells(current_ticker, 10).Interior.ColorIndex = 4
    End If

 
'Percent formatting
    ws.Cells(current_ticker, 11).Value = Format(percent_change, "0.00%")
    
    
    year_open = ws.Cells(i + 1, 3).Value
    current_ticker = current_ticker + 1
    
'Resetting the values
    Total_Stock_Volume = 0

    End If
 
Next i

    
'Finding maximum and minimum changes for the summary table
'Greatest % Increase->0, Greatest % Decrease->1, Greatest Volume->2
Dim SummaryVals(3) As Double
Dim SummaryTickers(3) As String

SummaryVals(0) = ws.Cells(2, 11)
SummaryVals(1) = ws.Cells(2, 11)
SummaryVals(2) = ws.Cells(2, 12)
SummaryTickers(0) = ws.Cells(2, 9)
SummaryTickers(1) = ws.Cells(2, 9)
SummaryTickers(2) = ws.Cells(2, 9)

LastRow_I = ws.Cells(Rows.Count, 9).End(xlUp).Row

For j = 2 To LastRow_I
    
    If ws.Cells(j, 11).Value > SummaryVals(0) Then
    SummaryVals(0) = ws.Cells(j, 11).Value
    SummaryTickers(0) = ws.Cells(j, 9).Value
    End If
    
    If ws.Cells(j, 11).Value < SummaryVals(1) Then
    SummaryVals(1) = ws.Cells(j, 11).Value
    SummaryTickers(1) = ws.Cells(j, 9).Value
    End If

    If ws.Cells(j, 12).Value > SummaryVals(2) Then
    SummaryVals(2) = ws.Cells(j, 12).Value
    SummaryTickers(2) = ws.Cells(j, 9).Value
    End If

Next j
    
ws.Cells(2, 16).Value = SummaryTickers(0)
ws.Cells(3, 16).Value = SummaryTickers(1)
ws.Cells(4, 16).Value = SummaryTickers(2)
ws.Cells(2, 17).Value = SummaryVals(0)
ws.Cells(3, 17).Value = SummaryVals(1)
ws.Cells(4, 17).Value = SummaryVals(2)

'Percent formatting
    ws.Cells(2, 17).Value = Format(SummaryVals(0), "Percent")
    ws.Cells(3, 17).Value = Format(SummaryVals(1), "Percent")


Next ws

End Sub
