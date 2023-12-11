Sub VBA_Challenge()

'Loop through all of the worksheets in the workbook

For Each ws In ThisWorkbook.Worksheets

' Label Headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
  ' Set initial variables
  Dim Ticker_Name As String
  Dim Total_Stock_Volume As Double
  Total_Stock_Volume = 0
  Dim Opening_Price As Double
  Dim Closing_Price As Double
  Dim Yearly_Change As Double
  Dim Percentage_Change As Double
  Dim Total_Ticker_Volume As Double
  Dim Max_IncreaseTicker_Name As String
  Dim Max_DecreaseTicker_Name As String
  Dim Max_Percentage As Double
  Max_Percentage = 0
  Dim Min_Percentage As Double
  Min_Percentage = 0
  Dim Max_Volume_Ticker_Name As String
  Dim Max_Volume As Double
  Max_Volume = 0

  ' Keep track of the location for each ticker name in the summary table
  Dim Summary_Row As Double
  Summary_Row = 2
  Dim OpenRow As Double
  OpenRow = 2
  
  Dim LastRow As Double
  LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Set opening price
Opening_Price = ws.Cells(OpenRow, 3).Value

  ' Loop through all ticker names
  For i = 2 To LastRow

    ' Check if we are still within the same ticker name, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the ticker name
      Ticker_Name = ws.Cells(i, 1).Value

'Set closing price
Closing_Price = ws.Cells(i, 6).Value

'Set yearly change
Yearly_Change = Closing_Price - Opening_Price

'Calculate
If Opening_Price <> 0 Then
Percentage_Change = (Yearly_Change / Opening_Price) * 100

End If

    'Add to the stock volume total
    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

      ' Print the ticker name in the summary row
      ws.Range("I" & Summary_Row).Value = Ticker_Name
      
      'Print the yearly change in the Summary Row, Column J
ws.Range("J" & Summary_Row).Value = Yearly_Change

'Use conditional formatting that will highlight positive change in green and negative change in red.
If (Yearly_Change > 0) Then
    ws.Range("J" & Summary_Row).Interior.ColorIndex = 4
    
    ElseIf (Yearly_Change < 0) Then
    
        ws.Range("J" & Summary_Row).Interior.ColorIndex = 3

End If

'Print the percentage change in the Summary Table,Column K
ws.Range("K" & Summary_Row).Value = (CStr(Percentage_Change) & "%")
    
       'Print the total stock volume in the summary row
    ws.Range("L" & Summary_Row).Value = Total_Stock_Volume

      ' Add one to the summary row
      Summary_Row = Summary_Row + 1
      
      'Set the opening price
Opening_Price = ws.Cells(i + 1, 3).Value

'Calculate
If (Percentage_Change > Max_Percentage) Then
        Max_Percentage = Percentage_Change
        Max_Increase_Ticker_Name = Ticker_Name

ElseIf (Percentage_Change < Min_Percentage) Then
       Min_Percentage = Percentage_Change
       Min_Decrease_Ticker_Name = Ticker_Name
       
    End If
    
 If (Total_Stock_Volume > Max_Volume) Then
    Max_Volume = Total_Stock_Volume
    Max_Volume_Ticker_Name = Ticker_Name

End If

        'Reset the stock volume total
      Total_Stock_Volume = 0
      
      'Reset the percent change
      Percentage_Change = 0
      
      ' If the cell immediately following a row is the same ticker name...
    Else

      ' Add to the stock volume total
      Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

End If

OpenRow = i + 1

Next i

'Print values in assigned cells
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("Q2").Value = (CStr(Max_Percentage) & "%")
ws.Range("Q3").Value = (CStr(Min_Percentage) & "%")
ws.Range("Q4").Value = Max_Volume
ws.Range("P2").Value = Max_Increase_Ticker_Name
ws.Range("P3").Value = Min_Decrease_Ticker_Name
ws.Range("P4").Value = Max_Volume_Ticker_Name
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

Next ws

End Sub


