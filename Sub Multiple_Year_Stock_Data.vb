Sub Multiple_Year_Stock_Data()
MsgBox "STOCKS!!!"

Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
ws.Activate
    
'Find last cell
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Set Header Columns
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"

    'Set up Header Rows for Bonus
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
     
'Set Variables
Dim Ticker As String
Dim Ticker_ As Integer
Dim Opening_Price As Double
Dim Closing_Price As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Total_Stock_Volume As Double

'Set Bonus Variables
Dim Greatest_Percent_Increase As Double
Dim Greatest_Percent_Increase_Ticker As String
Dim Greatest_Percent_Decrease As Double
Dim Greatest_Percent_Decrease_Ticker As String
Dim Greatest_Stock_Volume As Double
Dim Greatest_Stock_Volume_Ticker As String

    Ticker_Numbers = 0
    Ticker = ""
    Yearly_Change = 0
    Opening_Price = 0
    Percent_Change = 0
    Total_Stock_Volume = 0
    
    'Set Loop that will run
    For i = 2 To LastRow
    Ticker = ws.Cells(i, 1).Value
        
    If Opening_Price = 0 Then
    Opening_Price = ws.Cells(i, 3).Value
    End If
        
    'Find the Ticker Volume
    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        
    'Find the Next Ticker
    If ws.Cells(i + 1, 1).Value <> Ticker Then
    Ticker_Numbers = Ticker_Numbers + 1
    ws.Cells(Ticker_Numbers + 1, 9) = Ticker
            
    Closing_Price = ws.Cells(i, 6)
    Yearly_Change = Closing_Price - Opening_Price
    ws.Cells(Ticker_Numbers + 1, 10).Value = Yearly_Change
            
        If Opening_Price = 0 Then
        Percent_Change = 0
        Else
        Percent_Change = (Yearly_Change / Opening_Price)
        End If
        Opening_Price = 0
            
        ws.Cells(Ticker_Numbers + 1, 12).Value = Total_Stock_Volume
        Total_Stock_Volume = 0
             
        If Yearly_Change >= 0 Then
        ws.Cells(Ticker_Numbers + 1, 10).Interior.ColorIndex = 4
        Else
        ws.Cells(Ticker_Numbers + 1, 10).Interior.ColorIndex = 3
        End If
            
        ws.Cells(Ticker_Numbers + 1, 11).Value = Format(Percent_Change, "Percent")
        End If
        
    Next i

'Bonus
LastRow = ws.Cells(Rows.Count, "I").End(xlUp).Row
Greatest_Percent_Increase = ws.Cells(2, 11).Value
Greatest_Percent_Increase_Ticker = ws.Cells(2, 9).Value
Greatest_Percent_Decrease = ws.Cells(2, 11).Value
Greatest_Percent_Decrease_Ticker = ws.Cells(2, 9).Value
Greatest_Stock_Volume = ws.Cells(2, 12).Value
Greatest_Stock_Volume_Ticker = ws.Cells(2, 9).Value

'Proceed to Loop Thru
    For i = 2 To LastRow
    If Cells(i, 11).Value > Greatest_Percent_Increase Then
        Greatest_Percent_Increase = ws.Cells(i, 11).Value
        Greatest_Percent_Increase_Ticker = ws.Cells(i, 9).Value
        End If
        
    If Cells(i, 11).Value < Greatest_Percent_Decrease Then
        Greatest_Percent_Decrease = ws.Cells(i, 11).Value
        Greatest_Percent_Decrease_Ticker = ws.Cells(i, 9).Value
        End If
        
    If Cells(i, 12).Value > Greatest_Stock_Volume Then
        Greatest_Stock_Volume = ws.Cells(i, 12).Value
        Greatest_Stock_Volume_Ticker = ws.Cells(i, 9).Value
        End If

'Set Format
ws.Range("P2").Value = Format(Greatest_Percent_Increase_Ticker, "Percent")
ws.Range("Q2").Value = Format(Greatest_Percent_Increase, "Percent")
ws.Range("P3").Value = Format(Greatest_Percent_Decrease_Ticker, "Percent")
ws.Range("Q3").Value = Format(Greatest_Percent_Decrease, "Percent")
ws.Range("P4").Value = Greatest_Stock_Volume_Ticker
ws.Range("Q4").Value = Greatest_Stock_Volume

Next i

Next ws
   
End Sub
