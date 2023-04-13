Attribute VB_Name = "Module1"
Sub Stock_market()

'Declare and set worksheet
For Each ws In Worksheets
    
        Dim WorksheetName As String


'Create the column headings
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"


'Set initial and last row for worksheet
Dim LastRowA As Long
'For the current row
Dim i As Long
'Starting row of the ticker part
Dim j As Long


'Define Lastrow of worksheet Column A
LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Set new variables for prices and percent changes
Dim open_price As Double
open_price = 0
Dim close_price As Double
close_price = 0
Dim price_change As Double
price_change = 0

Dim TickerRow As Long
TickerRow = 2

'Setting the start row,another variable for the starting row
j = 2

'Do loop of current worksheet to Lastrow
For i = 2 To LastRowA

'Ticker symbol output

'Check if ticker name changed
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Write ticker in column I
                ws.Cells(TickerRow, 9).Value = ws.Cells(i, 1).Value

  

'Calculate change in Price
close_price = ws.Cells(i, 6).Value
open_price = ws.Cells(j, 3).Value
price_change = close_price - open_price
ws.Cells(TickerRow, 10) = price_change

   'For the color formatting
   If price_change > 0 Then
   ws.Cells(TickerRow, 10).Interior.ColorIndex = 4
   Else
    ws.Cells(TickerRow, 10).Interior.ColorIndex = 3
    End If
    
   'Calculate the percentage change
       If ws.Cells(j, 3).Value <> 0 Then
                    PerChange = price_change / open_price
                    
                    'Percent formating
                    ws.Cells(TickerRow, 11).Value = Format(PerChange, "Percent")
                    
                    Else
                    
                    ws.Cells(TickerRow, 11).Value = Format(0, "Percent")
                    
                    
        End If
        
        
        'Find the total volume in Column L
        ws.Cells(TickerRow, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
        
        
        
                
             
                    TickerRow = TickerRow + 1


j = i + 1

End If

Next i



'Find last non-blank cell in column I
        LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        'Getting ready for summary
        GreatVol = ws.Cells(2, 12).Value
        GreatIncr = ws.Cells(2, 11).Value
        GreatDecr = ws.Cells(2, 11).Value
        
            'Loop for summary
            For i = 2 To LastRowI
            
                'For greatest total volume,check if next value is greater,if it's true then take over a new value and populate ws.Cells
                If ws.Cells(i, 12).Value > GreatVol Then
                GreatVol = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatVol = GreatVol
                
                End If
                
                'For greatest increase,check if next value is greater,if it's true then take over a new value and populate ws.Cells
                
                If ws.Cells(i, 11).Value > GreatIncr Then
                GreatIncr = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatIncr = GreatIncr
                
                End If
                
                'For greatest decrease,check if next value is smaller, if yes take over a new value and populate ws.Cells
                If ws.Cells(i, 11).Value < GreatDecr Then
                GreatDecr = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatDecr = GreatDecr
                
                End If
                
            'Write summary results in ws.Cells
            ws.Cells(2, 17).Value = Format(GreatIncr, "Percent")
            ws.Cells(3, 17).Value = Format(GreatDecr, "Percent")
            ws.Cells(4, 17).Value = Format(GreatVol, "Scientific")
            
            Next i
            



Next ws

End Sub

