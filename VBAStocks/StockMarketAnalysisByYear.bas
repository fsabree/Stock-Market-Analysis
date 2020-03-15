Attribute VB_Name = "Module1"
' ## Instructions

' * Create a script that will loop through all the stocks for one year and output the following information.

' * The ticker symbol.

' * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

' * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

' * The total stock volume of the stock.

' * You should also have conditional formatting that will highlight positive change in green and negative change in red.

Sub StockMarketAnalysisByYear()

For Each ws In Worksheets

' Set variables
Dim Ticker As String
Dim YearOpen As Double
Dim YearlyChange As Double
Dim PercentChange As Double
Dim Volume As Double

Dim maxPercent As Double
Dim minPercent As Double
Dim BestTotalVolume As Double


lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
tableRow = 2
YearlyChange = 0
PercentChange = 0
Volume = 0
YearOpen = ws.Cells(2, 3).Value
maxPercent = 0
minPercent = 0
BestTotalVolume = 0

' Set headers for summary information

ws.Cells(1, 9).Value = "Ticker Symbol"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(2, 15).Value = "Greatest % increase"
ws.Cells(3, 15).Value = "Greatest % decrease"
ws.Cells(4, 15).Value = "Greatest total volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"



' Loop
For i = 2 To lastRow

                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            'Set Ticker Name
                ws.Range("I" & tableRow).Value = ws.Cells(i, 1).Value
                
            'Set Yearly Change
                
                YearlyChange = ws.Cells(i, 6).Value - YearOpen
                ws.Range("J" & tableRow).Value = YearlyChange
                       
            'Set Percent Change
                  
                 If (YearOpen = 0) Then
                      PercentChange = 0
                Else
                    PercentChange = YearlyChange / YearOpen
                
                End If
                ws.Range("K" & tableRow).Value = PercentChange
                ws.Range("K" & tableRow).Style = "Percent"
                
             'Set Total Stock Volume
                ws.Range("L" & tableRow).Value = Volume + ws.Cells(i, 7).Value
                
                'Conditional formatting of Yearly Change
                If (YearlyChange >= 0) Then
                    'Fill column with GREEN color for positive change - greater than 0
                    ws.Range("J" & tableRow).Interior.ColorIndex = 4
                ElseIf (YearlyChange < 0) Then
                    'Fill column with RED color for negative change - less than 0
                    ws.Range("J" & tableRow).Interior.ColorIndex = 3
                End If
                
                
               tableRow = tableRow + 1
               
               ' Reset the volume after each ticker value calculation.
               ' This is to make sure that volumes for each ticker are not added together.
               
               Volume = 0
               ' Next year open value
               YearOpen = ws.Cells(i + 1, 3)
               
             Else
             'Calculate total volume for each ticker symbol
             
               Volume = Volume + ws.Cells(i, 7).Value
               
               
        End If
    
    Next i
    
'### CHALLENGES

'Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume".
  
    
 lastSummRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
 
 'Loop
 
    For i = 2 To lastSummRow
            'Finds the Greatest % Increase
            
                If ws.Cells(i, 11).Value > maxPercent Then
                maxPercent = ws.Cells(i, 11).Value
                ws.Cells(2, 17) = maxPercent
                ws.Cells(2, 16) = ws.Cells(i, 9).Value
                ws.Cells(2, 17).Style = "Percent"
                
            'Finds the Greatest % Decrease
                ElseIf ws.Cells(i, 11).Value < minPercent Then
                minPercent = ws.Cells(i, 11).Value
                ws.Cells(3, 17) = minPercent
                ws.Cells(3, 16) = ws.Cells(i, 9).Value
                ws.Cells(3, 17).Style = "Percent"
            
                End If
            
            'Finds the Greatest Total Volume
                If ws.Cells(i, 12).Value > BestTotalVolume Then
                BestTotalVolume = ws.Cells(i, 12).Value
                ws.Cells(4, 17) = BestTotalVolume
                ws.Cells(4, 16) = ws.Cells(i, 9).Value
                            
                End If
                    
        Next i
    
    
     ' Autofit to display data
    ws.Columns("A:Q").AutoFit
          
      Next ws


End Sub

Sub ClearCells()

For Each ws In Worksheets
'Clear summary table and the best and worst stock percentages.

    ws.Columns("I:Q").Clear
      
      Next ws
End Sub

