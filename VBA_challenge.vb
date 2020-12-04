Sub stock_market():
    ' set the variables
    Dim ws As Worksheet
    Dim last_row As Long
    Dim start As Long
    Dim openval As Double
    Dim closeval As Double
    Dim ticker As String
    Dim totalvol As Double
    
    'start loop through all worksheets
    For Each ws In Worksheets
    
    'select worksheet and go to the next worksheet
    Worksheets(ws.Name).Select

    ' set the titles
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percentage Changes"
    Range("L1").Value = "Total Stock Volume"
    
    'making the titles bold
    Range("I1").Font.Bold = True
    Range("J1").Font.Bold = True
    Range("K1").Font.Bold = True
    Range("L1").Font.Bold = True

    'define the last row
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Define variables
        start = 2
        nr = 0
    'for loop from A2 till the last row
    For Row = 2 To last_row:
    
    totalvol = totalvol + Cells(Row, 7).Value
    
    'detect the change in ticker
    If Cells(Row, 1).Value <> Cells(Row + 1, 1).Value Then
    'this if statement looks at the current row and the next row.
    'the if statement will only run if the next row is different to the current row.
    
    'this will only be true once it reaches A262
    ticker = Cells(start, 1).Value
    'this is how you save the value
    Range("I" & 2 + nr).Value = ticker
    
    'get the open val
    'opening value starts in a2
    openval = Cells(start, 3).Value
    
    'get the close val
    closeval = Cells(Row, 6).Value
    
    'total volume
    'it adds for every row
    'then takes a break when the next ticker changes
    Range("L" & 2 + nr).Value = totalvol
    'then starts adding again from 0
    totalvol = 0
    
    'yearlychange = closeval - openval
    yearly_change = closeval - openval
    
    If openval <> 0 Then
        
    'calculate the percent change
    percentage_changes = (closeval - openval) / openval
        
    Else
        percentage_changes = 0
        
    End If
        
    'putting the yearly change
    Range("J" & 2 + nr).Value = yearly_change
    
    'set the color so that positive is green and negative is red
    If yearly_change > 0 Then
    
    'set the color to green
    Range("J" & 2 + nr).Interior.ColorIndex = 4
    
    Else
    
    'set the color to red
    Range("J" & 2 + nr).Interior.ColorIndex = 3
    
    End If
    
    'putting the percentage changes
    Range("K" & 2 + nr).Value = percentage_changes
    'changing the format to percentage
    Range("K" & 2 + nr).NumberFormat = "0.00%"
    
    start = Row + 1
    
    nr = nr + 1
    
    End If
    
    Next Row
    
    Next ws

End Sub