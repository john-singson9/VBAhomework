Attribute VB_Name = "RibbonX_Code"
'Entry point for RibbonX button click
Sub ShowATPDialog(control As IRibbonControl)
    Application.Run ("fDialog")
End Sub

'Callback for RibbonX button label
Sub GetATPLabel(control As IRibbonControl, ByRef label)
    label = ThisWorkbook.Sheets("RES").Range("A10").Value
End Sub

Sub Stock_Loop()
' Create a loop through each year of stock data and grab the total amount of volume
' Display the ticker symbol to coincide with the total volume

' loop through all sheets
For Each ws In Worksheets

' Set variables for total stock volume, ticker symbols and Summary Table Row
Dim Summary_Table_Row As Integer
Dim Total_Stock As Double
Dim ticker As String

Total_Stock = 0
Summary_Table_Row = 2

' Determine the last row
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

' Loop through all the stocks
    For i = 2 To lastrow

    ' Check if we are still within the same Stock
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                ' Set the Ticker Symbol
                ticker = ws.Cells(i, 1).Value
                
                'Add to the Stock Total
                Total_Stock = Total_Stock + ws.Cells(i, 7).Value
                
                'Put the Stock on a Summary Table
                ws.Range("I" & Summary_Table_Row).Value = ticker
                
                'Put the Stock Total on a Summary Table
                ws.Range("L" & Summary_Table_Row).Value = Total_Stock
                
                'Add to the Summary Table Row
                Summary_Table_Row = Summary_Table_Row + 1
                
                'Reset the Total Stock
                Total_Stock = 0
                
                ' If they are both the same stock
        Else
        
                ' Add to the Stock Total
                Total_Stock = Total_Stock + ws.Cells(i, 7).Value
        
        End If

    Next i
    
Next ws

End Sub

Sub Moderate()
' Set variables for opening price, ending price, yearly change, percent change, and summary table row
Dim opening_price As Double
Dim closing_price As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim Summary_Table_Row As Double

Summary_Table_Row = 2

'Determine the last row
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

' Loop through all stocks
    For i = 2 To lastrow
        'Check for the same stock and opening price
        If Cells(i, 1).Value <> Cells(i - 1, 1).Value And Cells(i, 3).Value > 0.5 Then
            opening_price = Cells(i, 3).Value
        'Check if it's still the same stock and ending price
       Else
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            closing_price = Cells(i, 6).Value
            yearly_change = closing_price - opening_price
            Range("J" & Summary_Table_Row).Value = yearly_change
            percent_change = yearly_change / opening_price
            Range("K" & Summary_Table_Row).Value = percent_change
            Summary_Table_Row = Summary_Table_Row + 1
            Else
            End If
        End If
    Next i

End Sub

Sub Hard()

' loop through all worksheet
For Each ws In Worksheets

' declaring variables
Dim i
Dim Summary_Table As Double
Dim greater_volume As Double
Dim greater_percentage As Double
Dim least_percentage As Double
Dim ticker As String
Dim ticker2 As String
Dim ticker3 As String

lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
Summary_Table = 2
greater_volume = 0
greater_percentage = 0
least_percentage = 0


' greatest value
For i = 2 To lastrow
If ws.Cells(i, 12).Value > greater_volume Then
    greater_volume = ws.Cells(i, 12).Value
    ticker = ws.Cells(i, 9).Value
End If

' greatest percentage
If ws.Cells(i, 11).Value > greater_percentage Then
    greater_percentage = ws.Cells(i, 11).Value
    ticker2 = ws.Cells(i, 9).Value
End If

' least percentage
If ws.Cells(i, 11).Value < least_percentage Then
    least_percentage = ws.Cells(i, 11).Value
    ticker3 = ws.Cells(i, 9).Value
End If
Next i

' putting the greatest value on the table
ws.Range("p" & Summary_Table).Value = greater_volume
ws.Range("q" & Summary_Table).Value = ticker
Summary_Table = Summary_Table + 1

' putting the greatest percentage on the table
ws.Range("p" & Summary_Table).Value = greater_percentage
ws.Range("q" & Summary_Table).Value = ticker2
Summary_Table = Summary_Table + 1

' putting the least percentage on the table
ws.Range("p" & Summary_Table).Value = least_percentage
ws.Range("q" & Summary_Table).Value = ticker3

Next ws
End Sub

Sub yearlycolor()

' loop through all worksheets
For Each ws In Worksheets

lastrow = ws.Cells(Rows.Count, 10).End(xlUp).Row

' change colors on column j:  red for negative, green for positive
For i = 2 To lastrow
    If ws.Cells(i, 10).Value >= 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
    Else
        ws.Cells(i, 10).Interior.ColorIndex = 3
    End If
Next i
Next ws
End Sub
