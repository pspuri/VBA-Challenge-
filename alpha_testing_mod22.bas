Attribute VB_Name = "Module22"
Sub alphalisting()
 
'For use in each worksheet
' Declare Current as a worksheet object variable.
         'Dim Current As Worksheet
   ' Loop through all of the worksheets in the active workbook.
         'For Each Current In Worksheets

'Automating to find the last row
Dim LastRow As Double

'define last row in column 1
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
 

 'set variable for ticker symbol
Dim Ticker As String

'set variable for total volume
Dim Vol As Double

Vol = 0



'set variable for opening price
Dim Start_price As Double


'set variable for closing price
Dim End_Price As Double

'set variable for yearly change
Dim Yearly_Change As Double

'set variable for yearly percent change
Dim Change_Percent As Double



'Keep track of the location forTicker in the summary table
Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2


'loop through data for the tickers

For I = 2 To LastRow

'Set Start Price
    If Cells(I, 2).Value = 20150101 Then
    
    Start_price = Cells(I, 6).Value
    
    'Print start price
    Range("Q" & Summary_Table_Row).Value = Start_price


    End If
    
    'set yearly ending close price
    If Cells(I, 2).Value = 20151230 Then
    
    End_Price = Cells(I, 6).Value
    
    'Print end price
    Range("R" & Summary_Table_Row).Value = End_Price
        
    End If

'calculate difference per year
Yearly_Change = End_Price - Start_price

'print change
Range("N" & Summary_Table_Row).Value = Yearly_Change
 
    
'Check if we are still within the Ticker, if it is not...
    If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then

        'set Ticker Symbol
        Ticker = Cells(I, 1).Value

        'set Ticker Volume
        Vol = Vol + Cells(I, 7).Value


        'Print Ticker Symbol
        Range("M" & Summary_Table_Row).Value = Ticker

        'Print Ticker Volume
       Range("P" & Summary_Table_Row).Value = Vol
       
'calc yearly percent change
     Change_Percent = Yearly_Change / End_Price
     
     'print percent change
     Range("O" & Summary_Table_Row).Value = Change_Percent
    

        'increment summary table for next symbol
        Summary_Table_Row = Summary_Table_Row + 1

        'reset Total Volume
        Vol = 0
        
     

Else

    'Add to the Ticker Volume
    Vol = Vol + Cells(I, 7)
    
    
End If
        

Next I

'define last row for column 15 percent change
Dim LastRow1 As Double

'set highest percentage gain
Dim Greatest_Percent As Double

'set Last Row1
LastRow1 = Cells(Rows.Count, 15).End(xlUp).Row

'connditional to change color of cell if greater than or equal to zero

For j = 2 To LastRow1

If Cells(j, 15) >= 0 Then

Cells(j, 15).Interior.ColorIndex = 4

Else

Cells(j, 15).Interior.ColorIndex = 3

End If

'find max and min values and associated cells

Cells(2, 22).Value = Application.WorksheetFunction.Max(Columns("O"))

Cells(3, 22).Value = Application.WorksheetFunction.Min(Columns("O"))

'Find most total volume

Cells(4, 22).Value = Application.WorksheetFunction.Max(Columns("P"))

'Find the corresponding ticker for max,min, and greatest volume

'Cells(2, 21).Value = Application.WorksheetFunction.VLookup(Cells(4, 22).Value, Range("M2:N300"), 1, False)

If Cells(j, 15) = Cells(2, 22).Value Then
Cells(2, 21).Value = Cells(j, 13).Value
End If

If Cells(j, 15) = Cells(3, 22).Value Then
Cells(3, 21).Value = Cells(j, 13).Value
End If

If Cells(j, 16) = Cells(4, 22).Value Then
Cells(4, 21).Value = Cells(j, 13).Value
End If

Next j

 
'Next


End Sub
