VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub wallstreet()

Dim ws As Worksheet
'to select all worksheets
For Each ws In ThisWorkbook.Worksheets
    If ws.Visible = xlSheetVisible Then
        ws.Select False
    End If


Dim lastrow As Long
Dim Tickername As String
Dim Tickerrow As Long
Tickerrow = 2
Dim previousprice As Long
previousprice = 2
Dim Totalvolume As Double
 Totalvolume = 0
Dim Percentchange As Double
Dim yearlychange As Double
Dim yearlyopen As Double
Dim yearlyclose As Double
Dim mmax As Double
Dim mmin As Double
Dim mvolmax As Double

 lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 ws.Cells(1, 9).Value = "Ticker"
 ws.Cells(1, 10).Value = "Yearly Change"
 ws.Cells(1, 11).Value = "Percent Change"
 ws.Cells(1, 12).Value = "Total Stock Volume"
 ws.Cells(1, 16).Value = "Ticker"
 ws.Cells(1, 17).Value = "Value"
 ws.Cells(2, 15).Value = "Greatest % Increase"
 ws.Cells(3, 15).Value = "Greateset % Decrease"
 ws.Cells(4, 15).Value = "Greatest Total Volume"
 
For i = 2 To lastrow
  
  'to calculate the total volume
  Totalvolume = Totalvolume + ws.Cells(i, 7).Value
 
 'to find the next different Tickername
 If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
 Tickername = ws.Cells(i, 1).Value
 ws.Range("I" & Tickerrow).Value = Tickername
 
 'to calculate the yearlychange
 yearlyopen = ws.Range("C" & previousprice)
 yearlyclose = ws.Range("F" & i)
 yearlychange = yearlyclose - yearlyopen
 ws.Range("J" & Tickerrow) = yearlychange
 
 'to calculate the percentage change
  yearlyopen = ws.Range("C" & previousprice)
 Percentchange = yearlychange / yearlyopen
 ws.Range("K" & Tickerrow).Value = Percentchange
  ws.Range("K" & Tickerrow).NumberFormat = "0.00%"
  'input the value of totalvolume for each ticker
  ws.Range("L" & Tickerrow).Value = Totalvolume
 'reset Totalvolume for next "Ticker Symbol"
 
 Totalvolume = 0
 
 ' to highlight prositive change in green and negative change in red
 If yearlychange >= 0 Then
 ws.Range("K" & Tickerrow).Interior.ColorIndex = 4
 Else
 ws.Range("K" & Tickerrow).Interior.ColorIndex = 3
 End If
 Tickerrow = Tickerrow + 1
 
 'Move to next "Ticker Symbol" to find the value of yearlyopen
 previousprice = i + 1
 
 End If
 
 Next i

 
 mmax = ws.Cells(2, 11).Value
 mmax_row = 2
 mmin = ws.Cells(2, 11).Value
 mmin_row = 2
 mvolmax = ws.Cells(2, 12).Value
 mvolmax_row = 2
 
 lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
 For i = 2 To lastrow
 'to calculate the greatest % increase and find the Ticker
 If ws.Cells(i, 11).Value > mmax Then
   mmax = ws.Cells(i, 11).Value
   mmax_row = i
   End If
   
   'to calculate the greatest % decrease and find the Ticker
   If ws.Cells(i, 11).Value < mmin Then
   mmin = ws.Cells(i, 11).Value
   mmin_row = i
   
 End If
 'to calculate the greatest total volume and find the Ticker
 
 If ws.Cells(i, 12).Value > mvolmax Then
 mvolmax = ws.Cells(i, 12).Value
 mvolmax_row = i
 End If

 
Next i

   ws.Cells(2, 17).Value = mmax
   ws.Cells(2, 16).Value = ws.Cells(mmax_row, 9).Value
   ws.Cells(3, 17).Value = mmin
   ws.Cells(3, 16).Value = ws.Cells(mmin_row, 9).Value
   ws.Cells(4, 17).Value = mvolmax
   ws.Cells(4, 16).Value = ws.Cells(mvolmax_row, 9).Value

'Format the values to a percentage
 ws.Cells(2, 17).NumberFormat = "0.00%"
 ws.Cells(3, 17).NumberFormat = "0.00%"

 'move to next sheet
 Next ws
 
 End Sub



