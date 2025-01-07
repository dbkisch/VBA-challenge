# VBA-challenge
Sub stock_analyzer()
  ' Create a variable to hold the counter
Dim ws As Worksheet
For Each ws In Worksheets
  Dim i As Long
  Dim LastRow As Long
  Dim LastSummaryRow As Long
  Dim j As Long
  
  ' Set an initial variable for holding the ticker and prices
  Dim ticker As String
  Dim openprice As Double
  Dim closeprice As Double
  Dim MaxIncrease As Double
  Dim MaxDecrease As Double
  Dim MaxVolume As Double

  ' Set an initial variable for holding the total per ticker
  Dim totalvolume As Double
  totalvolume = 0
  MaxIncrease = 0
  MaxDecrease = 0
  MaxVolume = 0
  
  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Long
  Summary_Table_Row = 2

  LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
   
        ' Create new column headers in summary section
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
     '   ws.Range("N1").Value = "Open Price"
     '   ws.Range("O1").Value = "Close Price"
        
        ws.Range("O2").Value = "Greatest % Increase"
       ws.Range("O3").Value = "Greatest % Decrease"
       ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

           
  ' Loop through all stock prices

 'initialize openprice
 openprice = ws.Cells(2, 3).Value
 '    MsgBox (openprice)
     
  For i = 2 To LastRow
    
    ' Check if we have reached the last row for the ticker...
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      'grab ticker and close price
      ticker = ws.Cells(i, 1).Value
      closeprice = ws.Cells(i, 6).Value
      
      ' Add to the totalvolume
      totalvolume = totalvolume + ws.Cells(i, 7).Value

      ' Print the ticker in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = ticker

      ' Format columns in the Summary Table
 
      ws.Range("J" & Summary_Table_Row).Value = closeprice - openprice
            If (closeprice - openprice >= 0) Then
                  ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
            Else
                  ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
            End If
      ws.Range("J" & Summary_Table_Row).NumberFormat = "0.00"
    
      ws.Range("K" & Summary_Table_Row).Value = (closeprice - openprice) / openprice
            If ((closeprice - openprice) / openprice >= 0) Then
                  ws.Cells(Summary_Table_Row, 11).Interior.ColorIndex = 4
            Else
                  ws.Cells(Summary_Table_Row, 11).Interior.ColorIndex = 3
            End If
      ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
      
      ' Print the totalvolume to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = totalvolume
      
     ' ws.Range("N" & Summary_Table_Row).Value = openprice
     ' ws.Range("O" & Summary_Table_Row).Value = closeprice

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1

      ' Reset the totalvolume and openprice
      totalvolume = 0
      openprice = ws.Cells(i + 1, 3).Value

    ' If the cell immediately following a row is the same ticker..
    Else

      ' Add to the totalvolume
      totalvolume = totalvolume + ws.Cells(i, 7).Value
      
    End If

  Next i
'-------------------------------------------------------------
' Finally, loop through the summary table to determine the greatest increase, decrease and volume,
' and print it in the upper right corner
  
  
  LastSummaryRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
  'MsgBox (LastSummaryRow)
    MaxIncrease = Application.WorksheetFunction.Max _
    (Range(ws.Cells(2, 11), ws.Cells(LastSummaryRow, 11)))
    ws.Range("Q2").Value = MaxIncrease
  
    MaxDecrease = Application.WorksheetFunction.Min _
    (Range(ws.Cells(2, 11), ws.Cells(LastSummaryRow, 11)))
    ws.Range("Q3").Value = MaxDecrease
    ws.Range("Q3").NumberFormat = "0.00%"

    MaxVolume = Application.WorksheetFunction.Max _
    (Range(ws.Cells(2, 12), ws.Cells(LastSummaryRow, 12)))
    ws.Range("Q4").Value = MaxVolume
    
    For j = 2 To LastSummaryRow
        If ws.Cells(j, 11).Value = MaxIncrease Then
            ws.Range("P2").Value = ws.Cells(j, 9).Value
        ElseIf ws.Cells(j, 11).Value = MaxDecrease Then
            ws.Range("P3").Value = ws.Cells(j, 9).Value
        ElseIf ws.Cells(j, 12).Value = MaxVolume Then
            ws.Range("P4").Value = ws.Cells(j, 9).Value
        End If
    Next j
    
    
Next ws

End Sub
