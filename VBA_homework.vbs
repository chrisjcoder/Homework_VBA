Sub stock_market()



  ' Set an initial variable for holding the ticker name
  Dim ticker As String

  ' Set an initial variable for holding the total vol per ticker
  Dim vol_Total As Double
  vol_Total = 0

  ' Keep track of the location for each ticker name in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  Range("I1").Value = "Ticker"
  
  Range("J1").Value = "Total Stock Volume"

  ' Loop through all tickers
  For i = 2 To 760192

    ' Check if we are still within the same ticker, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the ticker name
      ticker = Cells(i, 1).Value

      ' Add to the volume Total
      vol_Total = vol_Total + Cells(i, 7).Value

      ' Print the ticker name in the Summary Table
      Range("I" & Summary_Table_Row).Value = ticker

      ' Print the volume Amount to the Summary Tab
      Range("J" & Summary_Table_Row).Value = vol_Total

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the volume Total
      vol_Total = 0

    ' If the cell immediately following a row is the same ticker name...
    Else

      ' Add to the volume Total
      vol_Total = vol_Total + Cells(i, 7).Value

    End If

  Next i

End Sub


