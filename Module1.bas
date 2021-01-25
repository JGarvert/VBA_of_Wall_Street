Attribute VB_Name = "Module1"
Sub RunAll():

Dim ws As Worksheet
Application.ScreenUpdating = False
For Each ws In Worksheets
   ws.Select
   Call vbaWallStreet
   Next
   Application.ScreenUpdating = True
End Sub

Sub vbaWallStreet():

   ' add headers for summary
   Cells(1, 9).Value = "Ticker"
   Cells(1, 10).Value = "Yearly Change"
   Cells(1, 11).Value = "Percent Change"
   Cells(1, 12).Value = "Total Stock Value"
   Cells(1, 16).Value = "Ticker"
   Cells(1, 17).Value = "Value"
   Cells(2, 15).Value = "Greatest % Increase"
   Cells(3, 15).Value = "Greatest % Decrease"
   Cells(4, 15).Value = "Greatest Total Volume"

 
   ' Make listing of all unique tickers
   ' Set variables to hold ticker, beginning, and ending prices.  And total volumn.
   Dim ticker As String
   Dim ticker_beg As Double
   Dim ticker_end As Double
   ticker_beg = Range("C2").Value
   ticker_end = 0
   ticker_vol = Range("G2")
 
   ' Keep track of each ticker in a summary table
   Dim Summ_row As Integer
   Summ_row = 2
 
   ' Set rowcount to run down all rows with data
   endrow = Cells(Rows.Count, 1).End(xlUp).Row

   For i = 2 To endrow
    ' add up ticker volumn until the ticker is different
    ticker_vol = ticker_vol + Cells(i, 7)
    Range("L" & Summ_row) = ticker_vol
      
       ' see if after loop ticker has changed
       If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
       ' set ticker name and beginning value
      ticker = Cells(i, 1).Value
      ticker_end = Cells(i, 6).Value
      ticker_vol = Cells(i + 1, 7).Value
   
      ' print ticker name in summary table and beginning value
      Range("I" & Summ_row).Value = ticker
      Range("J" & Summ_row).Interior.ColorIndex = 3
      Range("J" & Summ_row) = ticker_end - ticker_beg
      ' handle division by zero
         If ticker_beg <> 0 Then
         Range("K" & Summ_row) = Format((ticker_end - ticker_beg) / ticker_beg, "Percent")
         End If
      ' Set conditional formatting
         If (ticker_end - ticker_beg) >= 0 Then
         Range("J" & Summ_row).Interior.ColorIndex = 4
         End If
   
    ' reset beginning ticker value
    ticker_beg = Cells(i + 1, 3).Value
    ' add one to summary row
      Summ_row = Summ_row + 1

   End If
Next i

' populate summary cells
Max_pct_increase = Application.WorksheetFunction.Max(Columns("K"))
Min_pct_decrease = Application.WorksheetFunction.Min(Columns("K"))
Max_volume = Application.WorksheetFunction.Max(Columns("L"))

Cells(2, 17).Value = Format(Max_pct_increase, "Percent")
Cells(3, 17).Value = Format(Min_pct_decrease, "Percent")
Cells(4, 17).Value = Max_volume

Max_pct_ticker_row = Application.WorksheetFunction.Match(Max_pct_increase, Columns("K"), 0)
Max_pct_ticker = Cells(Max_pct_ticker_row, 9)
Cells(2, 16) = Max_pct_ticker

Min_pct_ticker_row = Application.WorksheetFunction.Match(Min_pct_decrease, Columns("K"), 0)
Min_pct_ticker = Cells(Min_pct_ticker_row, 9)
Cells(3, 16) = Min_pct_ticker

Max_vol_ticker_row = Application.WorksheetFunction.Match(Max_volume, Columns("L"), 0)
Max_vol_ticker = Cells(Max_vol_ticker_row, 9)
Cells(4, 16) = Max_vol_ticker

End Sub
