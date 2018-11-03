Sub amountagg()
For Each ws In Worksheets

  Dim Ticker_Name As String


  Dim Ticker_amount As Double
  Ticker_amount = 0
  lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

  
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Total Stock Volume"

  For i = 2 To lastrow

    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      
      Ticker_Name = ws.Cells(i, 1).Value

      
      Ticker_amount = Ticker_amount + ws.Cells(i, 7).Value

   
      ws.Range("I" & Summary_Table_Row).Value = Ticker_Name

      ws.Range("J" & Summary_Table_Row).Value = Ticker_amount

      Summary_Table_Row = Summary_Table_Row + 1
      
     
      Ticker_amount = 0

    
    Else

     
      Ticker_amount = Ticker_amount + ws.Cells(i, 7).Value

    End If

  Next i
Next ws
End Sub
