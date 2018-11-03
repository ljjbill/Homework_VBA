Sub Moderate()
For Each ws In Worksheets

  Dim Ticker_Name As String

Dim Open_amount As Double
Dim Close_amount As Double
Dim Yearly_change As Double
Dim Percentage_change As Double
  Dim Ticker_amount As Double
  Ticker_amount = 0
  lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
ws.Cells(lastrow + 1, 1).Value = "DUMMY"
  ws.Cells(lastrow + 1, 3).Value = 100
  ws.Cells(lastrow + 1, 6).Value = 120
  
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
ws.Cells(1, 9).Value = 10
ws.Cells(1, 10).Value = 11
ws.Cells(1, 11).Value = 12
ws.Cells(1, 12).Value = 13
ws.Cells(1, 13).Value = 15
ws.Cells(1, 6).Value = 14
Set rg = Range("J2", Range("J2").End(xlDown))
Set cond1 = rg.FormatConditions.Add(xlCellValue, xlGreater, "=0")
Set cond2 = rg.FormatConditions.Add(xlCellValue, xlLess, "=0")

  For i = 2 To lastrow + 1
  

    
    if ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then

      
      Ticker_Name = ws.Cells(i, 1).Value

      
    Open_amount = ws.Cells(i, 3).Value

   
      ws.Range("I" & Summary_Table_Row).Value = Ticker_Name

     
      
      ws.Range("M" & Summary_Table_Row).Value = Open_amount
      

      
      
     
      Ticker_amount = ws.Cells(i, 7).Value
      
      Close_amount = ws.Cells(i - 1, 6).Value
      ws.Range("N" & Summary_Table_Row - 1).Value = Close_amount
      Yearly_change = ws.Cells(Summary_Table_Row - 1, 14).Value - ws.Cells(Summary_Table_Row - 1, 13).Value
      ws.Range("J" & Summary_Table_Row - 1).Value = Yearly_change
      Percentage_change = Yearly_change / ws.Cells(Summary_Table_Row - 1, 13).Value
      Format_percentage = Format(Percentage_change, "Percent")
      ws.Range("K" & Summary_Table_Row - 1).Value = Format_percentage
      
Summary_Table_Row = Summary_Table_Row + 1
    
    Else

     
      'Close_amount = ws.Cells(i + 1, 6).Value
      'ws.Range("N" & Summary_Table_Row - 1).Value = Close_amount
      'Yearly_change = ws.Cells(Summary_Table_Row - 1, 14).Value - ws.Cells(Summary_Table_Row - 1, 13).Value
      'ws.Range("J" & Summary_Table_Row - 1).Value = Yearly_change
      'Percentage_change = Yearly_change / ws.Cells(Summary_Table_Row - 1, 13).Value
      'Format_percentage = Format(Percentage_change, "Percent")
      'ws.Range("K" & Summary_Table_Row - 1).Value = Format_percentage
      Ticker_amount = Ticker_amount + ws.Cells(i, 7).Value
      ws.Range("L" & Summary_Table_Row - 1).Value = Ticker_amount

    End If

  Next i
  ws.Cells(1, 6).Value = "<close>"
  ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
  ws.Columns("M").EntireColumn.Delete
  ws.Columns("M").EntireColumn.Delete
  lastrow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
  ws.Cells(lastrow2, 9).Value = ""
  ws.Cells(lastrow2, 12).Value = ""
  ws.Cells(lastrow2, 13).Value = ""
  ws.Cells(1, 13).Value = ""
  ws.Cells(1, 14).Value = ""
  ws.Cells(lastrow + 1, 1).Value = ""
  ws.Cells(lastrow + 1, 3).Value = ""
  ws.Cells(lastrow + 1, 6).Value = ""
With cond1
.Interior.Color = vbGreen
End With
 
With cond2
.Interior.Color = vbRed
End With

Next ws

sheets("2015").activate

Set rg = Range("J2", Range("J2").End(xlDown))
Set cond1 = rg.FormatConditions.Add(xlCellValue, xlGreater, "=0")
Set cond2 = rg.FormatConditions.Add(xlCellValue, xlLess, "=0")

With cond1
.Interior.Color = vbGreen
End With
 
With cond2
.Interior.Color = vbRed
End With

sheets("2014").activate

Set rg = Range("J2", Range("J2").End(xlDown))
Set cond1 = rg.FormatConditions.Add(xlCellValue, xlGreater, "=0")
Set cond2 = rg.FormatConditions.Add(xlCellValue, xlLess, "=0")

With cond1
.Interior.Color = vbGreen
End With
 
With cond2
.Interior.Color = vbRed
End With

End Sub





