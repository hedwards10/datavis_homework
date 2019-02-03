Sub WallStreet()

'Easy:

    'Create a loop that will go through all the stocks

    For Each ws In Worksheets

    'Take ticker symbol

    Dim Ticker As String

    'Take the total Volume of the stock

    Dim Total_Volume As Variant
    Total_Volume = 0

     'Keep track of the location for each Ticker in Row I
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

    'Put the Ticker and Total_Volume values in the sheet

    LastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To LastRow
    
    Ticker = ws.Cells(i, 1).Value

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

    'Insert Headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Total Volume"

    ' Add to the Total Volume
      Total_Volume = Total_Volume + ws.Cells(i, 7).Value

      ' Print the Ticker in Row I
      ws.Range("I" & Summary_Table_Row).Value = Ticker

      ' Print the Total Volume in Row J
      ws.Range("J" & Summary_Table_Row).Value = Total_Volume

      ' Add one to the Ticker in Row I
       Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Total Volume
      Total_Volume = 0

    ' If the cell immediately following a row is the same Ticker
    Else

      ' Add to the Total Volume
      Total_Volume = Total_Volume + ws.Cells(i, 7).Value

    End If
    
    Next i
    
    Next ws

End Sub