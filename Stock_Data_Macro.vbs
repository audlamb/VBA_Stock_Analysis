Sub stock_ticker()

Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
    
    Dim Ticker As String
    
    Dim Volume_Total As LongLong
    Volume_Total = 0
    
    Dim Summary_Table_Row As Long
    Summary_Table_Row = 2
  
    Range("I1").Value = "Ticker"
  
    Range("J1").Value = "Total Volume"
    
    Dim LastRow As LongLong
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To LastRow
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            Ticker = Cells(i, 1).Value
            
            Volume_Total = Volume_Total + Cells(i, 7).Value
            
            Range("I" & Summary_Table_Row).Value = Ticker
            
            Range("J" & Summary_Table_Row).Value = Volume_Total
            
            Summary_Table_Row = Summary_Table_Row + 1
      
            Volume_Total = 0
        Else
            Volume_Total = Volume_Total + Cells(i, 7).Value
        End If
    Next i
  
Next ws

MsgBox ("Program Complete")

End Sub
