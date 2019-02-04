Sub working()


Dim ws As Worksheet
Dim i As LongLong
Dim tickername As String
Dim tickercounter As Long
Dim lRow As LongLong

For Each ws In ThisWorkbook.Worksheets
 ws.Select

tickercounter = 0
tickername = Cells(i, 1).Value

'lRow = Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To 500
  
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Range("k" & tickercounter + 2).Value = Cells(i, 1).Value
            tickercounter = tickercounter + 1
        End If
 
    Next i

Next ws

End Sub

