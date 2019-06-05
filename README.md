# Homework2

Sub SumVolume()
   Dim i As Long
   Dim lngLastRow As Long
   Dim strCurrentTicker As String
   Dim ws As Worksheet
   Dim dblSum As Double
   Dim lngRow As Long

   For Each ws In Application.ThisWorkbook.Sheets
       lngRow = 2
       lngLastRow = ws.Cells(1048576, 1).End(xlUp).Row

       ws.Cells(1, 9).Value = “CurrentTicker”
       ws.Cells(1, 10).Value = “Total Stock Volume”

       For i = 2 To lngLastRow
           strCurrentTicker = ws.Cells(i, 1).Value
           dblSum = dblSum + ws.Cells(i, 7).Value

           If strCurrentTicker <> ws.Cells(i + 1, 1).Value Then
               ws.Cells(lngRow, 9).Value = strCurrentTicker
               ws.Cells(lngRow, 10).Value = dblSum
               dblSum = 0
               lngRow = lngRow + 1
           End If
       Next i
   Next ws
End Sub
