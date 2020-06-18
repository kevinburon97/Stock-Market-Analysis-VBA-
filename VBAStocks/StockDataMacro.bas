Attribute VB_Name = "Module1"

Option Explicit

Sub StockData()
Dim ws As Worksheet
For Each ws In Worksheets

'Set the ticker in each row
Dim i As Long
Dim numrows As Long
Dim LastRow As Long
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

numrows = 2

    For i = 2 To LastRow
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            ws.Cells(numrows, 9).Value = ws.Cells(i, 1).Value
            numrows = numrows + 1
        End If
    Next i

Dim rowCounter As Long
Dim RCpctg As Long
Dim RCstockVol As Double
Dim totalVolume As Double
Dim q As Double 'q is the cell value for close
Dim w As Double 'w is the cell value for open
rowCounter = 2
RCpctg = 2 'RC = row counter
RCstockVol = 2
totalVolume = 0
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("k1").Value = "Percent Change"
ws.Range("l1").Value = "Total volume"
    
    For i = 2 To LastRow
        totalVolume = totalVolume + ws.Cells(i, 7).Value
  
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            q = ws.Cells(i, 6).Value
            w = ws.Cells(rowCounter, 3).Value
            ws.Cells(RCpctg, 11).Value = 0
            ws.Cells(RCpctg, 10).Value = q - w 'net change
                If ws.Cells(RCpctg, 10).Value = 0 Or w = 0 Then
                    ws.Cells(RCpctg, 11).Value = 0
                Else
                    ws.Cells(RCpctg, 11).Value = ws.Cells(RCpctg, 10).Value / w 'percentage change
                End If
            ws.Cells(RCpctg, 11).NumberFormat = "0.00%" ' display as percentage
                
                If ws.Cells(RCpctg, 10).Value > 0 Then ' color of cells
                    ws.Cells(RCpctg, 10).Interior.Color = vbGreen
                Else
                    ws.Cells(RCpctg, 10).Interior.Color = vbRed
                End If
        rowCounter = i + 1
    
        ws.Cells(RCstockVol, 12).Value = totalVolume
        totalVolume = 0
        RCpctg = RCpctg + 1
        RCstockVol = RCstockVol + 1
    End If
Next i

ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"
ws.Range("O1").Value = "Ticker"
ws.Range("P1").Value = "Value"

For i = 2 To LastRow
    If ws.Cells(i, 11).Value > ws.Cells(2, 16).Value Then
        ws.Cells(2, 16).Value = ws.Cells(i, 11).Value
        ws.Cells(2, 15).Value = ws.Cells(i, 9).Value
        ws.Cells(2, 16).NumberFormat = "0.00%"
    End If
 
    If ws.Cells(i, 11).Value < ws.Cells(3, 16).Value Then
        ws.Cells(3, 16).Value = ws.Cells(i, 11).Value
        ws.Cells(3, 15).Value = ws.Cells(i, 9).Value
        ws.Cells(3, 16).NumberFormat = "0.00%"
    End If
 
    If ws.Cells(i, 12).Value > ws.Cells(4, 16).Value Then
        ws.Cells(4, 16).Value = ws.Cells(i, 12).Value
        ws.Cells(4, 15).Value = ws.Cells(i, 9).Value
        
    End If
Next i
Next ws

End Sub

