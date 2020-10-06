Attribute VB_Name = "Module1"
Sub HW()
    Dim i As Long
    Dim Ticker As String
    Dim NxtTicker As String
    Dim Count As Integer
    Dim OpenValue As Double
    Dim CloseValue As Double
    Dim LastRow As Long
    Dim TotalVol As LongLong
    Dim ws_name As String
    Dim ws As Worksheet
    Dim GreatVal As String
    Dim GreatPct As Double
    Dim DecreaseVal As String
    Dim DecreasePct As Double
    Dim GreatVol As String
    Dim GreatVolVal As LongLong
    
    
    
    
    For Each ws In Worksheets
        ws.Activate
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Yearly Change"
        ws.Cells(1, 11) = "Percent Change"
        ws.Cells(1, 12) = "Total Stock Volume"
        Count = 2
        OpenValue = 0
        CloseValue = 0
        TotalVol = 0
        LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row
        ws_name = ws.Name
        MsgBox (ws_name)
        
        OpenValue = ws.Cells(2, 3).Value
        
        For i = 2 To LastRow
    
            Ticker = ws.Cells(i, 1).Value
            NxtTicker = ws.Cells(i + 1, 1).Value
            TotalVol = TotalVol + ws.Cells(i, 7).Value
            
            
             
            If Ticker <> NxtTicker Then
                If OpenValue = 0 And ws.Cells(i + 1, 3) <> 0 Then
                    OpenValue = ws.Cells(i + 1, 3).Value
                End If
                CloseValue = ws.Cells(i, 6).Value
                ws.Cells(Count, 9).Value = Ticker
                ws.Cells(Count, 10).Value = CloseValue - OpenValue
                
                If ws.Cells(Count, 10).Value > 0 Then
                    ws.Cells(Count, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(Count, 10).Interior.ColorIndex = 3
                End If
                    
                ws.Cells(Count, 11).Value = ((CloseValue - OpenValue) / OpenValue)
                ws.Cells(Count, 11).NumberFormat = "0.00%"
                ws.Cells(Count, 12).Value = TotalVol
                
                Count = Count + 1
                OpenValue = ws.Cells(i, 3).Value
                
            End If
            
        Next i
        
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        
        GreatPct = Application.WorksheetFunction.Max(ws.Range("K:K"))
        ws.Cells(2, 14).Value = "Greatest % Increase"
        'ws.Cells(2, 15).Value = Cells(GreatPct, 9).Value
        ws.Cells(2, 16).Value = GreatPct
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        'ws.Cells(3,16).Value =
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        'ws.Cells(5, 14).Value =
    Next ws
End Sub


