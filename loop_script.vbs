Sub LoopScript()
' Initialize Variables and Headers
    Range("J1").Value = "Ticker"
    Range("K1").Value = "Quarterly Change"
    Range("L1").Value = "Percent Change"
    Range("M1").Value = "Total Stock Volume"
    Range("P2").Value = "Greatest % Increase"
    Range("P3").Value = "Greatest % Decrease"
    Range("P4").Value = "Greatest Stock Volume"
    Range("Q1").Value = "Ticker"
    Range("R1").Value = "Value"
    
    Dim ticker As String
    Dim start_value As Double
    Dim end_value As Double
    Dim volume As Double
    Dim j As Integer

    j = 2
    volume = 0
    ticker = Range("A2").Value
    start_value = Range("C2").Value

' Loop through rows and accumulate if the value belongs to the same ticker symbol
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        If ticker = Cells(i, 1).Value Then
            volume = volume + Cells(i, 7).Value
            end_value = Cells(i, 6).Value
        Else
            ticker = Cells(i, 1).Value
            volume = Cells(i, 7).Value
            start_value = Cells(i, 3).Value
            end_value = Cells(i, 6).Value
            j = j + 1
        End If
        Cells(j, 10).Value = ticker
        Cells(j, 11).Value = end_value - start_value
        Cells(j, 12).Value = Round((end_value - start_value) / start_value, 4)
        Cells(j, 13).Value = volume
        ' Format to show proper number format
        Cells(j, 12).NumberFormat = "0.00%"
        Cells(j, 13).NumberFormat = "0"
    Next i
    
    Dim increase As Double
    Dim decrease As Double
    Dim max_volume As Double
    increase = 0
    decrease = 0
    max_volume = 0

    ' Loop through to change the color depending on the value as well as comparing percentage increase, percentage decrease and max volume
    For k = 2 To Cells(Rows.Count, 11).End(xlUp).Row
    
        If Cells(k, 11).Value > 0 Then
            Cells(k, 11).Interior.ColorIndex = 4
            
            If Cells(k, 12).Value > increase Then
                Range("R2").Value = Cells(k, 12).Value
                Range("Q2").Value = Cells(k, 10).Value
                increase = Cells(k, 12).Value
            End If
        ElseIf Cells(k, 11).Value < 0 Then
            Cells(k, 11).Interior.ColorIndex = 3
            If Cells(k, 12).Value < decrease Then
                Range("R3").Value = Cells(k, 12).Value
                Range("Q3").Value = Cells(k, 10).Value
                decrease = Cells(k, 12).Value
            End If
        Else
            Cells(k, 11).Interior.ColorIndex = 2
        End If
        
        If Cells(k, 13).Value > max_volume Then
            Range("R4").Value = Cells(k, 13).Value
            Range("Q4").Value = Cells(k, 10).Value
            max_volume = Cells(k, 13).Value
        End If
    Next k
    ' Format to reflect proper number format
    Range("R2").NumberFormat = "0.00%"
    Range("R3").NumberFormat = "0.00%"
End Sub





