Attribute VB_Name = "Module1"
Sub stockmarket_bigdata()
Dim i As Double
Dim totalval As Double
Dim ws As Worksheet
Dim start_index As Double
Dim end_index As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim max_stock As Double
Dim greatest_increase As Double
Dim greatest_decrease As Double

For Each ws In ActiveWorkbook.Sheets
ws.Activate
    Range("I1:P1000").Clear
    Cells(1, 9) = "Ticker"
    Cells(1, 10) = "Yearly Change"
    Cells(1, 11) = "Percent Change"
    Cells(1, 12) = "Total Stock Volume"
    Cells(2, 14) = "Greatest % increase"
    Cells(3, 14) = "Greatest % decrease"
    Cells(4, 14) = "Greatest total volume"
    Cells(1, 15) = "Ticker"
    Cells(1, 16) = "Value"
    start_index = 2
            
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        If Cells(i, 1) = Cells(i + 1, 1) Then
            totalval = totalval + Cells(i, 7)
        ElseIf Cells(i, 1) <> Cells(i + 1, 1) Then
            end_index = i
            yearly_change = Cells(end_index, 6) - Cells(start_index, 3)
            Cells(Cells(Rows.Count, 10).End(xlUp).Row + 1, 10) = yearly_change
            If yearly_change < 0 Then
                Cells(Cells(Rows.Count, 10).End(xlUp).Row, 10).Interior.ColorIndex = 3
                Else
                Cells(Cells(Rows.Count, 10).End(xlUp).Row, 10).Interior.ColorIndex = 4
                End If
            If Cells(start_index, 3) = 0 Then
                Cells(Cells(Rows.Count, 11).End(xlUp).Row + 1, 11) = "N/A"
            Else
                percent_change = yearly_change / Cells(start_index, 3)
                Cells(Cells(Rows.Count, 11).End(xlUp).Row + 1, 11) = percent_change
            End If
            
            start_index = end_index + 1
            
            totalval = totalval + Cells(i, 7)
            Cells(Cells(Rows.Count, 9).End(xlUp).Row + 1, 9) = Cells(i, 1)
            Cells(Cells(Rows.Count, 12).End(xlUp).Row + 1, 12) = totalval
            totalval = 0
        End If
    Next i
    
    max_stock = 0
    greatest_increase = 0
    greatest_decrease = 0
    
    For i = 2 To Cells(Rows.Count, 9).End(xlUp).Row
        
        If Cells(i, 12) > max_stock Then
        max_stock = Cells(i, 12)
        Cells(4, 15) = Cells(i, 9)
        Cells(4, 16) = max_stock
        End If
        
        If Cells(i, 11) <> "N/A" Then
        
        If Cells(i, 11) > greatest_increase Then
        greatest_increase = Cells(i, 11)
        Cells(2, 15) = Cells(i, 9)
        Cells(2, 16) = greatest_increase
                        
        ElseIf Cells(i, 11) < greatest_decrease Then
        greatest_decrease = Cells(i, 11)
        Cells(3, 15) = Cells(i, 9)
        Cells(3, 16) = greatest_decrease
        
        End If
        
    End If
        
    Next i
   
Next

End Sub
