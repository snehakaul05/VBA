Sub AnalyzeAllSheets():

Dim ws_count As Integer
ws_count = ActiveWorkbook.Worksheets.Count

For x = 1 To ws_count
    Worksheets(x).Activate
    SummarizeSheet
Next x

End Sub

Sub AnalyzeSheets()

    Dim row As Long
    Dim column As Long
    Dim stock As String
    Dim newstock As String
    Dim outputrow As Long
    Dim volume As Double
    Dim start As Double
    Dim percentageChange As Double
    Dim delta As Double

    With ActiveSheet.Sort
        .SortFields.Add Key:=Range("A1"), Order:=xlAscending
        .SortFields.Add Key:=Range("B1"), Order:=xlAscending
        .SetRange Range("A1").CurrentRegion
        .Header = xlYes
        .Apply
    End With

    output = 2
    volume = 0
    rowmax = ActiveSheet.UsedRange.Rows.Count
    colmax = ActiveSheet.UsedRange.Columns.Count
    stock = Cells(2, 1).Value
    start = Cells(2, 3).Value

    Cells(1, colmax + 2).Value = "Ticker"
    Cells(1, colmax + 3).Value = "Yearly Change"
    Cells(1, colmax + 4).Value = "Percent Change"
    Cells(1, colmax + 5).Value = "Total Stock Volume"

For Row = 2 To rowmax
    stock = Cells(Row, 1).Value
    newstock = Cells(Row + 1, 1).Value
    volume = volume + Cells(Row, 7)

    If stock <> newstock Then

    Cells(output, colmax + 2).Value = stock
    delta = Cells(Row, 6).Value - start

    If start <> 0 Then 
        percentageChange = Cells(Row, 6).Value / start - 1
    ElseIf start = 0 Then
        percentageChange = 0
        Cells(output, colmax + 4).Value = percentageChange
    End If

    Cells(output, colmax + 5).Value = volume

    output = output + 1

    start = Cells(Row + 1, 3).Value
    volume = 0

    End If

Next Row

endrow = outputrow - 1

For Row = 2 To endrow
    positive = Cells(Row, colmax + 3) >= 0

    If positive = "True" Then
            Cells(Row, colmax + 3).Interior.Color = RGB(0, 235, 30)
        
        Else
            Cells(Row, colmax + 3).Interior.Color = RGB(235, 20, 20)
         End If
Next Row

Dim max_change As Double
    Dim min_change As Double
    Dim max_volume As Double
    Dim name_max_change As String
    Dim name_min_change As String
    Dim name_max_volume As String
    
    max_change = 0
    min_change = 0
    max_volume = 0

For Row = 2 To endrow
        If Cells(Row, colmax + 4).Value >= max_change Then
            max_change = Cells(Row, colmax + 4).Value
            name_max_change = Cells(Row, colmax + 2).Value
        End If
        
        If Cells(Row, colmax + 4).Value < min_change Then
            min_change = Cells(Row, colmax + 4).Value
            name_min_change = Cells(Row, colmax + 2).Value
        End If
        
        If Cells(Row, colmax + 5).Value >= max_volume Then
            max_volume = Cells(Row, colmax + 5).Value
            name_max_volume = Cells(Row, colmax + 2).Value
        End If
    Next Row

    Cells(1, colmax + 7).Value = "Metric"
    Cells(1, colmax + 8).Value = "Ticker"
    Cells(1, colmax + 9).Value = "Value"
    
    Cells(2, colmax + 7).Value = "Greatest % Increase"
    Cells(2, colmax + 8).Value = name_max_change
    Cells(2, colmax + 9).Value = max_change
    
    Cells(3, colmax + 7).Value = "Greatest % Decrease"
    Cells(3, colmax + 8).Value = name_min_change
    Cells(3, colmax + 9).Value = min_change
       
    Cells(4, colmax + 7).Value = "Greatest Volume"
    Cells(4, colmax + 8).Value = name_max_volume
    Cells(4, colmax + 9).Value = max_volume

    Range("J:J").NumberFormat = "$###,###.00"
    Range("K:K").NumberFormat = "0.00%"
    Range("L:L").NumberFormat = "###,###,###,###"
    Range("P2").NumberFormat = "0.00%"
    Range("P3").NumberFormat = "0.00%"
    Range("P4").NumberFormat = "###,###,###,###"
    Columns("L").ColumnWidth = 20
    Columns("N").ColumnWidth = 20

End Sub