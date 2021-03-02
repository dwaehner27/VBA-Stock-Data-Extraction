Sub stockmarket()
'set header for outputs
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
'set variables
    Dim yearlyvalue As Double
    Dim ticker As String
    Dim openx As Double
    Dim closex As Double
    Dim volume As Double
    Dim summarytablerow As Integer
    Dim percent As Double
    volume = 0
    openx = 0
    closex = 0
    volume = 0
    ticker = " "
    summarytablerow = 2
    yearlyvalue = closex - openx
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
'loops
    For i = 2 To lastrow
        If Cells(i + 1, 1).Value = Cells(i, 1).Value And Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
            ticker = Cells(i, 1).Text
            openx = Cells(i, 3).Value
            Cells(summarytablerow, 9).Value = ticker
            volume = Cells(i, 7).Value
        ElseIf Cells(i + 1, 1) = Cells(i, 1) Then
            volume = Cells(i, 7).Value + volume
        ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value And Cells(i - 1, 1).Value = Cells(i, 1) Then
            closex = Cells(i, 6).Value
            volume = Cells(i, 7).Value + volume
            Cells(summarytablerow, 10).Value = closex - openx
            Cells(summarytablerow, 12).Value = volume
            Cells(summarytablerow, 11).Value = openx / closex
            summarytablerow = summarytablerow + 1
            volume = 0
            yearlyvalue = 0
            percent = 0
            ticker = " "
            openx = 0
            closex = 0
        End If
    Next i
'formatting
    For i = 2 To lastrow
        If Cells(i, 10) > 0 Then
            Cells(i, 10).Interior.ColorIndex = 4
        ElseIf Cells(i, 10) < 0 Then
            Cells(i, 10).Interior.ColorIndex = 3
        Else
            Cells(i, 10).Interior.ColorIndex = 0
        End If
    Next i
    
End Sub
