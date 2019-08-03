Sub StockYearlyandPctChange()
    'Define Variables
    Dim ticker As String
    Dim volume As Long
    Dim RunningTotal As Double
    Dim PasteOffset As Integer
    Dim FirstTime As Integer
    Dim RowNum As Long
    'Variables added for Medium Solution
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim PercentChange As Double
    Dim YearlyChange As Double

    'Set Column Names
    Cells(1, 9).Value = "Ticker"
    Cells(1, 12).Value = "Total Volume"
    'Variables added for Medium Solution
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    
    'Determine Total Rows in each sheet for use in first "For" Statement
    RowNum = Cells(Rows.Count, 1).End(xlUp).Row 
    'Set Variables
    FirstTime = 0
    PasteOffset = 2
    'Create Loop
    For i = 2 To RowNum
        'Test if current ticker is same as ticker for last row
        If Cells(i, 1).Value = Cells(i, 1).Offset(1, 0).Value Then
            FirstTime = FirstTime + 1
            'Add Volume cell to running total for current ticker
            RunningTotal = RunningTotal + Cells(i, 7).Value
            If FirstTime = 1 Then
            'If Statement to grab first opening price for stock
                OpeningPrice = Cells(i, 3).Value
            Else
            End If
        Else
            RunningTotal = RunningTotal + Cells(i, 7).Value
            'Write the Ticker and Total Volume in Solution Area
            Cells(PasteOffset, 9).Value = Cells(i, 1).Value
            Cells(PasteOffset, 12).Value = RunningTotal
            'Grab last closing price for stock
            ClosingPrice = Cells(i, 6).Value
            If OpeningPrice <> 0 Then
                'Calculations for PercentChange and YearlyChange
                PercentChange = ((ClosingPrice - OpeningPrice) / OpeningPrice)
                YearlyChange = ClosingPrice - OpeningPrice
            Else
                PercentChange = 0
                YearlyChange = 0
            End If
            Cells(PasteOffset, 11).Value = PercentChange
            'Set Percent Change to % Cell Format
            Cells(PasteOffset, 11).NumberFormat = "0.00%"
            Cells(PasteOffset, 10).Value = YearlyChange
                'Set Conditional Formatting
                If Cells(PasteOffset, 10).Value > 0 Then
                    Cells(PasteOffset, 10).Interior.ColorIndex = 4
                Else
                    Cells(PasteOffset, 10).Interior.ColorIndex = 3
                End If
            RunningTotal = 0
            'Increment paste offset so next group doesnt overwrite previous
            PasteOffset = PasteOffset + 1
            'Reset "FirstTime" counter for next ticker group
            FirstTime = 0
        End If
    Next i
    'Check        
    MsgBox ("It Runs")
    
End Sub