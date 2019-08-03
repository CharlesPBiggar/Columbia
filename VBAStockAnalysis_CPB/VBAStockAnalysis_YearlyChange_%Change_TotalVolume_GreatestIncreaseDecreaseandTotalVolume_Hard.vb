Sub StockAnalysisHard()
    'Define Variables
    Dim ticker As String
    Dim volume As Long
    Dim RunningTotal As Double
    Dim PasteOffset As Integer
    Dim FirstTime As Integer
    Dim RowNum As Long
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim PercentChange As Double
    Dim YearlyChange As Double
    'Variables added for Hard Solution
    Dim GreatestPctIncrease As Double
    Dim GPITicker As String
    Dim GreatestPctDecrease As Double
    Dim GPDTicker As String
    Dim GreatestTotalVolume As Double
    Dim GTVTicker As String
    Dim RowNum2 As Long

    'Set Column Names
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Volume"
    'Variables added for Hard Solution
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    'Determine Total Rows Column A in sheet for use in first "For" Statement
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
        
    'Determine Total Rows Column I in sheet for use in second "For" Statement
    RowNum2 = Cells(Rows.Count, 9).End(xlUp).Row
    GreatestTotalVolume = 0
    GreatestPctIncrease = 0
    GreatestPctDecrease = 0
    'Next Look to determine Greatest Values from previous loop
    For j = 2 To RowNum2
        If Cells(j, 12).Value > GreatestTotalVolume Then
            GreatestTotalVolume = Cells(j, 12).Value
            GTVTicker = Cells(j, 9).Value  
        End If
        'Sets Values for Greatest Total Volume Solution
        Cells(4, 17).Value = GreatestTotalVolume
        Cells(4, 16).Value = GTVTicker
        If Cells(j, 11).Value > GreatestPctIncrease Then
            GreatestPctIncrease = Cells(j, 11).Value
            GPITicker = Cells(j, 9).Value      
        End If
        'Sets Values for Greatest % Increase Solution
        Cells(2, 17).Value = GreatestPctIncrease
        Cells(2, 17).NumberFormat = "0.00%"
        Cells(2, 16).Value = GPITicker
        If Cells(j, 11).Value < GreatestPctDecrease Then
            GreatestPctDecrease = Cells(j, 11).Value
            GPDTicker = Cells(j, 9).Value       
        End If
        'Sets Values for Greatest % Decrease Solution
        Cells(3, 17).Value = GreatestPctDecrease
        Cells(3, 17).NumberFormat = "0.00%"
        Cells(3, 16).Value = GPDTicker
    Next j
    'Check        
    MsgBox ("Here is the data that you asked for!")
    
End Sub