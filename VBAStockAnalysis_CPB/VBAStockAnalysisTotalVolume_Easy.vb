Sub StockTotalVolume()
    'Define Variables
    Dim ticker As String
    Dim volume As Long
    Dim RunningTotal As Double
    Dim PasteOffset As Integer
    Dim FirstTime As Integer
    Dim RowNum As Long
    
    'Set Column Names
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Total Volume"

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
        Else
            RunningTotal = RunningTotal + Cells(i, 7).Value
            'Write the Ticker and Total Volume in Solution Area
            Cells(PasteOffset, 9).Value = Cells(i, 1).Value
            Cells(PasteOffset, 10).Value = RunningTotal
            RunningTotal = 0
            'Increment paste offset so next group doesnt overwrite previous
            PasteOffset = PasteOffset + 1
            'Reset "FirstTime" counter for next ticker group
            FirstTime = 0
        End If
    Next i
    'Check
    MsgBox ("Here is the data that you asked for!")
    
End Sub