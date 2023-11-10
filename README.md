# Module-2
Sub stock_analysis():

    ' Set dimensions
    Dim total As Double
    Dim i As Long
    Dim change As Double
    Dim j As Integer
    Dim start As Long
    Dim rowCount As Long
    Dim PercentChange As Double
    Dim days As Integer
    Dim dailyChange As Double
    Dim averageChange As Double
    Dim OpenPrice As Double
    Dim Closeprice As Double

    
    Dim ws As Worksheet
    
    ' Loop through each worksheet (tab) in the Excel file
    For Each ws In Worksheets
        ' Initialize values for each worksheet
        j = 0
        total = 0
        change = 0
        start = 2
        dailyChange = 0
        
        OpenPrice = 0
        Closeprice = 0
        TotalVolume = 0

        ' Set title row
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volue"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest & Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"

        ' get the row number of the last row with data
        rowCount = Cells(Rows.Count, "A").End(xlUp).Row

        For i = 2 To rowCount

            ' If ticker changes then print results
           If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' Stores results in variables
            Ticker = Cells(i, 1).Value
            
            Closeprice = ws.Cells(i, 6).Value
            
            PriceChange = Closeprice - OpenPrice


       End If

    Next i

Next ws

End Sub
