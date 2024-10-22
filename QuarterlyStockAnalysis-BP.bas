Attribute VB_Name = "Module2"
Sub QuarterlyStockAnalysis()

On Error GoTo ErrorHandler  ' enable error handling
    
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets ' loop through all sheets


    '------------------------------------------------------------------
    ' FIRST ANALYSIS
    '       for each stock, report quarterly change, percent change over quarter, and total stock volume over quarter;
    '       format positive quarterly changes in green and negative quarterly changes in red; format percent changes as 0.00%
    '------------------------------------------------------------------

    ' Declare input variables from data table
    Dim Ticker As String
    Dim OpeningPrice
    Dim ClosingPrice As Double
    ' Declare output table variables
    Dim QuarterlyChange
    Dim PercentChange As Double
    Dim TotalVolume As LongLong
    ' Declare other variables used in first analysis
    Dim LastRow As Long ' a variable that stores number of rows in data table
    Dim Row As Long ' an index for data table row
    Dim OutputRow As Long ' an index for output table row
        
    ' Add column labels to output table
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarterly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    ' Find last row of data table based on column A ("Cntl - Shift - End" method)
    LastRow = ws.Cells(Rows.count, 1).End(xlUp).Row
    
    ' Initialize variables
    TotalVolume = 0
    OutputRow = 2
    
    ' Loop through each row of data table
    For Row = 2 To LastRow

        ' if ticker not the same as prior ticker, we have first row of a stock: grab ticker, grab opening price for quarter, start running volume
        If ws.Cells(Row, 1).Value <> ws.Cells(Row - 1, 1).Value Then
            
            Ticker = ws.Cells(Row, 1).Value
            OpeningPrice = ws.Cells(Row, 3).Value
            TotalVolume = ws.Cells(Row, 7).Value
            
        ' else ticker same as prior ticker: increase running volume, check to see if at last row for stock; if last row for stock, calculate outputs and report to output table
        Else
        
            TotalVolume = TotalVolume + ws.Cells(Row, 7).Value
            
            If ws.Cells(Row, 1).Value <> ws.Cells(Row + 1, 1).Value Then 'if next row's stock does not match current row's stock, then we are at last row for stock
                
                ' grab closing price for stock, calculate quarterly change in stock price, calculate percent change in stock price (ensure not divide by 0)
                ClosingPrice = ws.Cells(Row, 6).Value
                QuarterlyChange = ClosingPrice - OpeningPrice
                If OpeningPrice <> 0 Then
                    PercentChange = (QuarterlyChange / OpeningPrice)
                Else
                    PercentChange = 0
                End If
                    
                ' report outputs to output table
                ws.Cells(OutputRow, 9).Value = Ticker
                ws.Cells(OutputRow, 10).Value = QuarterlyChange
                ws.Cells(OutputRow, 11).Value = PercentChange
                ws.Cells(OutputRow, 12).Value = TotalVolume

                ' advance to next output row
                OutputRow = OutputRow + 1

            End If
                 
        End If
  
     ' advance to next data row
    Next Row
    
    ' Format output table to match image in instructions
    Dim i As Long
    ws.Range("J2:J" & LastRow).NumberFormat = "0.00"
    For i = 2 To LastRow
        If Range("J" & i) > 0 Then
                ws.Cells(i, 10).Interior.Color = RGB(0, 255, 0)
        ElseIf Range("J" & i) < 0 Then
               ws.Cells(i, 10).Interior.Color = RGB(255, 0, 0)
        End If
    Next i
    ws.Range("K2:K" & LastRow).NumberFormat = "0.00%"
    ws.Columns("I:L").AutoFit
    
    
    '------------------------------------------------------------------
    ' SECOND ANALYSIS
    '       search over output table from first analysis for greatest % increase (GPI), greatest % decrease (GPD), and greatest total volume (GTV)
    '       report values of GPI, GPD, and GTV along with the respective tickers; format output as in instructions
    '------------------------------------------------------------------

    ' Declare variables used in second analysis
    Dim LastRow2 As Long ' variable that stores number of rows of data in output table from first analysis
    Dim GPIRow, GPDRow, GTVRow As Long 'used to store row of GPI/GPD/GTV
    
    ' Add column and row labels to output table
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"
    
    ' Find last row of data table based on column I, "Cntl - Shift - End" method
    LastRow2 = ws.Cells(Rows.count, 9).End(xlUp).Row
        
    ' Find stock with GPI from output table of first analysis; report its ticker and the GPI value
    ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & LastRow2)) ' finds and reports the GPI value
    GPIRow = Application.Match(ws.Range("Q2"), ws.Range("K2:K" & LastRow2), 0) + 1 ' identifies the row of GPI value; adds 1 to adjust for the header row
    ws.Range("P2").Value = ws.Cells(GPIRow, 9).Value ' finds and reports the ticker of GPI row
    
    'less compact form for GPI (similarly for GPD and GTV) that is more explicit:
        'Dim GPIValue As Double
        'Dim GPITicker As String
        'GPIValue = WorksheetFunction.Max(ws.Range("K2:K" & LastRow2))
        'GPIRow = Application.Match(GPIValue, ws.Range("K2:K" & LastRow2), 0) + 1 ' add 1 to adjust for the header row
        'GPITicker = ws.Cells(GPIRow, 9).Value
        'ws.Range("Q2") = GPIValue
        'ws.Range("P2") = GPITicker
    
    ' Find stock with GPD from output table of first analysis; report its ticker and the GPD value
    ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & LastRow2)) ' finds and reports the GPD value
    GPDRow = Application.Match(ws.Range("Q3"), ws.Range("K2:K" & LastRow2), 0) + 1 ' identifies the row of GPD value; adds 1 to adjust for the header row
    ws.Range("P3").Value = ws.Cells(GPDRow, 9).Value  ' finds and reports the ticker of GPD row

    ' Find stock with GTV from output table of first analysis; report its ticker and the GTV value
    ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & LastRow2)) ' finds and reports the GTV value
    GTVRow = Application.Match(ws.Range("Q4"), ws.Range("L2:L" & LastRow2), 0) + 1 ' identifies the row of GPD value; adds 1 to adjust for the header row
    ws.Range("P4").Value = ws.Cells(GTVRow, 9).Value ' finds and reports the ticker of GTV row
    
    ' Format output
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    ws.Range("Q4").NumberFormat = "#.00E+0"
    ws.Columns("O:Q").AutoFit
    
Next ws

Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description & " at row " & i

End Sub


' alt to application.match to identify cell for GPI (similar for GPD and GTV):
    'GPI = WorksheetFunction.Max(ws.Range("K2:K" & LastRow2))
    'Set CellGPI = ws.Range("K2:K" & LastRow2).Find(GPI) ' find function searches over numeric, but GPI is % which reads as string
    'GPITicker = CellGPI.Offset(0, -2).Value
    'ws.Range("Q2") = GPI
    'ws.Range("P2") = GPITicker
