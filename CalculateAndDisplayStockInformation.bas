Attribute VB_Name = "Module1"
' Declare constants for Stocks data column indexes
Public Const TICKER_COL As Integer = 1
Public Const OPEN_COL As Integer = 3
Public Const CLOSE_COL As Integer = 6
Public Const VOL_COL As Integer = 7

Sub DisplayStocksInformation()
    ' Declare variables used to describe the Stocks data range on each Worksheet
    Dim maxUsedRow As Long
    Dim maxDataCol As Long
    
    ' Iterate over all the Worksheets in this Workbook (Excel file)
    For wsIndex = 1 To Worksheets.Count
        
        ' Determine the used range for stocks data on the current Worksheet
        maxUsedRow = Worksheets(wsIndex).UsedRange.Rows.Count
        maxDataCol = VOL_COL
        
        ' Ensure the current Worksheet is active, so the following subroutine calls will direct output to it
        Worksheets(wsIndex).Activate
        
        ' Add Stocks Information header
        DisplayStocksInformationHeader maxDataCol
        
        ' Add 'Greatest' Stocks Information header
        DisplayGreatestStocksInformationHeader maxDataCol
        
        ' Add the summary information for each Stock on the current Worksheet
        DisplayWorksheetStocksSummary maxUsedRow, 2, maxDataCol + 2
            
    Next wsIndex
    
End Sub

Sub DisplayStocksInformationHeader(usedCol As Long)
    ' Output the individual Stocks information header (column names)
    Cells(1, usedCol + 2).value = "Ticker"
    Cells(1, usedCol + 3).value = "Quarterly Change"
    Cells(1, usedCol + 4).value = "Percent Change"
    Cells(1, usedCol + 5).value = "Total Stock Volume"
End Sub

Sub DisplayGreatestStocksInformationHeader(usedCol As Long)
    ' Output the 'Greatest' Stocks information header (column names)
    Cells(1, usedCol + 9).value = "Ticker"
    Cells(1, usedCol + 10).value = "Value"
End Sub

Sub DisplayWorksheetStocksSummary(maxDataRow As Long, startOutputRow As Long, startOutputCol As Long)
    ' Declare variables used to collate information or output information on individual Stocks
    Dim prevStockTicker As String
    Dim currStockTicker As String
    Dim stockOpen As Double
    Dim stockClose As Double
    Dim accumulatedStockVolume As Double
    Dim currOutputRow As Long
    Dim currOutputCol As Long
    
    ' Declare variables for tracking current ticker / output location
    prevStockTicker = ""
    currOutputRow = startOutputRow
    currOutputCol = startOutputCol
    
    ' Declare & initialise variables used to create statistics across all Stocks
    Dim greatestPercentageIncreaseTicker As String
    Dim greatestPercentageDecreaseTicker As String
    Dim greatestTotalVolumeTicker As String
    greatestPercentageIncreaseTicker = ""
    greatestPercentageDecreaseTicker = ""
    greatestTotalVolumeTicker = ""
    
    Dim greatestPercentageIncrease As Double
    Dim greatestPercentageDecrease As Double
    Dim greatestTotalVolume As Double
    greatestPercentageIncrease = 0
    greatestPercentageDecrease = 0
    greatestTotalVolume = 0
    
    ' Iterate over Tickers on this Worksheet, to collate then output information on each Stock
    For r = 2 To maxDataRow + 1
        ' Read the Stock Ticker from the current Stock row
        currStockTicker = Cells(r, TICKER_COL).value
        
        ' Output collated information if the Stock Ticker changed (unless the previous Ticker was blank), or
        ' we've reached the end of data on this Worksheet
        If ((currStockTicker <> prevStockTicker And prevStockTicker <> "") Or (r > maxDataRow)) Then
        
            ' Calculate accumulated statistics on the previous Stock
            Dim quarterlyChange As Double
            Dim percentageChange As Double
            
            quarterlyChange = stockClose - stockOpen
            percentageChange = quarterlyChange / stockOpen
            
            ' Output accumulated information on the previous Stock
            OutputCurrentStockSummary currOutputRow, currOutputCol, prevStockTicker, quarterlyChange, percentageChange, accumulatedStockVolume
            
            ' Update statistics considered across all Stocks
            If (percentageChange > 0 And percentageChange > greatestPercentageIncrease) Then
                greatestPercentageIncreaseTicker = prevStockTicker
                greatestPercentageIncrease = percentageChange
            End If
            
            If (percentageChange < 0 And percentageChange < greatestPercentageDecrease) Then
                greatestPercentageDecreaseTicker = prevStockTicker
                greatestPercentageDecrease = percentageChange
            End If
            
            If (accumulatedStockVolume > greatestTotalVolume) Then
                greatestTotalVolumeTicker = prevStockTicker
                greatestTotalVolume = accumulatedStockVolume
            End If
            
            ' Update the output position tracking
            currOutputRow = currOutputRow + 1
            currOutputCol = startOutputCol
        End If
        
        ' Gather other information from the current Stock row
        If (currStockTicker <> prevStockTicker) Then
            ' Initialise accumulator values for the new Stock
            stockOpen = Cells(r, OPEN_COL).value
            stockClose = Cells(r, CLOSE_COL).value
            accumulatedStockVolume = Cells(r, VOL_COL).value
            ' Note the current Ticker has changed from its previous value, so update Ticker tracking accordingly
            prevStockTicker = currStockTicker
        Else
            ' Update / Accumulate current Stock information
            stockClose = Cells(r, CLOSE_COL).value
            accumulatedStockVolume = accumulatedStockVolume + Cells(r, VOL_COL).value
        End If

    Next r
    
    ' Output "greatest" statistics found across all Stocks on this Worksheet
    OutputGreatestPercentageIncrease greatestPercentageIncreaseTicker, greatestPercentageIncrease, startOutputCol + 7
    OutputGreatestPercentageDecrease greatestPercentageDecreaseTicker, greatestPercentageDecrease, startOutputCol + 7
    OutputGreatestTotalVolume greatestTotalVolumeTicker, greatestTotalVolume, startOutputCol + 7
    
End Sub

Sub OutputCurrentStockSummary(outputRow As Long, startOutputCol As Long, _
                              currStockTicker As String, _
                              quarterlyChange As Double, percentageChange As Double, accumulatedStockVolume As Double)
    ' Output the Stock Ticker
    Cells(outputRow, startOutputCol).value = currStockTicker
    
    ' Output the Stock quarterly change, including number format 0.00 and background colour conditional formatting
    Cells(outputRow, startOutputCol + 1).value = quarterlyChange
    Cells(outputRow, startOutputCol + 1).NumberFormat = "0.00"
    If (quarterlyChange < 0) Then
        ' Set cell backgroun red for negative quarterly change
        Cells(outputRow, startOutputCol + 1).Interior.ColorIndex = 3
    ElseIf (quarterlyChange > 0) Then
        ' Set cell background green for positive quarterly change
        Cells(outputRow, startOutputCol + 1).Interior.ColorIndex = 4
    Else
        ' Set cell background white for zero quarterly change
        Cells(outputRow, startOutputCol + 1).Interior.ColorIndex = 2
    End If
    
    ' Output the Stock percentage change for the quarter, including percentage number formatting
    Cells(outputRow, startOutputCol + 2).value = percentageChange
    Cells(outputRow, startOutputCol + 2).NumberFormat = "0.00%"
    
    ' Output the Stock's accumulated Volume for the quarter
    Cells(outputRow, startOutputCol + 3).value = accumulatedStockVolume
End Sub

Sub OutputGreatestPercentageIncrease(ticker As String, value As Double, startOutputGreatestCol As Integer)
    Cells(2, startOutputGreatestCol - 1).value = "Greatest % Increase"
    Cells(2, startOutputGreatestCol + 0).value = ticker
    Cells(2, startOutputGreatestCol + 1).value = value
    Cells(2, startOutputGreatestCol + 1).NumberFormat = "0.00%"
End Sub

Sub OutputGreatestPercentageDecrease(ticker As String, value As Double, startOutputGreatestCol As Integer)
    Cells(3, startOutputGreatestCol - 1).value = "Greatest % Decrease"
    Cells(3, startOutputGreatestCol + 0).value = ticker
    Cells(3, startOutputGreatestCol + 1).value = value
    Cells(3, startOutputGreatestCol + 1).NumberFormat = "0.00%"
End Sub

Sub OutputGreatestTotalVolume(ticker As String, value As Double, startOutputGreatestCol As Integer)
    Cells(4, startOutputGreatestCol - 1).value = "Greatest Total Volume"
    Cells(4, startOutputGreatestCol + 0).value = ticker
    Cells(4, startOutputGreatestCol + 1).value = value
End Sub






