Attribute VB_Name = "Module1"
' Declare constants for Stocks data column indexes
Public Const TICKER_COL As Integer = 1
Public Const OPEN_COL As Integer = 3
Public Const CLOSE_COL As Integer = 6
Public Const VOL_COL As Integer = 7

Sub DisplayStocksInformation()
    ' Declare variables used to describe the Stocks data on each Worksheet
    Dim maxUsedRow As Long
    Dim maxDataCol As Long
    
    ' Iterate over all the Worksheets in this Workbook (Excel file)
    For wsIndex = 1 To Worksheets.Count
        
        ' Determine the used range for stocks data on the current Worksheet
        maxUsedRow = Worksheets(wsIndex).UsedRange.Rows.Count
        maxDataCol = VOL_COL
        'MsgBox ("Worksheet " & wsIndex & " maxUsedRow " & maxUsedRow & " maxDataCol " & maxDataCol)
        
        ' Ensure the current Worksheet is active, so the following subroutine calls will direct output to it
        Worksheets(wsIndex).Activate
        
        ' Add Stocks Information header
        DisplayStocksInformationHeader maxDataCol
        
        ' Add the Ticker information
        DisplayWorksheetStocksSummary maxUsedRow, 2, maxDataCol + 2
        
    Next wsIndex
End Sub

Sub DisplayStocksInformationHeader(usedCol As Long)
    ' Output the Stocks Information header (column names)
    Cells(1, usedCol + 2).Value = "Ticker"
    Cells(1, usedCol + 3).Value = "Quarterly Change"
    Cells(1, usedCol + 4).Value = "Percent Change"
    Cells(1, usedCol + 5).Value = "Total Stock Volume"
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
    
    ' Iterate over Tickers on this Worksheet, to collate then output information on each Stock
    For r = 2 To maxDataRow + 1
        ' Read the Stock Ticker from the current Stock row
        currStockTicker = Cells(r, TICKER_COL).Value
        
        ' Output collated information if the Stock Ticker changed (unless the previous Ticker was blank), or
        ' we've reached the end of data on this Worksheet
        If ((currStockTicker <> prevStockTicker And prevStockTicker <> "") Or (r > maxDataRow)) Then
        
            ' Output accumulated information on the previous Stock
            OutputCurrentStockSummary currOutputRow, currOutputCol, prevStockTicker, stockOpen, stockClose, accumulatedStockVolume
            
            ' Update the output position tracking
            currOutputRow = currOutputRow + 1
            currOutputCol = startOutputCol
        End If
        
        ' Gather other information from the current Stock row
        If (currStockTicker <> prevStockTicker) Then
            ' Initialise accumulator values for the new Stock
            stockOpen = Cells(r, OPEN_COL).Value
            stockClose = Cells(r, CLOSE_COL).Value
            accumulatedStockVolume = Cells(r, VOL_COL).Value
            ' Note the current Ticker has changed from its previous value, so update Ticker tracking accordingly
            prevStockTicker = currStockTicker
        Else
            ' Update / Accumulate current Stock information
            stockClose = Cells(r, CLOSE_COL).Value
            accumulatedStockVolume = accumulatedStockVolume + Cells(r, VOL_COL).Value
        End If

    Next r
    
End Sub

Sub OutputCurrentStockSummary(outputRow As Long, startOutputCol As Long, _
                              currStockTicker As String, _
                              stockOpen As Double, stockClose As Double, accumulatedStockVolume As Double)
    ' Output the Stock Ticker
    Cells(outputRow, startOutputCol).Value = currStockTicker
    
    ' Output the Stock quarterly change, including number format 0.00 and conditional formatting
    Dim quarterlyChange As Double
    quarterlyChange = stockClose - stockOpen
    Cells(outputRow, startOutputCol + 1).Value = quarterlyChange
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
    Cells(outputRow, startOutputCol + 2).Value = quarterlyChange / stockOpen
    Cells(outputRow, startOutputCol + 2).NumberFormat = "0.00%"
    
    ' Output the Stock's accumulated Volume for the quarter
    Cells(outputRow, startOutputCol + 3).Value = accumulatedStockVolume
End Sub







