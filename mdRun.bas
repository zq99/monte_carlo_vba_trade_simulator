Attribute VB_Name = "mdRun"
Option Explicit


Public Sub ClearUI()

'Purpose: Reset this tool and any input ranges

    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets("Control")
    ws.Range("OUTPUT").ClearContents
    
    Set ws = Nothing

End Sub


Public Sub StartMonteCarloSimulation()
    
'Purpose: This is the main routine that takes the parameters from the worksheet and runs the simulation
    
    Dim vntTradeList As Variant
    Dim iCalc As Integer
    Dim blnScreenUpdating As Boolean
    Dim collFinalResults As Collection
    Dim oResult As clsResult
    Dim lRow As Long
    Dim iCol As Integer
    Dim oSimulation As clsSimulation
    Dim ws As Worksheet
    Dim intTotalRuns As Integer
    Dim intLotSize As Integer
    Dim dblStartEquity As Double
    Dim dblMarginLimit As Double
    Dim intTradesInYear As Integer
    
    iCalc = Application.Calculation
    Application.Calculation = xlManual
    
    blnScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    ' clear the output from previous results
    Call ClearUI
    
    ' get the list of trades from the input worksheet
    vntTradeList = fncGetTrades()
    If UBound(vntTradeList) = 0 Then
        MsgBox "No trade list found!", vbExclamation + vbOKOnly, "Input Data"
        GoTo Exit_Here
    End If
    
    Set ws = ThisWorkbook.Sheets("Control")
    With ws
    
        ' parameters for the simulation
        intTotalRuns = .Range("TOTAL_RUNS").value
        intLotSize = .Range("LOT_SIZE").value
        intTradesInYear = .Range("TRADES_IN_YEAR").value
        dblStartEquity = .Range("START_EQUITY").value
        dblMarginLimit = .Range("MARGIN_LIMIT").value
        
        'create a simulation object to run with the parameters
        Set oSimulation = mdFactory.CreateSimulation(totalRuns:=intTotalRuns, _
            tradesInYear:=intTradesInYear, lotSize:=intLotSize, TradeList:=vntTradeList, _
            startEquity:=dblStartEquity, margin:=dblMarginLimit)
    
        If Not oSimulation Is Nothing Then
            lRow = .Range("OUTPUT_START_CELL").Row
            iCol = .Range("OUTPUT_START_CELL").Column
            
            'run the simulation
            Set collFinalResults = oSimulation.fncRunProcess()
            
            'output the results of the simulation
            If Not collFinalResults Is Nothing Then
                For Each oResult In collFinalResults
                   
                   .Cells(lRow, iCol).value = oResult.Equity
                   .Cells(lRow, iCol + 1).value = oResult.Ruin
                   .Cells(lRow, iCol + 2).value = oResult.MedianDrawdown
                   .Cells(lRow, iCol + 3).value = oResult.MedianProfit
                   .Cells(lRow, iCol + 4).value = oResult.MedianReturn
                   .Cells(lRow, iCol + 5).value = oResult.MedianReturnDD
    
                   lRow = lRow + 1
                Next oResult
            End If
        End If
    
        ws.Select
    End With
    
    MsgBox "Process complete!", vbOKOnly + vbInformation, "Simulation"
    
Exit_Here:

    Set ws = Nothing
    Set oResult = Nothing
    Set collFinalResults = Nothing
    Set oSimulation = Nothing
                
    Application.Calculation = iCalc
    Application.ScreenUpdating = blnScreenUpdating
    

End Sub


Function fncGetTrades() As Variant

'Purpose: return the input pnl trades as a one dimensional array

    Dim ws As Worksheet
    Dim rng As Range
    Dim arr As Variant
    Dim lnglastRow As Long
    Dim lngfirstRow As Long
    
    Set ws = ThisWorkbook.Worksheets("InputData")
    lnglastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lngfirstRow = 2
    Set rng = ws.Range("A" & lngfirstRow & ":A" & lnglastRow)
    arr = rng.value
    fncGetTrades = Application.Transpose(rng.value)
    
    Set ws = Nothing

End Function
