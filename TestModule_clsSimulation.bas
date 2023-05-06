Attribute VB_Name = "TestModule_clsSimulation"
Option Explicit

Public Sub Test_clsSimulation()
'Purpose: This tests the main simulation class for errors

    Dim objSimulation As clsSimulation
    Dim objResults As Collection
    Dim objResult As clsResult
    Dim vntTradeList As Variant
    Dim intTradesInYear As Integer
    Dim dblStartEquity As Double
    Dim dblMargin As Double
    Dim intLotSize As Integer
    Dim intTotalRuns As Integer
    Dim objLogger As clsTestLogger
    
    ' Initialize input data for the simulation
    intTradesInYear = 100
    vntTradeList = Array(10, 20, -10, -15, 30, 40, -20, 50, 10, -30)
    dblStartEquity = 10000
    dblMargin = 1000
    intLotSize = 1
    intTotalRuns = 2500
    
    ' Instantiate clsSimulation object
    Set objSimulation = New clsSimulation
    
    ' Instantiate clsTestLogger object
    Set objLogger = New clsTestLogger
    objLogger.SetClass objSimulation
    
    ' Initialize properties for the simulation
    objSimulation.InitiateProperties intTradesInYear, vntTradeList, dblStartEquity, dblMargin, intLotSize, intTotalRuns
    
    ' Run simulation process
    Set objResults = objSimulation.fncRunProcess
    
    Debug.Assert objResults.Count > 0
    
    ' Display results
    objLogger.PrintMessage "Equity", "Risk of Ruin", "Median Profit", "Median Drawdown", "Median Return", "Median Return/DD"
    For Each objResult In objResults
        objLogger.PrintMessage objResult.equity, objResult.Ruin, objResult.MedianProfit, objResult.MedianDrawdown, objResult.MedianReturn, objResult.MedianReturnDD
    Next objResult
    
    ' Clean up
    Set objResult = Nothing
    Set objResults = Nothing
    Set objSimulation = Nothing
    Set objLogger = Nothing

End Sub

