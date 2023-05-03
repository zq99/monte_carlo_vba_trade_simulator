Attribute VB_Name = "TestModule_clsSimulation"
Option Explicit

Public Sub Test_clsSimulation()
'Purpose: This tests the main simulation class for errors

    Dim objSimulation As clsSimulation
    Dim objResults As Collection
    Dim objResult As clsResult
    Dim TradeList As Variant
    Dim tradesInYear As Integer
    Dim startEquity As Double
    Dim margin As Double
    Dim lotSize As Integer
    Dim totalRuns As Integer
    Dim objLogger As clsTestLogger
    
    ' Initialize input data for the simulation
    tradesInYear = 100
    TradeList = Array(10, 20, -10, -15, 30, 40, -20, 50, 10, -30)
    startEquity = 10000
    margin = 1000
    lotSize = 1
    totalRuns = 2500
    
    ' Instantiate clsSimulation object
    Set objSimulation = New clsSimulation
    
    ' Instantiate clsTestLogger object
    Set objLogger = New clsTestLogger
    objLogger.SetClass objSimulation
    
    ' Initialize properties for the simulation
    objSimulation.InitiateProperties tradesInYear, TradeList, startEquity, margin, lotSize, totalRuns
    
    ' Run simulation process
    Set objResults = objSimulation.fncRunProcess
    
    Debug.Assert objResults.Count > 0
    
    ' Display results
    Debug.Print "Equity", "Risk of Ruin", "Median Profit", "Median Drawdown", "Median Return", "Median Return/DD"
    For Each objResult In objResults
        Debug.Print objResult.equity, objResult.Ruin, objResult.MedianProfit, objResult.MedianDrawdown, objResult.MedianReturn, objResult.MedianReturnDD
    Next objResult
    
    ' Clean up
    Set objResult = Nothing
    Set objResults = Nothing
    Set objSimulation = Nothing
    Set objLogger = Nothing

End Sub

