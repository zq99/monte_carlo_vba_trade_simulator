Attribute VB_Name = "TestModule_clsEquityCurve"
Option Explicit

Public Sub Test_clsEquityCurve()
'Purpose: This tests the equity curve class for errors

    ' Declare variables
    Dim oEquityCurve As clsEquityCurve
    Dim dblStartEquity As Double
    Dim dblValue As Double
    Dim dblReturn As Double
    Dim dblReturnOverDrawdown As Double
    Dim objLogger As clsTestLogger

    ' Initialize variables
    dblStartEquity = 10000
    dblValue = 500
    
    ' Create a new instance of clsEquityCurve
    Set oEquityCurve = New clsEquityCurve
    
    ' Instantiate clsTestLogger object
    Set objLogger = New clsTestLogger
    objLogger.SetClass oEquityCurve
    
    ' Test the InitializeStartEquity method
    oEquityCurve.InitializeStartEquity dblStartEquity
    Debug.Assert oEquityCurve.EquityAmount = dblStartEquity
    Debug.Assert oEquityCurve.MaxEquityAmount = dblStartEquity
    
    ' Test the Add method
    oEquityCurve.Add dblValue
    Debug.Assert oEquityCurve.EquityAmount = dblStartEquity + dblValue
    Debug.Assert oEquityCurve.MaxEquityAmount = dblStartEquity + dblValue
    
    ' Test the GetReturn method
    dblReturn = oEquityCurve.GetReturn
    Debug.Assert dblReturn = (dblValue / dblStartEquity)
    
    ' Test the GetReturnOverDrawdown method
    dblReturnOverDrawdown = oEquityCurve.GetReturnOverDrawdown
    If oEquityCurve.Drawdown = 0 Then
        Debug.Assert dblReturnOverDrawdown = 0
    Else
        Debug.Assert dblReturnOverDrawdown = dblReturn / oEquityCurve.Drawdown
    End If
    
    ' Test the calculateDrawdown method
    oEquityCurve.Add -1000
    oEquityCurve.calculateDrawdown
    Debug.Assert oEquityCurve.Drawdown = 1 - ((dblStartEquity + dblValue - 1000) / (dblStartEquity + dblValue))
    
    ' Test the IsRuined property
    oEquityCurve.IsRuined = True
    Debug.Assert oEquityCurve.IsRuined = True
    
    ' Test history
    ' Initialize start equity
    oEquityCurve.InitializeStartEquity 1000
    
    ' Add values to the equity curve
    oEquityCurve.Add 100
    oEquityCurve.Add -50
    oEquityCurve.Add 200
    
    ' Get the equity history
    Dim equityHistory As Variant
    equityHistory = oEquityCurve.GetEquityHistory
    
    ' Test if the equity history has the expected length
    Debug.Assert oEquityCurve.GetHistoryCount = 4
    
    objLogger.PrintMessage ("output equity history contains the expected values..")
    Dim i As Integer
    For i = LBound(equityHistory) To UBound(equityHistory)
        objLogger.PrintMessage (i & vbTab & equityHistory(i))
    Next
    
    Debug.Assert equityHistory(0) = 1000
    Debug.Assert equityHistory(1) = 1100
    Debug.Assert equityHistory(2) = 1050
    Debug.Assert equityHistory(3) = 1250

    
    ' Clean up
    Set oEquityCurve = Nothing
    Set objLogger = Nothing

End Sub

