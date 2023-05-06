Attribute VB_Name = "TestModule_clsResult"
Option Explicit

Public Sub Test_clsResult()
'Purpose: This tests the results class for errors

    Dim objResult As clsResult
    Dim objLogger As clsTestLogger
    
    
    ' Instantiate clsResult object
    Set objResult = New clsResult
    
    ' Instantiate clsTestLogger object
    Set objLogger = New clsTestLogger
    objLogger.SetClass objResult
    
    ' Set property values
    objResult.equity = 10000
    objResult.Ruin = 0.1
    objResult.MedianDrawdown = 500
    objResult.MedianProfit = 2000
    objResult.MedianReturn = 0.2
    objResult.MedianReturnDD = 4
    
    ' Use Debug.Assert to check property values
    Debug.Assert objResult.equity = 10000
    Debug.Assert objResult.Ruin = 0.1
    Debug.Assert objResult.MedianDrawdown = 500
    Debug.Assert objResult.MedianProfit = 2000
    Debug.Assert objResult.MedianReturn = 0.2
    Debug.Assert objResult.MedianReturnDD = 4
    
    ' Clean up
    Set objResult = Nothing
    Set objLogger = Nothing

End Sub

