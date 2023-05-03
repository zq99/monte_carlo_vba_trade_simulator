Attribute VB_Name = "TestModule_All"
Option Explicit


Public Sub RunAllTests()
'**************************************************************************

' Purpose: Run this routine to test all the classes are working correctly

'**************************************************************************
    
    Call TestModule_clsResult.Test_clsResult
    Call TestModule_clsEquityCurve.Test_clsEquityCurve
    Call TestModule_clsSimulation.Test_clsSimulation

End Sub

