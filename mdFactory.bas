Attribute VB_Name = "mdFactory"
'Purpose: Factory method for creating the simulation object with it's variables already initialized

Option Explicit

Public Function CreateSimulation(ByVal totalRuns As Integer, ByVal lotSize As Integer, ByVal tradesInYear As Integer, ByVal TradeList As Variant, ByVal startEquity As Double, ByVal margin As Double) As clsSimulation

    If (totalRuns = 0) Or (lotSize = 0) And (tradesInYear = 0) Or (UBound(TradeList) = 0) Or (margin = 0) Then
        MsgBox "One or more parameters are invalid!", vbExclamation, "Invalid Inputs"
        Set CreateSimulation = Nothing
    End If
    
    Set CreateSimulation = New clsSimulation
    CreateSimulation.InitiateProperties tradesInYear:=tradesInYear, TradeList:=TradeList, startEquity:=startEquity, margin:=margin, lotSize:=lotSize, totalRuns:=totalRuns
    
End Function
