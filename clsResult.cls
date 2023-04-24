VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Purpose: This stores the result of running a simulation for Equity value N times

Option Explicit

Private mdblEquity As Double
Private mdblRuin As Double
Private mdblMedianDrawdown As Double
Private mdblMedianProfit As Double
Private mdblMedianReturn As Double
Private mdblMedianReturnDD As Double

' Equity
Property Let Equity(value As Double)
    mdblEquity = value
End Property

Property Get Equity() As Double
    Equity = mdblEquity
End Property

' Ruin
Property Let Ruin(value As Double)
    mdblRuin = value
End Property

Property Get Ruin() As Double
    Ruin = mdblRuin
End Property

' Median Drawdown
Property Let MedianDrawdown(value As Double)
    mdblMedianDrawdown = value
End Property

Property Get MedianDrawdown() As Double
    MedianDrawdown = mdblMedianDrawdown
End Property

' Median Profit
Property Let MedianProfit(value As Double)
    mdblMedianProfit = value
End Property

Property Get MedianProfit() As Double
    MedianProfit = mdblMedianProfit
End Property

' Median Return
Property Let MedianReturn(value As Double)
    mdblMedianReturn = value
End Property

Property Get MedianReturn() As Double
    MedianReturn = mdblMedianReturn
End Property

' Median Return Drawdown
Property Let MedianReturnDD(value As Double)
    mdblMedianReturnDD = value
End Property

Property Get MedianReturnDD() As Double
    MedianReturnDD = mdblMedianReturnDD
End Property

