VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEquityCurve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Purpose: This stores the result of a creating an Equity curve

Option Explicit

Private mdblMaxEquityAmount As Double
Private mdblEquityAmount As Double
Private mdblDrawdown As Double
Private mIsRuined As Boolean

Property Get MaxEquityAmount() As Double
    MaxEquityAmount = mdblMaxEquityAmount
End Property

Property Let MaxEquityAmount(ByVal value As Double)
    mdblMaxEquityAmount = value
End Property

Property Get EquityAmount() As Double
    EquityAmount = mdblEquityAmount
End Property

Property Let EquityAmount(ByVal value As Double)
    mdblEquityAmount = value
End Property

Property Get Drawdown() As Double
    Drawdown = mdblDrawdown
End Property


Property Get IsRuined() As Boolean
    IsRuined = mIsRuined
End Property

Property Let IsRuined(ByVal value As Boolean)
    mIsRuined = value
End Property

Private Sub Class_Initialize()
    mIsRuined = False
    mdblDrawdown = 0
    mdblEquityAmount = 0
    mdblMaxEquityAmount = 0
End Sub


Public Sub calculateDrawdown()
    Dim value As Double
    value = 1 - (mdblEquityAmount / mdblMaxEquityAmount)
    If value > mdblDrawdown Then mdblDrawdown = value
End Sub

Public Sub InitializeStartEquity(ByVal dblEquity As Double)
    mdblEquityAmount = dblEquity
    mdblMaxEquityAmount = dblEquity
End Sub

