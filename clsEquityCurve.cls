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

Implements INameProvider

Private mdblMaxEquityAmount As Double
Private mdblEquityAmount As Double
Private mdblDrawdown As Double
Private mIsRuined As Boolean
Private mdblStartEquity As Double
Private marrEquityHistory() As Variant

Private Function INameProvider_GetClassName() As String
    INameProvider_GetClassName = "clsEquityCurve"
End Function

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
    mdblStartEquity = 0
End Sub

Public Sub calculateDrawdown()
    Dim value As Double
    value = 1 - (mdblEquityAmount / mdblMaxEquityAmount)
    If value > mdblDrawdown Then mdblDrawdown = value
End Sub

Public Sub InitializeStartEquity(ByVal dblEquity As Double)
    mdblStartEquity = dblEquity
    mdblEquityAmount = dblEquity
    mdblMaxEquityAmount = dblEquity
    ReDim marrEquityHistory(0 To 0)
    marrEquityHistory(0) = dblEquity
End Sub

Public Sub Add(ByVal dblValue As Double)
    mdblEquityAmount = mdblEquityAmount + dblValue
    Call AddHistory(mdblEquityAmount)
    
    If mdblEquityAmount > mdblMaxEquityAmount Then
        mdblMaxEquityAmount = mdblEquityAmount
    Else
        Me.calculateDrawdown
    End If
End Sub

Public Function GetReturn() As Double
    GetReturn = ((mdblEquityAmount / mdblStartEquity) - 1)
End Function

Public Function GetReturnOverDrawdown() As Double
    If Me.Drawdown <> 0 Then
        GetReturnOverDrawdown = Me.GetReturn() / Me.Drawdown()
    End If
End Function

Private Sub AddHistory(ByVal equity As Double)
    ' Check if the array is empty (not yet initialized)
    If IsEmpty(marrEquityHistory) Then
        ReDim marrEquityHistory(0 To 0) ' Initialize the array with one element
    Else
        ReDim Preserve marrEquityHistory(0 To UBound(marrEquityHistory) + 1) ' Resize the array, preserving existing values and adding one more element
    End If
    
    marrEquityHistory(UBound(marrEquityHistory)) = equity ' Assign the new value to the last element
End Sub


Public Function GetEquityHistory() As Variant
    GetEquityHistory = marrEquityHistory
End Function

Private Sub ClearHistory()
    ReDim marrEquityHistory(0 To 0) ' Initialize the array with one element
End Sub


