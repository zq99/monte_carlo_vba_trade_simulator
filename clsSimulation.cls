VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSimulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Purpose: This class encapsulates the entire logic for the Monte Carlo Simulation

Option Explicit


Private mintLotSize As Integer
Private mintTotalRuns As Integer
Private mvntTradeList As Variant
Private mdblStartEquity As Double
Private mdblMargin As Double
Private mdblEquityIncrement As Double
Private mintTradesInYear As Integer

Private Const DEFAULT_RUNS As Integer = 2500
Private Const DEFAULT_LOTSIZE As Integer = 1

Property Get lotSize() As Integer
    lotSize = mintLotSize
End Property


Property Get totalRuns() As Integer
    totalRuns = mintTotalRuns
End Property


Private Sub Class_Initialize()
    mintLotSize = 1
End Sub


Public Sub InitiateProperties(ByVal tradesInYear As Integer, ByVal TradeList As Variant, _
                              ByVal startEquity As Double, ByVal margin As Double, _
                              Optional ByVal lotSize As Integer = DEFAULT_LOTSIZE, _
                              Optional ByVal totalRuns As Integer = DEFAULT_RUNS)

'Purpose: This creates a method that can be used as a factory process to initialize the class

    mintLotSize = lotSize
    mintTotalRuns = totalRuns
    mvntTradeList = TradeList
    mdblMargin = margin
    mdblStartEquity = startEquity
    mdblEquityIncrement = mdblStartEquity / 4
    mintTradesInYear = tradesInYear

End Sub


Public Function fncRunProcess() As Collection

' Purpose: this is the main routine that starts the simulation
' The results are return in a collection of clsResult objects

    Dim dblBeginEquity As Double
    Dim dblEnd As Double
    Dim oResult As clsResult
    Dim collResults As New Collection
    
On Error GoTo Err_Handler
    
    dblEnd = mdblStartEquity + 10 * mdblEquityIncrement
        
    For dblBeginEquity = mdblStartEquity To dblEnd Step mdblEquityIncrement
        
        Set oResult = fncProcessStartEquity(dblBeginEquity, mintTotalRuns)
        If Not oResult Is Nothing Then
            collResults.Add oResult
        End If
        
    Next dblBeginEquity
    
    Set fncRunProcess = collResults

Exit_Here:
    Set collResults = Nothing
    Set oResult = Nothing
    Exit Function
Err_Handler:
    MsgBox Err.Description, vbCritical, "Error"
    GoTo Exit_Here
End Function


Private Function fncProcessStartEquity(ByVal dblBeginEquity As Double, Optional ByVal intTotalRuns As Integer = DEFAULT_RUNS) As clsResult

' Purpose: run the simulation for this starting equity amount

' dblBeginEquity = Base start Equity value for the output
' intTotalRuns = This is the number of times to run the simulation for each equity (default is 2500)

    Dim dblTotalriskofruin As Double
    Dim intIteration As Integer
    Dim oEquityCurve As New clsEquityCurve
    Dim wsCalc As Worksheet
    Dim oResult As New clsResult
    Dim lRow As Long
    
    Const STR_CALC_SHEET As String = "Calc"

    dblTotalriskofruin = 0

    Call DeleteSheetIfExists(STR_CALC_SHEET)
    Set wsCalc = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsCalc.name = STR_CALC_SHEET
    
    With wsCalc
    
        For intIteration = 1 To intTotalRuns
        
            Set oEquityCurve = fncBuildEquityCurve(dblBeginEquity, mintTradesInYear, mvntTradeList, mintLotSize)
            
            If oEquityCurve.IsRuined Then
                dblTotalriskofruin = dblTotalriskofruin + 1
            End If
            
            .Range("A" & intIteration).value = oEquityCurve.EquityAmount
            .Range("B" & intIteration).value = oEquityCurve.Drawdown
            .Range("C" & intIteration).value = oEquityCurve.GetReturn
            .Range("D" & intIteration).value = oEquityCurve.GetReturnOverDrawdown
        
        Next intIteration
        
        dblTotalriskofruin = dblTotalriskofruin / mintTotalRuns
    End With
    
    wsCalc.Calculate
    lRow = wsCalc.Range("A1").Cells.CurrentRegion.Rows.Count
      
    oResult.Equity = dblBeginEquity
    oResult.Ruin = dblTotalriskofruin
    
    oResult.MedianProfit = Application.WorksheetFunction.Median(wsCalc.Range("A1:A" & lRow)) - dblBeginEquity
    oResult.MedianDrawdown = Application.WorksheetFunction.Median(wsCalc.Range("B1:B" & lRow))
    oResult.MedianReturn = Application.WorksheetFunction.Median(wsCalc.Range("C1:C" & lRow))
    oResult.MedianReturnDD = Application.WorksheetFunction.Median(wsCalc.Range("D1:D" & lRow))

    Set fncProcessStartEquity = oResult
    
    Call DeleteSheetIfExists(STR_CALC_SHEET)
    Set oEquityCurve = Nothing
    Set wsCalc = Nothing
    Set oResult = Nothing
    
End Function


Private Function fncBuildEquityCurve(ByVal dblBeginEquity As Double, ByVal intTradesInYear As Integer, _
                                     ByVal vntTradeList As Variant, Optional ByVal intLotSize As Integer = DEFAULT_LOTSIZE) As clsEquityCurve

' Purpose: This takes a list of trade PNL, then it picks at random a trade PNL value in order to construct an equity curve
'          it does this for the number times specified for how many trades are likely to be placed in a year (intTradesInYear)

' dblBeginEquity = this the start of the base to build the equity curve
' inTradesInYear = number of trades you expect to make per year
' vntTradeList = list of trade result PNL for existing trades made (one dimensional array)
' intLotSize = quantity of the asset traded (default is 1 per trade result)

    Dim intTrades As Integer
    Dim dblTradevalue As Double
    Dim intTradenumber As Integer
    Dim oEquityCurve As New clsEquityCurve
    
    oEquityCurve.InitializeStartEquity (dblBeginEquity)
    
    For intTrades = 1 To intTradesInYear
    
        Randomize (Time)
        
        intTradenumber = Application.WorksheetFunction.RandBetween(LBound(vntTradeList), UBound(vntTradeList))
        
        dblTradevalue = vntTradeList(intTradenumber)
         
        If oEquityCurve.EquityAmount < mdblMargin Then
            oEquityCurve.IsRuined = True
            GoTo Exit_Function
        End If
        
        oEquityCurve.Add (intLotSize * dblTradevalue)
    
    Next
    
Exit_Function:
    Set fncBuildEquityCurve = oEquityCurve
    Set oEquityCurve = Nothing
End Function


Private Sub DeleteSheetIfExists(ByVal strName As String)

'Purpose: Removes interim sheets used in the calculation

    Dim sheet As Worksheet
    Dim blnAlerts As Boolean
    
    blnAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    
    For Each sheet In ThisWorkbook.Worksheets
        If sheet.name = strName Then
            sheet.Delete
            Exit For
        End If
    Next sheet
    
    Application.DisplayAlerts = blnAlerts
    Set sheet = Nothing
End Sub


