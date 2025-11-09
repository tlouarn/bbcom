Option Explicit

Public Function Test()

    Dim ticker As String
    Dim tradeDate As Date
    Dim adjustedPrice As Double
    Dim unadjustedPrice As Double
    
    ticker = "CON GY"
    tradeDate = #9/17/2025#
    
    adjustedPrice = GetPrice(ticker, tradeDate)
    unadjustedPrice = GetPrice(ticker, tradeDate, False)
   
    If adjustedPrice <> 54.598 Or unadjustedPrice <> 72.98 Then
        Err.Raise vbObjectError, "Test", "Tests not passed"
    End If

End Function


Public Function GetPrice(ByVal ticker As String, ByVal tradeDate As Date, Optional ByVal adjusted As Boolean = True) As Double

    Dim wrapper As New bbcom
    
    Dim tickers(1 To 1) As String
    Dim fields(1 To 1) As String
    Dim res() As Variant
    
    tickers(1) = ticker & " Equity"
    fields(1) = "PX_LAST"
    
    If adjusted Then
        res = wrapper.historicalData(tickers, fields, tradeDate, tradeDate, , , , , , , , , False, True, True, False)
    Else
        res = wrapper.historicalData(tickers, fields, tradeDate, tradeDate, , , , , , , , , False, False, False, False)
    End If
    
    GetPrice = res(1, 2)(1)

End Function

