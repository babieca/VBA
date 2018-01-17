Public Function AVol( _
    Prices As Range, _
    Optional Dividends, _
    Optional DataInterval = 1, _
    Optional AnnualTradingDays = 252 _
    ) As Double
    



Dim MyVal As Variant
Dim DailyReturns As Variant
Dim DailyReturnsSquared As Variant

' Set the TimePeriod value to 1.
'(This value is used differently in variants of this function.)
TimePeriod = 1

'Validate data:
    'Range must have at least 3 datapoints (= 2 intervals) for calculating standard deviation
    'NOTE:  DataCount is used later in the function as well.
    DataCount = Application.WorksheetFunction.Count(Prices)
    If DataCount < 3 Then
        AVol = CVErr(xlErrRef)
        Exit Function
    End If
    

    'Data ranges must be only ONE column wide OR one row high
    If Prices.Rows.Count > 1 Then
        If Prices.Columns.Count > 1 Then
            AVol = CVErr(xlErrRef)
            Exit Function
        End If
    End If
    If Not IsMissing(Dividends) Then
        If Dividends.Rows.Count > 1 Then
            If Dividends.Columns.Count > 1 Then
                AVol = CVErr(xlErrRef)
                Exit Function
            End If
        End If
    End If

    'DIV range must be same size as PRX range:
    PricesSize = 0
    For Each MyVal In Prices
        PricesSize = PricesSize + 1
    Next
    If Not IsMissing(Dividends) Then
        DivSize = 0
        For Each MyVal In Dividends
            DivSize = DivSize + 1
        Next
        If DivSize <> PricesSize Then
            AVol = CVErr(xlErrRef)
            Exit Function
        End If
    End If
      
'Deal with AnnualTradingDays.  In the US public markets,
'a typical year has 252 trading days.  Since price movement
'occurs only on trading days, a year is counted as the number
'of trading days rather than 365 days.
'
'Some markets or assets may experience price movements
'for a different number of days.  This optional variable
'allows the user to select an alternative number of
'annual trading days.  If no input is provided, the number
'is 252 by default.

If Not IsMissing(AnnualTradingDays) Then
    If Not IsNumeric(AnnualTradingDays) Then
        AVol = CVErr(xlErrValue)
        Exit Function
    ElseIf AnnualTradingDays > 365 Then
        AVol = CVErr(xlErrValue)
        Exit Function
    End If
End If
      

'Now adjust for the DATA INTERVAL passed to the function -
'i.e., does the Prices range contain daily, weekly, monthly, or other
'data?
'
'DATA INTERVAL can be passed to the function as a NUMBER,
'represented in NUMBER OF TRADING DAYS (i.e., for weekly volatility, pass "5"
'OR, may be passed with a one-character key as follows:
'   A = ANNUAL
'   S = SEMIANNUAL
'   Q = QUARTERLY
'   M = MONTHLY
'   B = BIWEEKLY
'   W = WEEKLY
'   D = DAILY

If Not IsMissing(DataInterval) Then
    If Not IsNumeric(DataInterval) Then
        If Not TypeName(DataInterval) = "String" Then
            AVol = CVErr(xlErrValue)
            Exit Function
        ElseIf Len(DataInterval) > 1 Then
            AVol = CVErr(xlErrValue)
            Exit Function
        Else
            DataInterval = UCase(DataInterval)
            Select Case DataInterval
                Case "A"
                    DataInterval = AnnualTradingDays
                Case "S"
                    DataInterval = 0.5 * AnnualTradingDays
                Case "Q"
                    DataInterval = 0.25 * AnnualTradingDays
                Case "M"
                    DataInterval = AnnualTradingDays / 12
                Case "B"
                    DataInterval = AnnualTradingDays / 26
                Case "W"
                    DataInterval = AnnualTradingDays / 52
                Case "D"
                    DataInterval = 1
                Case Else
                    AVol = CVErr(xlErrValue)
                    Exit Function
            End Select
        End If
    End If
End If




'Now prepare to do the math.
DailyReturns = 0
DailyReturnsSquared = 0


'Find the first data within the Price range variable
i = 1
Do Until Not IsEmpty(PrevVal)
    PrevVal = Prices(i)
    i = i + 1
Loop
FirstData = i


'Now caculate daily returns
TodaysDiv = 0
N = 1
For Each NewVal In Prices
    If N >= FirstData Then
        If Not IsEmpty(Prices(N)) Then
            If IsMissing(Dividends) Then
                TodaysDiv = 0
            Else
                TodaysDiv = Dividends(N)
            End If
            
            ThisDailyReturn = Application.WorksheetFunction.Ln((NewVal + TodaysDiv) / PrevVal)
            DailyReturns = DailyReturns + ThisDailyReturn
            DailyReturnsSquared = DailyReturnsSquared + (ThisDailyReturn ^ 2)
            PrevVal = NewVal
        End If
    End If
    N = N + 1
Next

Part1 = DailyReturnsSquared / (DataCount - 2)
Part2 = DailyReturns ^ 2 / ((DataCount - 1) * (DataCount - 2))

AVol = ((Part1 - Part2) ^ 0.5) * ((TimePeriod * AnnualTradingDays / DataInterval) ^ 0.5)

End Function

Public Function SharpeR( _
    Prices As Range, _
    RiskFree As Variant _
    ) As Double
    
    Dim i, j As Integer

    Dim MyVal As Variant
    Dim ProdDailyReturns As Variant
    Dim RiskFreeT As Variant

    Dim AnnualizedReturn As Double
    Dim AnnualizedRFReturn As Double
    Dim StandardDev As Double
    Dim ExcessReturn As Double
    Dim nValues As Integer
    
    Dim TimePeriod As Double
    Dim AnnualTradingDays As Double
    Dim DataInterval As Double
    
    TimePeriod = 1
    AnnualTradingDays = 252
    DataInterval = AnnualTradingDays / 12
    
    'Now prepare to do the math.
    ProdDailyReturns = 0
    ProdRFReturns = 0
    
    i = 1
    Do Until Not IsEmpty(PrevVal)
        PrevVal = Prices(i)
        i = i + 1
    Loop
    
    k1 = 0
    ProdDailyReturns = 1
    For Each NewVal In Prices
        ThisDailyReturn = (NewVal / PrevVal) - 1
        ProdDailyReturns = ProdDailyReturns * (1 + ThisDailyReturn)
        PrevVal = NewVal
        k1 = k1 + 1
    Next
    k1 = k1 - 1
    
    k2 = 0
    ProdRFReturns = 1
    For Each NewVal In RiskFree
        ProdRFReturns = ProdRFReturns * (1 + NewVal)
        k2 = k2 + 1
    Next
    
    AnnualizedReturn = (ProdDailyReturns ^ (12 / k1)) - 1
    AnnualizedRFReturn = (ProdRFReturns ^ (12 / k2)) - 1
    
    ExcessReturn = AnnualizedReturn - AnnualizedRFReturn
    StandardDev = AVol(Prices, , "M")
    
    SharpeR = ExcessReturn / StandardDev
    
End Function


Function GetLastRow(strSheet, strColum)
 Dim MyRange As Range
 Dim lngLastRow As Long

    Set MyRange = Worksheets(strSheet).Range(strColum & "1")

    GetLastRow = Cells(65536, MyRange.Column).End(xlUp).Row
End Function
