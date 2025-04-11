Function CalculateWeeklyValue(amount As Double, period As String) As Double
    Dim WeeklyValue As Double

    Select Case LCase(period)
        Case "2", "forthnight", "fortnight"
            WeeklyValue = amount / 2
        Case "4", "month"
            WeeklyValue = (amount * 12) / 52
        Case "8", "bi month", "bimonth", "bi-month", "bi-monthly"
            WeeklyValue = (amount * 6) / 52
        Case "12", "quarter", "quarterly"
            WeeklyValue = (amount * 4) / 52
        Case "26", "6 month", "six month", "semi-annual", "semiannual"
            WeeklyValue = (amount * 2) / 52
        Case "52", "year", "annual"
            WeeklyValue = amount / 52
        Case Else
            WeeklyValue = 0
    End Select

    CalculateWeeklyValue = WeeklyValue
End Function
