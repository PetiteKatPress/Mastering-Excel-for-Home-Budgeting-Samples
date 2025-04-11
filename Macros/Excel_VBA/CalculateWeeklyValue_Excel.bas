Function CalculateWeeklyValue(amount As Double, period As String) As Double
    Dim WeeklyValue As Double
    
    Select Case LCase(period)
        Case "2", "forthnight", "fortnight"
            ' Forthnight/Fortnight (every two weeks)
            WeeklyValue = amount / 2
        Case "4", "month"
            ' Month
            WeeklyValue = (amount * 12) / 52
        Case "8", "bi month", "bimonth", "bi-month", "bi-monthly"
            ' Bi-Month (every two months)
            WeeklyValue = (amount * 6) / 52
        Case "12", "quarter", "quarterly"
            ' Quarter (every three months)
            WeeklyValue = (amount * 4) / 52
        Case "26", "6 month", "six month", "semi-annual", "semiannual"
            ' 6 Months (semi-annual)
            WeeklyValue = (amount * 2) / 52
        Case "52", "year", "annual"
            ' Year (annual)
            WeeklyValue = amount / 52
        Case Else
            ' If the period is not recognized, return 0
            WeeklyValue = 0
    End Select
    
    CalculateWeeklyValue = WeeklyValue
End Function