'--- cmbPayPeriods_Change ---
Sub cmbPayPeriods_Change()
    Dim iSelection As String
    Dim iValue As Integer

    iSelection = Worksheets("Expenses - Budget").Range("M5")

    Select Case iSelection
        Case "Year": iValue = 1
        Case "Month": iValue = 12
        Case "Fortnight": iValue = 26
        Case "Week": iValue = 52
        Case Else
            MsgBox "Unexpected value selected. Please choose Year, Month, Fortnight, or Week.", vbExclamation
            Exit Sub
    End Select

    Worksheets("Expenses - Budget").Range("N5") = iValue
End Sub
