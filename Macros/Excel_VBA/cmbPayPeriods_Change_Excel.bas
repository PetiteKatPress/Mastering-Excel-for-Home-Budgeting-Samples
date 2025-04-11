Sub cmbPayPeriods_Change()
    'This routine will determine the selected pay period and
    'put the relevant value directly into the Annual Payments cell and from there
    'any cells referencing it will automatically recalculate, magic.
    'This value is then used in all pay period based calculations

    
    '1. Declare the variables we will need
    Dim iSelection As Integer   'To store the users selection
    Dim iValue As Integer       'To store the value we will be passing back to worksheet
    
    '2. Assign selection to a variable
    iSelection = Worksheets("Expenses - Budget").Range("M5")
    
    '3. Check what user selected and assign the proper amount of pay periods per year to iValue
    Select Case iSelection
        Case 1  'Year
            iValue = 1
        Case 2  'Month
            iValue = 12
        Case 3  'Fortnight
            iValue = 26
        Case 4  'Week
            iValue = 52
        Case Else
            'Handle unexpected values by messaging user and then leaving.
            'We call this using a dropdown so this should never happen.
            'Just being safe
            MsgBox "Expected values are 1, 2, 3 or 4", vbExclamation
            Exit Sub
    End Select
    
    '4. Return value to the worksheets N5 directly and leave
    Worksheets("Expenses - Budget").Range("N5") = iValue
    Exit Sub
    
End Sub
