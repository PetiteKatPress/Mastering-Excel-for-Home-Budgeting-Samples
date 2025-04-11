Public Function CalculateTax(WeeklyIncome As Long, ClaimingTFT As String, Taxable As String) As Double

'1. Details about the function
'--------------------------------------------
'Author:    Thomas Pettit
'Date:      01 June 2024
'Purpose:   Calculate Tax, Excel Training
'--------------------------------------------
'Input parameters:
'   WeeklyIncome:   The income we will calculate tax for, equivalent to one weeks income, whole dollars only
'   ClaimingTFT:    Is the Tx Free Threshold be claimed, "Y" or "N"
'   Taxable:        Is income taxable, "Y" or "N", function will just return a zero if this "Y"
'
'Output:
'   Function returns a Double type value representing the tax payable on the income based on parameters passed
'   and the value recorded in the tax and associated lookup tables
'--------------------------------------------
    
 '2. Setup generic error handler
 On Error GoTo ErrorHandler         'A generic error handler to capture anything that might go wrong
    
    '3. Declare all variables that will be required
    '--- To reference the named ranges tables
    Dim TaxTable As Range           'To refer the Tax Table
    Dim OtherTaxPayable As Range    'To refer to the additional tax table
    Dim MedicareLevyTax As Range    'For the medicare levy
    
    '--- Storeage for deductions
    Dim PreTaxDeductions As Double  'Any deductions to come out of GROSS, reduces taxable income
    Dim PostTaxDeductions As Double 'Any deductions to come out of GROSS, calculate on taxable income
    Dim NETDeductions As Double     'Deductions to come out of NET amount, donations
    Dim MedicareLevy As Double      'How much MedicareLEvy to pay
    Dim MedicareRate As Double      'The Medicare rate used to calculate the levy
    Dim TotalTax As Double          'Total tax payable. includes all extras
    Dim IncomeExcess As Double      'Income over the brackets lower amount
    Dim TFTRate As Double           'This is the amount payable for the taxfree threshold
    
    '--- Pointers for table row position
    Dim Row As Range                'Row of current table will be in here for easy access
    Dim IsFirstRow As Boolean       'Indicates if on first row of table
    
    '--- Tax Bracket variable
    Dim TaxRate As Double           'Brackets tax rate
    Dim LowerLimit As Double        'Brackets lower amount
    Dim UpperLimit As Double        'Brackets upper amount
    Dim BracketTax As Double        'Brackets lower amount
    Dim NetPay As Double            'Net pay, after all deductions
    Dim TaxableIncome As Double     'How much you will pay tax on
    
'--------------------------------------------
'Begin processing
'--------------------------------------------
    
'4. Setup taxTables to the named ranges in the workbook
    Set TaxTable = ThisWorkbook.Names("PAYG_Tax_Table").RefersToRange
    Set OtherTaxPayable = ThisWorkbook.Names("Other_Tax_Payable").RefersToRange
    Set MedicareLevyTax = ThisWorkbook.Names("Medicare_Levy").RefersToRange
    
'5. Initialize variables
    TotalTax = 0
    PreTaxDeductions = 0
    PostTaxDeductions = 0
    NETDeductions = 0
    
    
'6.Check if income is taxable
    If Taxable = "N" Then
        'Leave function, will return 0
        GoTo LeaveFunction
    End If
    
    
'7. Calculate pre tax deductions
    'Add any % based values
    PreTaxDeductions = PreTaxDeductions + (WeeklyIncome * OtherTaxPayable.Rows.Cells(1, 1).Value)
    'and total fixed $ amounts
    PreTaxDeductions = PreTaxDeductions + OtherTaxPayable.Rows.Cells(1, 2).Value
    'Calculate the taxable income
    TaxableIncome = WeeklyIncome - PreTaxDeductions
    'Calculate the TaxFree Threshold amount for this job on TaxableIncome
    
'8. Calculate post tax deductions
    'Add any % based values
    PostTaxDeductions = PostTaxDeductions + (TaxableIncome * OtherTaxPayable.Rows.Cells(1, 3).Value)
    'and total fixed $ amounts
    PostTaxDeductions = PostTaxDeductions + OtherTaxPayable.Rows.Cells(1, 4).Value
    
    
'9.  Loop through each row in the taxTable and calculate the tax payable on taxable income
    IsFirstRow = True
    For Each Row In TaxTable.Rows
        'Assign the rows values to variables, easier to use
        LowerLimit = Row.Cells(1, 3).Value
        UpperLimit = Row.Cells(1, 4).Value
        
        'Test if claiming the Tax-Free Threshold
        If ClaimingTFT = "N" And IsFirstRow = True Then
            'No, get the rate and calculate the amount
            'Get rate from next bracket
            TaxRate = Row.Cells(2, 5).Value
            'Calculate the Bracket amount
            BracketTax = Row.Cells(1, 4).Value * TaxRate
                    
            IsFirstRow = False
        Else
            'Yes, use the value shown in table
            TaxRate = Row.Cells(1, 5).Value
            BracketTax = Row.Cells(1, 7).Value
            IsFirstRow = False
        End If
        
        'Now start checking
        'Test taxable income is more than zero
        If TaxableIncome > 0 Then
            'Test is less than or equal to the upper limit in the current bracket,
            '1000000 means we have reached the end tax brackets so everything earned above
            'this bracket is paid at this brackets rate
            If TaxableIncome <= UpperLimit Or UpperLimit = 1000000 Then
                'This is our top bracket for the income, calculate the tax on remaining income at this rate
                IncomeExcess = TaxableIncome - LowerLimit
                TotalTax = TotalTax + (IncomeExcess * TaxRate)
                
                'Leave the for loop, there is no more to calculate
                Exit For
            Else
                'We still have more income after this bracket so just take the bracketets tax amount
                TotalTax = TotalTax + BracketTax
            End If
        Else
            'Taxable income is zero, return zero and go to the exit section
            TotalTax = 0
            
            'Leave the routine
            GoTo LeaveFunction
        End If

    Next Row
    
'10. Calculate Medicare Levy
    ' Setup Medicare variable from the table values
    MedicareThreshold = MedicareLevyTax.Rows.Cells(1, 1).Value / 52
    MedicareRate = MedicareLevyTax.Rows.Cells(1, 2).Value
    
    'Test if the income is over the threshhold
    If TaxableIncome > MedicareThreshold Then
        'Income is over the threshhold
        'Calculate MedicareLevy on taxable income
        MedicareLevy = (TaxableIncome) * MedicareRate
    Else
        'Below threshold so zero
        MedicareLevy = 0
    End If

'11. Calculate NetPay
    NetPay = TaxableIncome - TotalTax

'12. Calculate post tax deductions
    'Add any % based values
    NETDeductions = NETDeductions + (NetPay * OtherTaxPayable.Rows.Cells(1, 5).Value)
    'and total fixed $ amounts
    NETDeductions = NETDeductions + MedicareLevy + OtherTaxPayable.Rows.Cells(1, 6).Value
    
LeaveFunction:
    '13. Return the total tax + all deductions
    CalculateTax = Round(TotalTax + PostTaxDeductions + NETDeductions, 2)
    
    'Leave the function
    Exit Function
    
ErrorHandler:
    '14. Tell user what went wrong
    MsgBox "An error occurred: " & Err.Description, vbExclamation
    
    'Return a zero, we do not want to return any other value in an error situation
    CalculateTax = 0
    
    'Leave the function
    Exit Function
    
End Function
