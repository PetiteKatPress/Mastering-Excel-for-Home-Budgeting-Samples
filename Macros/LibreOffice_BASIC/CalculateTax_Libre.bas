'--- CalculateTax Function - LibreOffice Version ---
' Tested: April 13, 2025 - Working

Public Function CalculateTax(WeeklyIncome As Long, ClaimingTFT As String, Taxable As String) As Double
    On Error GoTo ErrorHandler

    '--- Variable Declarations ---
    Dim TaxTable As Range
    Dim OtherTaxPayable As Range
    Dim MedicareLevyTax As Range
    Dim PreTaxDeductions As Double
    Dim PostTaxDeductions As Double
    Dim NETDeductions As Double
    Dim MedicareLevy As Double
    Dim MedicareRate As Double
    Dim TotalTax As Double
    Dim IncomeExcess As Double
    Dim TFTRate As Double
    Dim Row As Range
    Dim IsFirstRow As Boolean
    Dim TaxRate As Double
    Dim LowerLimit As Double
    Dim UpperLimit As Double
    Dim BracketTax As Double
    Dim NetPay As Double
    Dim TaxableIncome As Double

    '--- Set Direct Range References ---
    Set TaxTable = Worksheets("Lookup Tables").Range("B10:H14")
    Set OtherTaxPayable = Worksheets("Lookup Tables").Range("C26:H26")
    Set MedicareLevyTax = Worksheets("Lookup Tables").Range("B30:C30")

    '--- Initialize Totals ---
    TotalTax = 0
    PreTaxDeductions = 0
    PostTaxDeductions = 0
    NETDeductions = 0

    '--- Check if Income is Taxable ---
    If Taxable = "N" Then GoTo LeaveFunction

    '--- Calculate Pre-Tax Deductions ---
    PreTaxDeductions = WeeklyIncome * OtherTaxPayable.Cells(1, 1).Value
    PreTaxDeductions = PreTaxDeductions + OtherTaxPayable.Cells(1, 2).Value
    TaxableIncome = WeeklyIncome - PreTaxDeductions

    '--- Calculate Post-Tax Deductions ---
    PostTaxDeductions = TaxableIncome * OtherTaxPayable.Cells(1, 3).Value
    PostTaxDeductions = PostTaxDeductions + OtherTaxPayable.Cells(1, 4).Value

    '--- Determine Tax Based on Brackets ---
    IsFirstRow = True
    For Each Row In TaxTable.Rows
        LowerLimit = Row.Cells(1, 3).Value
        UpperLimit = Row.Cells(1, 4).Value

        If ClaimingTFT = "N" And IsFirstRow Then
            TaxRate = Row.Cells(2, 5).Value
            BracketTax = Row.Cells(1, 4).Value * TaxRate
            IsFirstRow = False
        Else
            TaxRate = Row.Cells(1, 5).Value
            BracketTax = Row.Cells(1, 7).Value
            IsFirstRow = False
        End If

        If TaxableIncome > 0 Then
            If TaxableIncome <= UpperLimit Or UpperLimit = 1000000 Then
                IncomeExcess = TaxableIncome - LowerLimit
                TotalTax = TotalTax + (IncomeExcess * TaxRate)
                Exit For
            Else
                TotalTax = TotalTax + BracketTax
            End If
        Else
            TotalTax = 0
            GoTo LeaveFunction
        End If
    Next Row

    '--- Calculate Medicare Levy ---
    MedicareThreshold = MedicareLevyTax.Cells(1, 1).Value / 52
    MedicareRate = MedicareLevyTax.Cells(1, 2).Value

    If TaxableIncome > MedicareThreshold Then
        MedicareLevy = TaxableIncome * MedicareRate
    Else
        MedicareLevy = 0
    End If

    '--- Calculate Final NET Deductions ---
    NetPay = TaxableIncome - TotalTax
    NETDeductions = (NetPay * OtherTaxPayable.Cells(1, 5).Value)
    NETDeductions = NETDeductions + MedicareLevy + OtherTaxPayable.Cells(1, 6).Value

'--- Normal Exit ---
LeaveFunction:
    CalculateTax = Round(TotalTax + PostTaxDeductions + NETDeductions, 2)
    Exit Function

'--- Error Handler ---
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbExclamation
    CalculateTax = 0
    Exit Function
End Function
