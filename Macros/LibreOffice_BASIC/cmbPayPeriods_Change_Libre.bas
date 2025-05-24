'Attribute VBA_ModuleType=VBAModule
'The following line MUST be at the top of each module in Libre to provide VBA Support
Option VBASupport 1
' VBA Module - LibreOffice Specific Version

'--- cmbPayPeriods_Change - LibreOffice Version ---
' Tested: April 13, 2025 - Working

Sub cmbPayPeriods_Change()

	'--- Variable Declarations ---
    Dim iSelection As String
    Dim iValue As Integer

	'--- Set Direct Range References ---
    iSelection = Worksheets("Expenses - Budget").Range("M5")

    '--- Examine users selection, store pay periods in variable ---
    Select Case iSelection
        Case "Year": iValue = 1
        Case "Month": iValue = 12
        Case "Fortnight": iValue = 26
        Case "Week": iValue = 52
        Case "Quarter": iValue = 4
        Case "BiAnnual": iValue = 2
        Case Else
            MsgBox "Unexpected value selected. Please choose Year,BiAnnual, Quarter, Month, Fortnight, or Week.", vbExclamation
            Exit Sub
    End Select

	'--- Update worksheet with pay periods, will auto trigger a recalculation in sheet ---
    Worksheets("Expenses - Budget").Range("N5") = iValue
    
End Sub
