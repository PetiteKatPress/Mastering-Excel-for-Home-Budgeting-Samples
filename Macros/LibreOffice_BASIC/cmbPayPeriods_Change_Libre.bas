Sub cmbPayPeriods_Change()
    Dim oDoc As Object
    Dim oSheet As Object
    Dim iSelection As Integer
    Dim iValue As Integer

    oDoc = ThisComponent
    oSheet = oDoc.CurrentController.ActiveSheet

    iSelection = oSheet.getCellRangeByName("M5").Value

    Select Case iSelection
        Case 1
            iValue = 1
        Case 2
            iValue = 12
        Case 3
            iValue = 26
        Case 4
            iValue = 52
        Case Else
            MsgBox "Expected values are 1, 2, 3 or 4", 48, "Invalid Selection"
            Exit Sub
    End Select

    oSheet.getCellRangeByName("N5").Value = iValue
End Sub
