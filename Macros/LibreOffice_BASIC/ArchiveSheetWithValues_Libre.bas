'--- ArchiveSheetWithValues_Libre.bas ---
' Tested: April 13, 2025 - Working

Sub ArchiveSheetWithValues()
    Dim oDoc As Object
    Dim oSheets As Object
    Dim oSheet As Object
    Dim oNewSheet As Object
    Dim sSheetName As String
    Dim sNewName As String
    Dim oDispatcher As Object
    Dim oDrawPage As Object
    Dim oShape As Object
    Dim nCopyIndex As Integer

    oDoc = ThisComponent
    oSheets = oDoc.Sheets
    oSheet = oDoc.CurrentController.ActiveSheet

    sSheetName = oSheet.Name
    sNewName = sSheetName & "_" & Format(Now, "YYYY_MMM_DD")

    ' Handle duplicate sheet names automatically
    nCopyIndex = 1
    Do While oSheets.hasByName(sNewName)
        sNewName = sSheetName & "_" & Format(Now, "YYYY_MMM_DD") & "_Copy" & nCopyIndex
        nCopyIndex = nCopyIndex + 1
    Loop

    ' Copy the active sheet
    oSheets.copyByName(sSheetName, sNewName, oSheets.getCount())
    oNewSheet = oSheets.getByName(sNewName)

    ' Clear content first to prevent "Replace existing" popup
    oNewSheet.getCellRangeByPosition(0, 0, oNewSheet.Columns.Count - 1, oNewSheet.Rows.Count - 1).clearContents(7)

    ' Select the new sheet and paste as values
    oDispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
    oDoc.CurrentController.select(oNewSheet)
    oDispatcher.executeDispatch(oDoc.CurrentController.Frame, ".uno:SelectAll", "", 0, Array())
    oDispatcher.executeDispatch(oDoc.CurrentController.Frame, ".uno:Copy", "", 0, Array())
    oDispatcher.executeDispatch(oDoc.CurrentController.Frame, ".uno:PasteOnlyValue", "", 0, Array())

    ' Remove buttons from the new sheet
    oDrawPage = oNewSheet.DrawPage
    Dim i As Integer
    For i = oDrawPage.getCount() - 1 To 0 Step -1
        oShape = oDrawPage.getByIndex(i)
        If oShape.SupportsService("com.sun.star.drawing.ControlShape") Then
            Dim oControl As Object
            oControl = oShape.Control
            If Not IsNull(oControl) Then
                If oControl.SupportsService("com.sun.star.form.component.CommandButton") Then
                    oDrawPage.remove(oShape)
                End If
            End If
        End If
    Next i

    ' Save changes
    oDoc.store()
End Sub
