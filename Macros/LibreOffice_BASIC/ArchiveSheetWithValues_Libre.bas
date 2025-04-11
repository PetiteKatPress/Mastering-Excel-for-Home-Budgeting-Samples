Sub ArchiveSheetWithValues()
    Dim oDoc As Object
    Dim oSheets As Object
    Dim oSheet As Object
    Dim oNewSheet As Object
    Dim sSheetName As String
    Dim sNewName As String
    Dim oDispatcher As Object

    oDoc = ThisComponent
    oSheets = oDoc.Sheets
    oSheet = oDoc.CurrentController.ActiveSheet

    sSheetName = oSheet.Name
    sNewName = sSheetName & "_" & Format(Now, "YYYY_MMM_DD")

    oSheet.copyByName(sSheetName, sNewName, oSheets.getCount())

    oNewSheet = oSheets.getByName(sNewName)

    ' Copy all cells and paste as values
    oDispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
    oDoc.CurrentController.select(oNewSheet)
    oDispatcher.executeDispatch(oDoc.CurrentController.Frame, ".uno:SelectAll", "", 0, Array())
    oDispatcher.executeDispatch(oDoc.CurrentController.Frame, ".uno:Copy", "", 0, Array())
    oDispatcher.executeDispatch(oDoc.CurrentController.Frame, ".uno:PasteSpecial", "", 0, Array())

    ' Save changes
    oDoc.store()
End Sub
