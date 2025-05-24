'--- ArchiveSheetWithValues_Libre.bas ---
' Tested: May 24, 2025 - Working

Sub ArchiveSheetWithValues()
    Dim oDoc As Object
    Dim oSheets As Object
    Dim oSheet As Object
    Dim oNewSheet As Object
    Dim sSheetName As String, sNewName As String
    Dim nCopyIndex As Integer
    Dim row As Long, col As Long
    Dim oCell As Object, oNewCell As Object
    Dim maxRows As Long, maxCols As Long
    Dim oDrawPage As Object, oShape As Object
    Dim i As Integer
    Dim oCursor As Object
	Dim archiveMsg As String

    oDoc = ThisComponent
    oSheets = oDoc.Sheets
    oSheet = oDoc.CurrentController.ActiveSheet
    

	
	' Create a cursor to detect the used range
	oCursor = oSheet.createCursor()
	oCursor.gotoEndOfUsedArea(False)
	maxRows = oCursor.RangeAddress.EndRow
	maxCols = oCursor.RangeAddress.EndColumn
 
    
    sSheetName = oSheet.Name
    sNewName = sSheetName & "_" & Format(Now, "YYYY_MMM_DD")

    ' Ensure unique name
    nCopyIndex = 1
    Do While oSheets.hasByName(sNewName)
        sNewName = sSheetName & "_" & Format(Now, "YYYY_MMM_DD") & "_Copy" & nCopyIndex
        nCopyIndex = nCopyIndex + 1
    Loop

    ' Create new blank sheet
    oSheets.insertNewByName(sNewName, oSheets.getCount())
    oNewSheet = oSheets.getByName(sNewName)

    ' Get size of source sheet
	oCursor = oSheet.createCursor()
	oCursor.gotoEndOfUsedArea(False)
	maxRows = oCursor.RangeAddress.EndRow
	maxCols = oCursor.RangeAddress.EndColumn

	
	archiveMsg = "Archived copy of '" & sSheetName & "' created on " & Format(Now, "YYYY-MMM-DD at HH:MM AM/PM")
	
	oNewSheet.getCellByPosition(0, 0).String = archiveMsg
	oNewSheet.Rows.getByIndex(0).Height = 500 ' taller row for readability


    ' Copy resolved values only
    For row = 0 To maxRows
        For col = 0 To maxCols
            oCell = oSheet.getCellByPosition(col, row)
            oNewCell = oNewSheet.getCellByPosition(col, row + 1)

            ' Set value or string directly — no formulas
            If oCell.Type = com.sun.star.table.CellContentType.VALUE Then
                oNewCell.Value = oCell.Value
            ElseIf oCell.Type = com.sun.star.table.CellContentType.TEXT Then
                oNewCell.String = oCell.String
            ElseIf oCell.Type = com.sun.star.table.CellContentType.FORMULA Then
                If oCell.FormulaResultType = com.sun.star.table.CellContentType.VALUE Then
                    oNewCell.Value = oCell.Value
                ElseIf oCell.FormulaResultType = com.sun.star.table.CellContentType.TEXT Then
                    oNewCell.String = oCell.String
                End If
            End If
            
           	With oNewCell
			    .CharWeight = oCell.CharWeight
			    .CharPosture = oCell.CharPosture
			    .CharFontName = oCell.CharFontName
			    .CharHeight = oCell.CharHeight
			    .IsTextWrapped = oCell.IsTextWrapped
			    .HoriJustify = oCell.HoriJustify
			    .VertJustify = oCell.VertJustify
			    .NumberFormat = oCell.NumberFormat
			    .CellBackColor = oCell.CellBackColor
			End With
 
            
            
        Next col
    Next row

    ' Remove all drawing shapes (buttons etc)
    oDrawPage = oNewSheet.DrawPage
    For i = oDrawPage.getCount() - 1 To 0 Step -1
        oShape = oDrawPage.getByIndex(i)
        oDrawPage.remove(oShape)
    Next i

    ' Remove form controls if any
    If oNewSheet.DrawPage.Forms.hasElements() Then
        Dim oForms As Object
        oForms = oNewSheet.DrawPage.Forms
        For i = oForms.Count - 1 To 0 Step -1
            oForms.removeByIndex(i)
        Next i
    End If

    oDoc.store()
End Sub
