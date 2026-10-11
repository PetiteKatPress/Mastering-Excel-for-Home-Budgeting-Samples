Option Explicit

Public Sub ArchiveSheetWithValues()

'1. Details about the subroutine
'--------------------------------------------
'Author:    Thomas Pettit
'Date:      01 June 2024, revised 8 October 2026
'Purpose:   Archive the active worksheet, normally Expense - Actuals
'--------------------------------------------
'Input parameters:
'   None
'
'Output:
'   Creates a values-only copy of the active worksheet; formulas are removed.
'   The new sheet uses the original worksheet name with the date appended.
'   Duplicate same-day archive names are rejected to protect existing archives.
'--------------------------------------------

'2. Declare variables
    Dim ws As Worksheet, wsNew As Worksheet, item As Object
    Dim originalSheets As Collection, newName As String, shpIndex As Long
    Dim oldEvents As Boolean, oldAlerts As Boolean, oldUpdating As Boolean
    Dim settingsCaptured As Boolean, saveAttempted As Boolean, completed As Boolean
    Dim failureNumber As Long, failureText As String, cleanupText As String
    Dim report As String, candidate As Object, newCount As Long

    'Set up error handler
    On Error GoTo Failed 

'3. Validate that the active sheet is a worksheet in this workbook
    If Not TypeOf ActiveSheet Is Worksheet Then Err.Raise vbObjectError + 2001, , "Select a worksheet in this workbook."
    If Not (ActiveSheet.Parent Is ThisWorkbook) Then Err.Raise vbObjectError + 2002, , "Select a worksheet in this workbook."
    Set ws = ActiveSheet

'4. Prepare and validate the name for the new worksheet    
    newName = ws.Name & "_" & Format(Now, "yyyy mmm dd")
    ' Both guards run BEFORE copying; chart-sheet names also reserve names.
    If Len(newName) > 31 Then Err.Raise vbObjectError + 2003, , "Archive name exceeds 31 characters: " & newName
    Set originalSheets = New Collection
    For Each item In ThisWorkbook.Sheets
        If StrComp(item.Name, newName, vbTextCompare) = 0 Then Err.Raise vbObjectError + 2004, , "A sheet already has this name: " & newName
        originalSheets.Add item
    Next item

'5. Check the workbook status: read-only state, saved path, and structure protection
'   The source worksheet’s own protection is handled later, after copying.

    If ThisWorkbook.ReadOnly Then Err.Raise vbObjectError + 2005, , "Workbook is read-only."
    If Len(ThisWorkbook.Path) = 0 Then Err.Raise vbObjectError + 2006, , "Save this workbook before archiving."
    If ThisWorkbook.ProtectStructure Then Err.Raise vbObjectError + 2007, , "Workbook structure is protected."

'6. Prepare the environment
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    oldUpdating = Application.ScreenUpdating
    settingsCaptured = True
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

'7. Copy the source worksheet
    ws.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Set wsNew = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    If WasOriginalSheet(wsNew, originalSheets) Then Err.Raise vbObjectError + 2008, , "Cannot identify the newly copied worksheet safely."

    'Assign the new name
    wsNew.Name = newName

    ' Unprotect ONLY the new copy. Password-protected sources are unsupported.
    If wsNew.ProtectContents Or wsNew.ProtectDrawingObjects Or wsNew.ProtectScenarios Then wsNew.Unprotect Password:=""
    With wsNew.UsedRange
        .Value2 = .Value2
    End With

    'Find and remove any Form Control buttons from the newly copied sheet
    For shpIndex = wsNew.Shapes.Count To 1 Step -1
        With wsNew.Shapes(shpIndex)
            If .Type = msoFormControl Then
                If .FormControlType = xlButtonControl Then .Delete
            End If
        End With
    Next shpIndex

    'Reapply the source worksheet’s protection settings to the archive
    If ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios Then
        With ws.Protection
            wsNew.Protect Password:="", DrawingObjects:=ws.ProtectDrawingObjects, _
                Contents:=ws.ProtectContents, Scenarios:=ws.ProtectScenarios, _
                AllowFormattingCells:=.AllowFormattingCells, AllowFormattingColumns:=.AllowFormattingColumns, _
                AllowFormattingRows:=.AllowFormattingRows, AllowInsertingColumns:=.AllowInsertingColumns, _
                AllowInsertingRows:=.AllowInsertingRows, AllowInsertingHyperlinks:=.AllowInsertingHyperlinks, _
                AllowDeletingColumns:=.AllowDeletingColumns, AllowDeletingRows:=.AllowDeletingRows, _
                AllowSorting:=.AllowSorting, AllowFiltering:=.AllowFiltering, AllowUsingPivotTables:=.AllowUsingPivotTables
        End With
        wsNew.EnableSelection = ws.EnableSelection
    End If

    'Try to save the workbook with the archived worksheet
    saveAttempted = True
    ThisWorkbook.Save

    'Did it save?

    'No, report it
    If Not ThisWorkbook.Saved Then Err.Raise vbObjectError + 2009, , "Excel did not confirm that the workbook was saved."
    
    'Otherwise, Yes...report that too
    completed = True
    report = "Archive saved successfully: " & newName & vbCrLf & ThisWorkbook.FullName
    
    'and finish up
    GoTo Finish

'8. Error handling routine
Failed:
    'Get the error number and text description
    failureNumber = Err.Number
    failureText = Err.Description
    ' Never select a cleanup target by name. Identify the new object against
    ' the pre-copy snapshot, including a Copy failure before assignment.
    
    'This will cause errors in this routine to be ignored so we dont get into an endless loop
    On Error Resume Next

    'Run some checks and assign the message to cleanupText based on the error
    If settingsCaptured Then
        If wsNew Is Nothing Then
            For Each item In ThisWorkbook.Sheets
                If Not WasOriginalSheet(item, originalSheets) Then
                    newCount = newCount + 1
                    Set candidate = item
                End If
            Next item
            If newCount = 1 Then
                If TypeOf candidate Is Worksheet Then Set wsNew = candidate
            End If
            If newCount > 0 And wsNew Is Nothing Then cleanupText = "INCOMPLETE: newly added sheets could not be identified safely; retained for inspection."
        End If
        If Not wsNew Is Nothing Then
            If Not WasOriginalSheet(wsNew, originalSheets) Then
                If ThisWorkbook.ProtectStructure Then
                    cleanupText = "INCOMPLETE: new partial archive remains because workbook structure is protected."
                Else
                    Err.Clear
                    Application.DisplayAlerts = False
                    wsNew.Delete
                    If Err.Number = 0 Then
                        cleanupText = "The new partial archive was removed from memory."
                    Else
                        cleanupText = "INCOMPLETE: new partial archive remains. Cleanup error " & Err.Number & ": " & Err.Description
                    End If
                End If
            Else
                cleanupText = "Cleanup refused: target was an existing sheet."
            End If
        End If
        ws.Activate
    End If

    On Error GoTo 0
    report = "Archive NOT completed. Error " & failureNumber & ": " & failureText
    If Len(cleanupText) > 0 Then report = report & vbCrLf & cleanupText
    If saveAttempted Then report = report & vbCrLf & "Saving did not complete successfully. Disk state is unconfirmed; no cleanup save was attempted."
    If Not ThisWorkbook.Saved Then report = report & vbCrLf & "UNSAVED workbook changes remain. Review before saving or closing."

    'Continue to Finish to restore Excel application settings and show the result.

Finish:
    'Restore the original Excel application settings.
    'The source worksheet and Excel security settings are not changed.

    On Error Resume Next
    If settingsCaptured Then
        Err.Clear
        Application.ScreenUpdating = oldUpdating
        Application.DisplayAlerts = oldAlerts
        Application.EnableEvents = oldEvents
        If Err.Number <> 0 Then report = report & vbCrLf & "WARNING: application settings restoration failed: " & Err.Description
    End If
    On Error GoTo 0
    MsgBox report, IIf(completed, vbInformation, vbExclamation)
End Sub

Private Function WasOriginalSheet(ByVal target As Object, ByVal originals As Collection) As Boolean

    'Returns True if the sheet existed before the archive process began.
    Dim item As Object
    For Each item In originals
        If item Is target Then
            WasOriginalSheet = True
            Exit Function
        End If
    Next item
End Function

