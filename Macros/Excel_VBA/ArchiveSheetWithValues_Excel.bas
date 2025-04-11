Sub ArchiveSheetWithValues()
'1. Details about the function
'--------------------------------------------
'Author:    Thomas Pettit
'Date:      01 June 2024
'Purpose:   This routine will archive the current worksheet, Excel Training
'--------------------------------------------
'Input parameters: None, work on active worksheet
'Output: None, it is a sub
'--------------------------------------------
    '2. Declare variables
    Dim ws As Worksheet             'The Active worksheet
    Dim WorkSheetName As String     'Name of Active worksheet
    Dim wsNew As Worksheet          'The new worksheet
    Dim NewName As String           'New name for new worksheet
    Dim shp As Shape                'Pointer to controls on a worksheet (button)
    
    '3. Get the name of the active worksheet
    WorkSheetName = ActiveSheet.Name
    
    '4.  Set reference to the current worksheet so we can use a short name
    Set ws = ThisWorkbook.Sheets(WorkSheetName)
    
    '5. Generate new name based on current date, 31 characters limit
    NewName = WorkSheetName & "_" & Format(Now, "mmm_yyyy")
    
    '6. Copy the worksheet, set a pointer and rename it
    ws.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)   'Make it the last tab
    Set wsNew = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)      'Assign pointer to new ws
    wsNew.Name = NewName                                            'Rename it
    
    '7. Replace formulas with values across the entire copied sheet
    wsNew.Cells.Copy                                 'Copy all cells of the new worksheet
    wsNew.Cells.PasteSpecial Paste:=xlPasteValues    'Paste back over them but values only
    Application.CutCopyMode = False    'Remove values in copy/paste so we don't paste it anywhere
    
    '8. Disable or delete buttons, Cycle through the controls on ws, these include the button
    For Each shp In wsNew.Shapes
        'Test if it is a control
        If shp.Type = msoFormControl Then
            'Test if it is the button
            If shp.FormControlType = xlButtonControl Then
                '8.a To disable the button
                'shp.OLEFormat.Object.Enabled = False
                
                '8.b Delete the button
                 shp.Delete
            End If
        End If
    Next shp
    
    '9. Save the workbook
    ThisWorkbook.Save
    End Sub
