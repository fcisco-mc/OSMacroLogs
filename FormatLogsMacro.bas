Attribute VB_Name = "FormatLogsMacro"
Sub FormatLogs()

'Version 2:
    '# Fix: headers with _ characters are replaced with blank spaces
    '# Fix: Some logs were missing headers. Added a resume next activity to avoid the trigger of exception
'V2.1:
    '# Fix: set worksheet to the active (opened) worksheet instead of the hardcoded name "Sheet1"
'V3:
    '# Fix an infinite loop when the headers are not on the first line
    '# One macro for all SC logs
    '# If Autofilter is already applied, then it's not removed
'#V3.1: Adapted to Platform 11 headers

    Dim RowHeaderRange As Range, nameHeader As Range, instantHeader As Range, messageHeader As Range, StackHeader As Range
    Dim moduleNameHeader As Range, RequestKeyHeader As Range, EspaceIdHeader As Range, ActionNameHeader As Range
    Dim endPointHeader As Range, actionHeader As Range, durationHeader As Range, ScreenHeader As Range
    Dim myWorksheet As Worksheet
    
    'Gets rid of annoying message whenever you close the Excel file'
    Application.DisplayAlerts = False
    
    'Error Handler'
    On Error GoTo ErrorHandler:
    
    Set myWorksheet = ActiveWorkbook.ActiveSheet
    
    'Selects all cells and defines its row height'
    With myWorksheet
        .Activate
        .Cells.Select
        .Rows.RowHeight = 14
    End With
    
    'Defines the Headers Range. If no headers are found on the first line, it interrupts the code to avoid an infinite loop
    Set RowHeaderRange = myWorksheet.Cells.Range(Range("A1"), Range("A1").End(xlToRight))
    If RowHeaderRange(1, 1) = "" And RowHeaderRange(RowHeaderRange.Rows.Count, RowHeaderRange.Columns.Count) = "" Then
        GoTo InterruptExecution
    End If
        
    'Some logs have '_' characaters on the headers. The For Each loop replaces them with spaces
    For Each cell In RowHeaderRange.Cells
        cell.Value = Replace(cell.Value, "_", " ")
    Next cell
    
    'On error, it resumes next in case it does not find one or more of the headers below'
    On Error Resume Next
    'Finds Headers by exact match'
    'Common Headers
    Set instantHeader = RowHeaderRange.Cells.Find("Instant", LookAt:=xlWhole, MatchCase:=False)
    instantHeader.EntireColumn.ColumnWidth = 20
    
    Set RequestKeyHeader = RowHeaderRange.Cells.Find("Request Key", LookAt:=xlWhole, MatchCase:=False)
    RequestKeyHeader.EntireColumn.ColumnWidth = 35
    
    Set nameHeader = RowHeaderRange.Cells.Find("Name", LookAt:=xlWhole, MatchCase:=False)
    If nameHeader Is Nothing Then
        Set nameHeader = RowHeaderRange.Cells.Find("Espace Name", LookAt:=xlWhole, MatchCase:=False)
    End If
    nameHeader.EntireColumn.ColumnWidth = 20
    
    'Other Headers (General and Error mostly)
    Set ActionNameHeader = RowHeaderRange.Cells.Find("Action Name", LookAt:=xlWhole, MatchCase:=False)
    ActionNameHeader.EntireColumn.ColumnWidth = 18

    Set messageHeader = RowHeaderRange.Cells.Find("Message", LookAt:=xlWhole, MatchCase:=False)
    messageHeader.EntireColumn.ColumnWidth = 80
    
    Set StackHeader = RowHeaderRange.Cells.Find("Stack", LookAt:=xlWhole, MatchCase:=False)
    StackHeader.EntireColumn.ColumnWidth = 40
    
    Set moduleNameHeader = RowHeaderRange.Cells.Find("Module Name", LookAt:=xlWhole, MatchCase:=False)
    moduleNameHeader.EntireColumn.ColumnWidth = 20
    
    'Integration Headers
    Set endPointHeader = RowHeaderRange.Cells.Find("Endpoint", LookAt:=xlWhole, MatchCase:=False)
    endPointHeader.EntireColumn.ColumnWidth = 90
    
    Set actionHeader = RowHeaderRange.Cells.Find("Action", LookAt:=xlWhole, MatchCase:=False)
    actionHeader.EntireColumn.ColumnWidth = 90
    
    Set durationHeader = RowHeaderRange.Cells.Find("Duration", LookAt:=xlWhole, MatchCase:=False)
    durationHeader.EntireColumn.ColumnWidth = 10
    
    'Screen and Mobile Headers
    Set ScreenHeader = RowHeaderRange.Cells.Find("Screen", LookAt:=xlWhole, MatchCase:=False)
    ScreenHeader.EntireColumn.ColumnWidth = 30
    
    'This find is just to reset the "Exact match" when CTRL+F
    Random = RowHeaderRange.Cells.Find("", LookAt:=xlPart)
    
    'Error Handler 2 to avoid resuming next'
    On Error GoTo ErrorHandler:
    
    'Applies filter to Headers'
    If Not myWorksheet.AutoFilterMode Then
        myWorksheet.Cells.Rows(1).AutoFilter
    End If
    
    'Freezes Top Row'
    Rows("2:2").Select
    ActiveWindow.FreezePanes = True
    
    Exit Sub
    
    'Error Handler'
ErrorHandler:
        MsgBox ("Oops! Something went wrong. Make sure you selected the right type of Macro"), vbCritical
        MsgBox Err.Description
        Exit Sub
        
InterruptExecution:
        MsgBox ("Headers must be on the first line of the Excel file!"), vbCritical
        Exit Sub

End Sub


