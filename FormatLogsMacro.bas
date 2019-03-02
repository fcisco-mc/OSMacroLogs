Attribute VB_Name = "FormaLogsMacro"
Sub FormatErrorLogs()

'Version 2:
    '# Fix: headers with _ characters are replaced with blank spaces
    '# Fix: Some logs were missing headers. Added a resume next activity to avoid the trigger of exception
'V2.1:
    '# Fix: set worksheet to the active (opened) worksheet instead of the hardcoded name "Sheet1"
'V3:
    '# Fix an infinite loop when the headers are not on the first line
    '# One macro for all SC logs
    '# If Autofilter is already applied, then it's not removed

    Dim RowHeaderRange As Range, NameHeader As Range, InstantHeader As Range, MessageHeader As Range, StackHeader As Range
    Dim ModuleNameHeader As Range, RequestKeyHeader As Range, EspaceIdHeader As Range, ActionNameHeader As Range
    Dim EndpointHeader As Range, ActionHeader As Range, DurationHeader As Range, ScreenHeader As Range
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
    For Each Cell In RowHeaderRange.Cells
        Cell.Value = Replace(Cell.Value, "_", " ")
    Next Cell
    
    'On error, it resumes next in case it does not find one or more of the headers below'
    On Error Resume Next
    'Finds Headers by exact match'
    'Common Headers
    Set InstantHeader = RowHeaderRange.Cells.Find("Instant", Lookat:=xlWhole)
    InstantHeader.EntireColumn.ColumnWidth = 20
    
    Set RequestKeyHeader = RowHeaderRange.Cells.Find("Request Key", Lookat:=xlWhole)
    RequestKeyHeader.EntireColumn.ColumnWidth = 35
    
    Set NameHeader = RowHeaderRange.Cells.Find("Name", Lookat:=xlWhole)
    NameHeader.EntireColumn.ColumnWidth = 20
    
    'Other Headers (General and Error mostly)
    Set ActionNameHeader = RowHeaderRange.Cells.Find("Action Name", Lookat:=xlWhole)
    ActionNameHeader.EntireColumn.ColumnWidth = 18

    Set MessageHeader = RowHeaderRange.Cells.Find("Message", Lookat:=xlWhole)
    MessageHeader.EntireColumn.ColumnWidth = 80
    
    Set StackHeader = RowHeaderRange.Cells.Find("Stack", Lookat:=xlWhole)
    StackHeader.EntireColumn.ColumnWidth = 40
    
    Set ModuleNameHeader = RowHeaderRange.Cells.Find("Module Name", Lookat:=xlWhole)
    ModuleNameHeader.EntireColumn.ColumnWidth = 20
    
    'Integration Headers
    Set EndpointHeader = RowHeaderRange.Cells.Find("Endpoint", Lookat:=xlWhole)
    EndpointHeader.EntireColumn.ColumnWidth = 90
    
    Set ActionHeader = RowHeaderRange.Cells.Find("Action", Lookat:=xlWhole)
    ActionHeader.EntireColumn.ColumnWidth = 90
    
    Set DurationHeader = RowHeaderRange.Cells.Find("Duration", Lookat:=xlWhole)
    DurationHeader.EntireColumn.ColumnWidth = 10
    
    'Screen and Mobile Headers
    Set ScreenHeader = RowHeaderRange.Cells.Find("Screen", Lookat:=xlWhole)
    ScreenHeader.EntireColumn.ColumnWidth = 30
    
    'This find is just to reset the "Exact match" when CTRL+F
    Random = RowHeaderRange.Cells.Find("", Lookat:=xlPart)
    
    'Error Handler 2 to avoid resuming next'
    On Error GoTo ErrorHandler:
    
    'Applies filter to Headers'
    Set RowHeaderRange = myWorksheet.Cells.Rows(1)
    If Not myWorksheet.AutoFilterMode Then
        RowHeaderRange.AutoFilter
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



