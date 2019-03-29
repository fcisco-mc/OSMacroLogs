Attribute VB_Name = "AllMacros"
Sub CallUserForm()
    MacroForm.Show
End Sub


Public Sub DeviceUUID()

'v2: fixed an infinite loop if the logs did not contain any log message with Device UUID
Dim wb As Workbook
Dim myWorksheet As Worksheet, DataProcessSheet As Worksheet
Dim RowHeaderRange As Range, EnvInfHeader As Range, EnvInfColumn As Range, HeaderStart As Range, OSColumn As Range, RawDataRegion As Range
Dim DataWrite As Range
Dim DeviceInfo_Headers(3) As String
Dim i As Integer, m As Integer, j As Integer

'Gets rid of annoying message whenever you close the Excel file'
Application.ScreenUpdating = False
Application.DisplayAlerts = False

Set wb = ActiveWorkbook
Set myWorksheet = ActiveWorkbook.ActiveSheet
Set RowHeaderRange = myWorksheet.Cells.Range(Range("A1"), Range("A1").End(xlToRight))

'Some logs have '_' characaters on the headers. The For Each loop replaces them with spaces
    For Each cell In RowHeaderRange.Cells
        cell.Value = Replace(cell.Value, "_", " ")
    Next cell

'Finds "Environment Information Header" and defines its column
On Error GoTo EnvInf_ErrorHandler:
Set EnvInfHeader = RowHeaderRange.Cells.Find("Environment Information", Lookat:=xlWhole)

RowHeaderRange.AutoFilter Field:=EnvInfHeader.Column, Criteria1:="=*" & "DeviceUUID" & "*"

On Error Resume Next
Worksheets("DeviceInformation").Delete
'Adds a new workshet
Set DataProcessSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count), Type:=xlWorksheet)
DataProcessSheet.Name = "DeviceInformation"

On Error GoTo General_ErrorHandler:

'Defining the range of the environment information column
myWorksheet.Activate

RowHeaderRange.Cells.Find("Environment Information").Activate
Set EnvInfColumn = ActiveCell.EntireColumn.SpecialCells(xlCellTypeVisible)

'Headers in the new sheet
DeviceInfo_Headers(0) = "DeviceUUID"
DeviceInfo_Headers(1) = "Cordova"
DeviceInfo_Headers(2) = "OperatingSystem"
DeviceInfo_Headers(3) = "DeviceModel"

Set HeaderStart = DataProcessSheet.Range("A1").Cells
For i = LBound(DeviceInfo_Headers) To UBound(DeviceInfo_Headers)
    HeaderStart.Offset(0, i).Value = DeviceInfo_Headers(i)
Next i

'Setting the range of the envinfo on the new sheet
'Set EnvInfColumn = DataProcessSheet.Cells.Range(Range("A2"), Range("A2").End(xlDown))
Set DataWrite = DataProcessSheet.Range("A2")
For Each cell In EnvInfColumn.Cells.SpecialCells(xlCellTypeVisible)
    
    ' Replace line breaks with ; and "," with nothing: Some logs are formatted differently for some reason
        ' This guarantee the same format for all
    If Not cell.Value = "Environment Information" Then
        cell.Value = Replace(cell.Text, vbLf, ";")
        'Debug.Print Cell.Value
        cell.Value = Replace(cell.Text, ",", "")
        'Debug.Print Cell.Value
        For k = LBound(DeviceInfo_Headers) To UBound(DeviceInfo_Headers)
            'Parsing Device Information from string
            i = InStr(1, cell.Value, DeviceInfo_Headers(k))
            If i = 0 And k = 0 Then Exit For
            'Debug.Print i
            ' If i = 0 it is because it didn't found the DeviceInfo, hence, it is undefined
            If i = 0 Then
                DataWrite.Offset(0, k).Value = "Undefined"
                'Debug.Print DataWrite.Offset(0, k).Value
            Else
                m = InStr(i, cell.Value, ":")
                j = InStr(i, cell.Value, ";")
                DataWrite.Offset(0, k).Value = Mid(cell.Text, m + 1, j - m - 1)
                DataWrite.Offset(0, k).Value = Trim(DataWrite.Offset(0, k).Value)
                'Debug.Print DataWrite.Offset(0, k).Value
            End If
            'Debug.Print i, m, j
    
        Next k
        If i = 0 And k = 0 Then Exit For
        Set DataWrite = DataWrite.Offset(1, 0)
    End If
Next cell
    
'Parsing OS version for better visualization in Charts
DataProcessSheet.Activate
Set OSColumn = DataProcessSheet.Cells.Range(Range("C1"), Range("C1").End(xlDown))
DataProcessSheet.Range("E1").Value = "OperatingSystem_Version"

'v2_Exit sub if there's no logs with DeviceUUID
OSColumn.Cells(2, 1).Select
If OSColumn.Cells(2, 1).Value = "" Then
    DataProcessSheet.Delete
    MsgBox "Looks like these Error logs don't have any log with DeviceUUID in the Environment Information column. Make sure to select Error logs with Mobile errors.", vbInformation
    Exit Sub
End If


For Each cell In OSColumn

    If Not cell.Value = "Undefined" And Not cell.Text = DeviceInfo_Headers(2) Then
        i = InStr(1, cell.Value, " ")
        j = Len(cell.Value)
        'Debug.Print i, j
        cell.Offset(0, 2).Value = Mid(cell.Text, i + 1, j)
        cell.Value = Mid(cell.Text, 1, i)
        cell.Offset(0, 2).Value = Trim(cell.Offset(0, 2).Value)
    ElseIf cell.Text = DeviceInfo_Headers(2) Then
        'Skip to next cell
    Else
        cell.Offset(0, 2).Value = "Undefined"
    End If

Next cell

DataProcessSheet.Columns.AutoFit

Call CreatePVTable_Mobile


Exit Sub

EnvInf_ErrorHandler:
    MsgBox ("Couldn't find the header 'Environment Information' on the Logs"), vbCritical


General_ErrorHandler:
MsgBox Err.Description, vbCritical
End Sub



Private Sub CreatePVTable_Mobile()

Dim wb As Workbook
Dim ws As Worksheet, DataSheet As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim LastRow As Long, LasCol As Long
Dim PRange As Range
Dim ch As Chart

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Set wb = ActiveWorkbook
Set DataSheet = wb.Worksheets("DeviceInformation")

'Delete a sheet call "PivotTable" to avoid conflict when creating it
On Error Resume Next
Worksheets("PivotTable").Delete

On Error GoTo General_ErrorHandler:

'Creating Pivot Cache
Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        sourceData:=DataSheet.Name & "!" & DataSheet.Range("A1").CurrentRegion.Address, _
        Version:=xlPivotTableVersion15)

    
'Creating Pivot Table
    'TableDestination must be empty because the macro is on the Personal.xlsb and is not dynamic.
        ' More info: https://support.microsoft.com/en-us/help/940166/error-message-when-you-play-a-recorded-macro-to-create-a-pivottable-in
Set pt = pc.CreatePivotTable( _
    TableDestination:="", _
    TableName:="MyPivotTable")

Set ws = wb.ActiveSheet
    ws.Name = "PivotTable"

Set pf = pt.PivotFields("OperatingSystem")
        pf.Orientation = xlRowField
        pf.Position = 1

Set pf = pt.PivotFields("DeviceUUID")
        pf.Orientation = xlRowField
        pf.Position = 2
        
Set pf = pt.PivotFields("OperatingSystem_Version")
        pf.Orientation = xlRowField
        pf.Position = 3


pt.AddDataField pt.PivotFields("DeviceUUID"), , xlCount

'Checking if the Chart already exists
    Application.DisplayAlerts = False
    For Each ch In wb.Charts
        If ch.Name = "OS Analysis" Then ch.Delete
    Next ch

'Adding a Chart
Set ch = Charts.Add
    ch.Name = "OS Analysis"
    ch.SetSourceData pt.TableRange1
    ch.ChartType = xlColumnStacked




Exit Sub

General_ErrorHandler:
MsgBox Err.Description, vbCritical


End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


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
    Set instantHeader = RowHeaderRange.Cells.Find("Instant", Lookat:=xlWhole)
    instantHeader.EntireColumn.ColumnWidth = 20
    
    Set RequestKeyHeader = RowHeaderRange.Cells.Find("Request Key", Lookat:=xlWhole)
    RequestKeyHeader.EntireColumn.ColumnWidth = 35
    
    Set nameHeader = RowHeaderRange.Cells.Find("Name", Lookat:=xlWhole)
    nameHeader.EntireColumn.ColumnWidth = 20
    
    'Other Headers (General and Error mostly)
    Set ActionNameHeader = RowHeaderRange.Cells.Find("Action Name", Lookat:=xlWhole)
    ActionNameHeader.EntireColumn.ColumnWidth = 18

    Set messageHeader = RowHeaderRange.Cells.Find("Message", Lookat:=xlWhole)
    messageHeader.EntireColumn.ColumnWidth = 80
    
    Set StackHeader = RowHeaderRange.Cells.Find("Stack", Lookat:=xlWhole)
    StackHeader.EntireColumn.ColumnWidth = 40
    
    Set moduleNameHeader = RowHeaderRange.Cells.Find("Module Name", Lookat:=xlWhole)
    moduleNameHeader.EntireColumn.ColumnWidth = 20
    
    'Integration Headers
    Set endPointHeader = RowHeaderRange.Cells.Find("Endpoint", Lookat:=xlWhole)
    endPointHeader.EntireColumn.ColumnWidth = 90
    
    Set actionHeader = RowHeaderRange.Cells.Find("Action", Lookat:=xlWhole)
    actionHeader.EntireColumn.ColumnWidth = 90
    
    Set durationHeader = RowHeaderRange.Cells.Find("Duration", Lookat:=xlWhole)
    durationHeader.EntireColumn.ColumnWidth = 10
    
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


'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Sub Integration()

Dim wb As Workbook
Dim myWorksheet As Worksheet, DataProcessSheet As Worksheet
Dim RowHeaderRange As Range, durationHeader As Range, endPointHeader As Range, DataWrite As Range, actionHeader As Range, typeHeader As Range, sourceHeader As Range
Dim instantHeader As Range, DataWriteInstant As Range, instantData As Range, durationData As Range, sourceData As Range, endpointData As Range, actionData As Range, typeData As Range, nameHeader As Range
Dim nameData As Range
Dim i As Integer, j As Integer
Dim execTimeValue As Double
Dim queryName As String

'# v1.1: Added Source, eSpaceName, Endpoint and Type to Filter Fields

Application.DisplayAlerts = False

Set wb = ActiveWorkbook

Set myWorksheet = ActiveWorkbook.ActiveSheet
Set RowHeaderRange = myWorksheet.Cells.Range(Range("A1"), Range("A1").End(xlToRight))

Set durationHeader = RowHeaderRange.Cells.Find("Duration", Lookat:=xlWhole)
Set sourceHeader = RowHeaderRange.Cells.Find("Source", Lookat:=xlWhole)
Set endPointHeader = RowHeaderRange.Cells.Find("Endpoint", Lookat:=xlWhole)
Set actionHeader = RowHeaderRange.Cells.Find("Action", Lookat:=xlWhole)
Set typeHeader = RowHeaderRange.Cells.Find("Type", Lookat:=xlWhole)
Set instantHeader = RowHeaderRange.Cells.Find("Instant", Lookat:=xlWhole)
Set nameHeader = RowHeaderRange.Cells.Find("Name", Lookat:=xlWhole)

On Error Resume Next
Worksheets("IntegrationData").Delete
'Adds a new workshet
Set DataProcessSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count), Type:=xlWorksheet)
DataProcessSheet.Name = "IntegrationData"

On Error GoTo General_ErrorHandler:

Set instantData = DataProcessSheet.Range("A1")
instantData.Value = "Instant"

Set durationData = DataProcessSheet.Range("B1")
durationData.Value = "Duration"

Set sourceData = DataProcessSheet.Range("C1")
sourceData.Value = "Source"

Set endpointData = DataProcessSheet.Range("D1")
endpointData.Value = "Endpoint"

Set actionData = DataProcessSheet.Range("E1")
actionData.Value = "Action"

Set typeData = DataProcessSheet.Range("F1")
typeData.Value = "Type"

Set nameData = DataProcessSheet.Range("G1")
nameData.Value = "eSpace Name"


myWorksheet.Activate

Set InstantColumn = myWorksheet.Cells.Range(instantHeader, instantHeader.End(xlDown))

Set DataWrite = DataProcessSheet.Range("A2")

For Each cell In InstantColumn.Cells.SpecialCells(xlCellTypeVisible)
    If Not cell.Value = "Instant" Then
        'Writting Instant
        DataWrite.Value = cell.Value
        
        'Duration
        DataWrite.Offset(0, durationData.Column - instantData.Column).Value = cell.Offset(0, (durationHeader.Column - instantHeader.Column)).Value
        
        'Source
        DataWrite.Offset(0, sourceData.Column - instantData.Column).Value = cell.Offset(0, (sourceHeader.Column - instantHeader.Column)).Value
        
        'Endpoint
        DataWrite.Offset(0, endpointData.Column - instantData.Column).Value = cell.Offset(0, (endPointHeader.Column - instantHeader.Column)).Value
        
        'Action
        DataWrite.Offset(0, actionData.Column - instantData.Column).Value = cell.Offset(0, (actionHeader.Column - instantHeader.Column)).Value
        
        'Type
        DataWrite.Offset(0, typeData.Column - instantData.Column).Value = cell.Offset(0, (typeHeader.Column - instantHeader.Column)).Value
        
        'eSpace Name
        DataWrite.Offset(0, nameData.Column - instantData.Column).Value = cell.Offset(0, (nameHeader.Column - instantHeader.Column)).Value
        
        Set DataWrite = DataWrite.Offset(1, 0)
        
    End If
Next cell

DataProcessSheet.Columns.AutoFit
DataProcessSheet.Visible = xlSheetHidden

Call CreatePVTable_Integration

Exit Sub

General_ErrorHandler:
MsgBox Err.Description, vbCritical

End Sub


Private Sub CreatePVTable_Integration()

Dim wb As Workbook
Dim ws As Worksheet, DataSheet As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim LastRow As Long, LasCol As Long
Dim PRange As Range
Dim ch As Chart

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Set wb = ActiveWorkbook


Set DataSheet = wb.Worksheets("IntegrationData")

'Delete a sheet call "PivotTable" to avoid conflict when creating it
On Error Resume Next
Worksheets("PivotTable").Delete

On Error GoTo General_ErrorHandler:

'Creating Pivot Cache
Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        sourceData:=ThisWorkbook.ActiveSheet.Name & "!" & DataSheet.Range("A1").CurrentRegion.Address, _
        Version:=xlPivotTableVersion15)

'Creating Pivot Table
    'TableDestination must be empty because the macro is on the Personal.xlsb and is not dynamic.
        ' More info: https://support.microsoft.com/en-us/help/940166/error-message-when-you-play-a-recorded-macro-to-create-a-pivottable-in
Set pt = pc.CreatePivotTable( _
    TableDestination:="", _
    TableName:="MyPivotTable")

Set ws = wb.ActiveSheet
    ws.Name = "PivotTable"
    
'Row Fields
Set pf = pt.PivotFields("Instant")
        pf.Orientation = xlRowField
        pf.Position = 1

'Column Fields
Set pf = pt.PivotFields("Action")
        pf.Orientation = xlColumnField
        pf.Position = 1
        
'Filter Fields
Set pf = pt.PivotFields("Type")
        pf.Orientation = xlPageField
        pf.Position = 1

Set pf = pt.PivotFields("eSpace Name")
        pf.Orientation = xlPageField
        pf.Position = 2

Set pf = pt.PivotFields("Source")
        pf.Orientation = xlPageField
        pf.Position = 3

Set pf = pt.PivotFields("Endpoint")
        pf.Orientation = xlPageField
        pf.Position = 4
        
'pt.AddDataField pt.PivotFields("Execution Time"), , xlSum
pt.AddDataField pt.PivotFields("Duration"), , xlAverage
        
'Checking if the Chart already exists
    Application.DisplayAlerts = False
    For Each ch In wb.Charts
        If ch.Name = "Integration Analysis Chart" Then ch.Delete
    Next ch

'Adding a Chart
Set ch = Charts.Add
    ch.Name = "Integration Analysis Chart"
    ch.SetSourceData pt.TableRange1
    ch.ChartType = xlLineMarkers


Exit Sub

General_ErrorHandler:
MsgBox Err.Description, vbCritical

End Sub


'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Sub SlowSql()

Dim wb As Workbook
Dim myWorksheet As Worksheet, DataProcessSheet As Worksheet
Dim messageHeader As Range, RowHeaderRange As Range, moduleNameHeader As Range, queryHeader As Range, execTimeHeader As Range, messageColumn As Range, DataWrite As Range
Dim instantHeader As Range, DataWriteInstant As Range, instantData As Range, nameHeader As Range, nameHeadProc As Range, modNameProc As Range
Dim i As Integer, j As Integer, k As Integer
Dim execTimeValue As Double
Dim queryName As String, eSpaceName As String, moduleName As String

'# v1.1: Added Slow Extensions and eSpace Names
'# v1.2: Validated if the Maximum Number of log entries was exceeded

Application.DisplayAlerts = False

Set wb = ActiveWorkbook

Set myWorksheet = ActiveWorkbook.ActiveSheet
Set RowHeaderRange = myWorksheet.Cells.Range(Range("A1"), Range("A1").End(xlToRight))

'Finding headers and their locations
Set messageHeader = RowHeaderRange.Cells.Find("Message", Lookat:=xlWhole)
Set moduleNameHeader = RowHeaderRange.Cells.Find("Module Name", Lookat:=xlWhole)
Set instantHeader = RowHeaderRange.Cells.Find("Instant", Lookat:=xlWhole)
Set nameHeader = RowHeaderRange.Cells.Find("Name", Lookat:=xlWhole)
    
' Autofilter for SlowSql and SlowExtension
RowHeaderRange.AutoFilter Field:=moduleNameHeader.Column, Criteria1:=Array("SLOWSQL", "SLOWEXTENSION"), Operator:=xlFilterValues


On Error Resume Next
Worksheets("SlowSQL").Delete
'Adds a new workshet
Set DataProcessSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count), Type:=xlWorksheet)
DataProcessSheet.Name = "SlowSQL"

On Error GoTo General_ErrorHandler:

'Setting the headers in the data source of the pivot table
Set queryHeader = DataProcessSheet.Range("B1")
queryHeader.Value = "Query"

Set execTimeHeader = DataProcessSheet.Range("C1")
execTimeHeader.Value = "Execution Time"

Set instantData = DataProcessSheet.Range("A1")
instantData.Value = "Instant"

Set nameHeadProc = DataProcessSheet.Range("D1")
nameHeadProc = "eSpace Name"

Set modNameProc = DataProcessSheet.Range("E1")
modNameProc = "Module Name"

myWorksheet.Activate
messageHeader.Activate
'Set messageColumn = ActiveCell.EntireColumn.SpecialCells(xlCellTypeVisible)
Set messageColumn = myWorksheet.Cells.Range(messageHeader, messageHeader.End(xlDown))

Set DataWrite = DataProcessSheet.Range("B2")
Set DataWriteInstant = DataProcessSheet.Range("A2")
For Each cell In messageColumn.Cells.SpecialCells(xlCellTypeVisible)
    ' Added a validation to see if the maximum number of log entries was exceeded. This message is also marked as slow and does not have the expected format
    If Not (cell.Value = "Message" Or InStr(1, cell.Value, "The maximum number") > 0) Then
        'Debug.Print Cell.Offset(0, -(messageColumn.Column - instantHeader.Column)).Value
        'Instant (Message column is the reference location)
        DataWriteInstant.Value = cell.Offset(0, -(messageColumn.Column - instantHeader.Column)).Value
        
        ' Query string
        i = InStr(1, cell.Value, "took")
        queryName = Mid(cell.Value, 1, i - 1)
        DataWrite.Value = queryName
        
        'Execution time
        i = i + 4
        j = InStr(i, cell.Value, "ms")
        execTimeValue = Mid(cell.Value, i, j - i)
        DataWrite.Offset(0, 1).Value = execTimeValue
        
        'eSpace Name
        eSpaceName = cell.Offset(0, nameHeader.Column - messageColumn.Column).Value
        DataWrite.Offset(0, 2).Value = eSpaceName
        
        'ModuleName
        moduleName = cell.Offset(0, moduleNameHeader.Column - messageColumn.Column).Value
        DataWrite.Offset(0, 3).Value = moduleName
        
        Set DataWrite = DataWrite.Offset(1, 0)
        Set DataWriteInstant = DataWriteInstant.Offset(1, 0)
        
    End If

Next cell

DataProcessSheet.Columns.AutoFit
DataProcessSheet.Visible = xlSheetHidden

Call CreatePVTable_SlowSql

Exit Sub

General_ErrorHandler:
MsgBox Err.Description, vbCritical


End Sub

Private Sub CreatePVTable_SlowSql()

Dim wb As Workbook
Dim ws As Worksheet, DataSheet As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim LastRow As Long, LasCol As Long
Dim PRange As Range
Dim ch As Chart

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Set wb = ActiveWorkbook


Set DataSheet = wb.Worksheets("SlowSQL")

'Delete a sheet call "PivotTable" to avoid conflict when creating it
On Error Resume Next
Worksheets("PivotTable").Delete

On Error GoTo General_ErrorHandler:

'Creating Pivot Cache
Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        sourceData:=DataSheet.Name & "!" & DataSheet.Range("A1").CurrentRegion.Address, _
        Version:=xlPivotTableVersion15)

'Creating Pivot Table
    'TableDestination must be empty because the macro is on the Personal.xlsb and is not dynamic.
        ' More info: https://support.microsoft.com/en-us/help/940166/error-message-when-you-play-a-recorded-macro-to-create-a-pivottable-in
Set pt = pc.CreatePivotTable( _
    TableDestination:="", _
    TableName:="MyPivotTable")

Set ws = wb.ActiveSheet
    ws.Name = "PivotTable"

'Row Field
Set pf = pt.PivotFields("Instant")
        pf.Orientation = xlRowField
        pf.Position = 1

'Column Field
Set pf = pt.PivotFields("Query")
        pf.Orientation = xlColumnField
        pf.Position = 1

'Filter Fields
Set pf = pt.PivotFields("eSpace Name")
        pf.Orientation = xlPageField
        pf.Position = 1
        
Set pf = pt.PivotFields("Module Name")
        pf.Orientation = xlPageField
        pf.Position = 2



        
'pt.AddDataField pt.PivotFields("Execution Time"), , xlSum
pt.AddDataField pt.PivotFields("Execution Time"), , xlAverage
        
'Checking if the Chart already exists
    Application.DisplayAlerts = False
    For Each ch In wb.Charts
        If ch.Name = "SlowSQL Analysis" Then ch.Delete
    Next ch

'Adding a Chart
Set ch = Charts.Add
    ch.Name = "SlowSQL Analysis"
    ch.SetSourceData pt.TableRange1
    ch.ChartType = xlLineMarkers


Exit Sub

General_ErrorHandler:
MsgBox Err.Description, vbCritical

End Sub

