Attribute VB_Name = "Integration"
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
'# v1.2: Adapted to Platform 11 headers

Application.DisplayAlerts = False

'This call is to guarantee header uniformization
Call AllMacros.FormatLogs

Set wb = ActiveWorkbook

Set myWorksheet = ActiveWorkbook.ActiveSheet
Set RowHeaderRange = myWorksheet.Cells.Range(Range("A1"), Range("A1").End(xlToRight))

Set durationHeader = RowHeaderRange.Cells.Find("Duration")
Set sourceHeader = RowHeaderRange.Cells.Find("Source", LookAt:=xlWhole, MatchCase:=False)
Set endPointHeader = RowHeaderRange.Cells.Find("Endpoint", LookAt:=xlWhole, MatchCase:=False)
Set actionHeader = RowHeaderRange.Cells.Find("Action", LookAt:=xlWhole, MatchCase:=False)
Set typeHeader = RowHeaderRange.Cells.Find("Type", LookAt:=xlWhole, MatchCase:=False)
Set instantHeader = RowHeaderRange.Cells.Find("Instant", LookAt:=xlWhole, MatchCase:=False)

'Validation added for Platform 11
Set nameHeader = RowHeaderRange.Cells.Find("Name", LookAt:=xlWhole, MatchCase:=False)
If nameHeader Is Nothing Then
    Set nameHeader = RowHeaderRange.Cells.Find("Espace Name", LookAt:=xlWhole, MatchCase:=False)
End If


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
    If Not LCase(cell.Value) = "instant" Then
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

