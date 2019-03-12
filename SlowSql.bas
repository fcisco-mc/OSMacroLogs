Attribute VB_Name = "SlowSql"
Sub SlowSql()

Dim wb As Workbook
Dim myWorksheet As Worksheet, DataProcessSheet As Worksheet
Dim messageHeader As Range, RowHeaderRange As Range, moduleNameHeader As Range, queryHeader As Range, execTimeHeader As Range, messageColumn As Range, DataWrite As Range
Dim instantHeader As Range, DataWriteInstant As Range, instantData As Range, nameHeader As Range, nameHeadProc As Range, modNameProc As Range
Dim i As Integer, j As Integer, k As Integer
Dim execTimeValue As Double
Dim queryName As String, eSpaceName As String, moduleName As String

Application.DisplayAlerts = False

Set wb = ActiveWorkbook

Set myWorksheet = ActiveWorkbook.ActiveSheet
Set RowHeaderRange = myWorksheet.Cells.Range(Range("A1"), Range("A1").End(xlToRight))

'Finding headers and their locations
Set messageHeader = RowHeaderRange.Cells.Find("Message", Lookat:=xlWhole)
Set moduleNameHeader = RowHeaderRange.Cells.Find("Module Name", Lookat:=xlWhole)
Set instantHeader = RowHeaderRange.Cells.Find("Instant", Lookat:=xlWhole)
Set nameHeader = RowHeaderRange.Cells.Find("Name", Lookat:=xlWhole)

'RowHeaderRange.AutoFilter Field:=moduleNameHeader.Column, Criteria1:="=*" & "SLOWSQL" & "*"
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
    If Not cell.Value = "Message" Then
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

Call CreatePVTable

Exit Sub

General_ErrorHandler:
MsgBox Err.Description, vbCritical


End Sub

Private Sub CreatePVTable()

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
