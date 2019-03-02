Attribute VB_Name = "SlowSql"
Sub SlowSql()

Dim wb As Workbook
Dim myWorksheet As Worksheet, DataProcessSheet As Worksheet
Dim messageHeader As Range, RowHeaderRange As Range, moduleNameHeader As Range, queryHeader As Range, execTimeHeader As Range, messageColumn As Range, DataWrite As Range
Dim instantHeader As Range, DataWriteInstant As Range, instantData As Range
Dim i As Integer, j As Integer
Dim execTimeValue As Double
Dim queryName As String

Application.DisplayAlerts = False

Set wb = ActiveWorkbook

Set myWorksheet = ActiveWorkbook.ActiveSheet
Set RowHeaderRange = myWorksheet.Cells.Range(Range("A1"), Range("A1").End(xlToRight))

Set messageHeader = RowHeaderRange.Cells.Find("Message", Lookat:=xlWhole)
Set moduleNameHeader = RowHeaderRange.Cells.Find("Module Name", Lookat:=xlWhole)
Set instantHeader = RowHeaderRange.Cells.Find("Instant", Lookat:=xlWhole)

RowHeaderRange.AutoFilter Field:=moduleNameHeader.Column, Criteria1:="=*" & "SLOWSQL" & "*"

On Error Resume Next
Worksheets("SlowSQL").Delete
'Adds a new workshet
Set DataProcessSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count), Type:=xlWorksheet)
DataProcessSheet.Name = "SlowSQL"

On Error GoTo General_ErrorHandler:

Set queryHeader = DataProcessSheet.Range("B1")
queryHeader.Value = "Query"

Set execTimeHeader = DataProcessSheet.Range("C1")
execTimeHeader.Value = "Execution Time"

Set instantData = DataProcessSheet.Range("A1")
instantData.Value = "Instant"

'Defining the range of the environment information column
myWorksheet.Activate
messageHeader.Activate
'Set messageColumn = ActiveCell.EntireColumn.SpecialCells(xlCellTypeVisible)
Set messageColumn = myWorksheet.Cells.Range(messageHeader, messageHeader.End(xlDown))

Set DataWrite = DataProcessSheet.Range("B2")
Set DataWriteInstant = DataProcessSheet.Range("A2")
For Each Cell In messageColumn.Cells.SpecialCells(xlCellTypeVisible)
    If Not Cell.Value = "Message" Then
        If Cell.Value = "" Then Exit For
        'Debug.Print Cell.Offset(0, -(messageColumn.Column - instantHeader.Column)).Value
        DataWriteInstant.Value = Cell.Offset(0, -(messageColumn.Column - instantHeader.Column)).Value
        
        i = InStr(1, Cell.Value, "took")
        queryName = Mid(Cell.Value, 1, i - 1)
        DataWrite.Value = queryName
        
        i = i + 4
        j = InStr(i, Cell.Value, "ms")
        execTimeValue = Mid(Cell.Value, i, j - i)
        DataWrite.Offset(0, 1).Value = execTimeValue
        
        Set DataWrite = DataWrite.Offset(1, 0)
        Set DataWriteInstant = DataWriteInstant.Offset(1, 0)
        
    End If

Next Cell

DataProcessSheet.Columns.AutoFit

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
        SourceData:=DataSheet.Name & "!" & DataSheet.Range("A1").CurrentRegion.Address, _
        Version:=xlPivotTableVersion15)

'Creating Pivot Table
    'TableDestination must be empty because the macro is on the Personal.xlsb and is not dynamic.
        ' More info: https://support.microsoft.com/en-us/help/940166/error-message-when-you-play-a-recorded-macro-to-create-a-pivottable-in
Set pt = pc.CreatePivotTable( _
    TableDestination:="", _
    TableName:="MyPivotTable")

Set ws = wb.ActiveSheet
    ws.Name = "PivotTable"
    
Set pf = pt.PivotFields("Instant")
        pf.Orientation = xlRowField
        pf.Position = 1

Set pf = pt.PivotFields("Query")
        pf.Orientation = xlColumnField
        pf.Position = 1
        
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
