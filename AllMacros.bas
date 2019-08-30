Attribute VB_Name = "AllMacros"
Option Explicit
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'                                                                           *** AllMacros ***
' #v1.2:
        ' Device UUID: # v2.1
        ' Format Logs: # v3.1
        ' Integration: # v1.2
        ' SlowSQL: # v1.4
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Sub CallUserForm()
    MacroForm.Show
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'                                                                       *** Timers Performance Macro ***
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Sub Timers()

Dim wb As Workbook
Dim myWorksheet As Worksheet, DataProcessSheet As Worksheet
Dim RowHeaderRange As Range, nameHeader As Range, timerName As Range

Application.DisplayAlerts = False

'This call is to guarantee header uniformization
Call FormatLogs

Set wb = ActiveWorkbook

Set myWorksheet = ActiveWorkbook.ActiveSheet
Set RowHeaderRange = myWorksheet.Cells.Range(Range("A1"), Range("A1").End(xlToRight))



'Validation added for Platform 11
Set timerName = RowHeaderRange.Cells.Find("name(", LookAt:=xlPart, MatchCase:=False)
If Not timerName Is Nothing Then
    timerName.Value = "cyclicjobname"
ElseIf timerName Is Nothing Then
    Set timerName = RowHeaderRange.Cells.Find("cyclicjobname", LookAt:=xlWhole, MatchCase:=False)
End If

Set nameHeader = RowHeaderRange.Cells.Find("name", LookAt:=xlWhole, MatchCase:=False)
If Not nameHeader Is Nothing Then
    nameHeader.Value = "espacename"
ElseIf nameHeader Is Nothing Then
    Set nameHeader = RowHeaderRange.Cells.Find("espacename", LookAt:=xlWhole, MatchCase:=False)
End If

Set DataProcessSheet = myWorksheet

Call CreatePVTable("Timers", DataProcessSheet, 2)

End Sub


'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'                                                                       *** Screens Performance Macro ***
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Sub Screens()

Dim wb As Workbook
Dim myWorksheet As Worksheet, DataProcessSheet As Worksheet
Dim RowHeaderRange As Range, nameHeader As Range

Application.DisplayAlerts = False

'This call is to guarantee header uniformization
Call FormatLogs

Set wb = ActiveWorkbook

Set myWorksheet = ActiveWorkbook.ActiveSheet
Set RowHeaderRange = myWorksheet.Cells.Range(Range("A1"), Range("A1").End(xlToRight))

'Validation added for Platform 11
Set nameHeader = RowHeaderRange.Cells.Find("name", LookAt:=xlWhole, MatchCase:=False)
If Not nameHeader Is Nothing Then
    nameHeader.Value = "espacename"
ElseIf nameHeader Is Nothing Then
    Set nameHeader = RowHeaderRange.Cells.Find("espacename", LookAt:=xlWhole, MatchCase:=False)
End If

Set DataProcessSheet = myWorksheet

Call CreatePVTable("Screens", DataProcessSheet, 2)


End Sub


'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'                                                                       *** Mobile Macro ***
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Sub DeviceUUID(iOS_cb As Boolean, Android_cb As Boolean)

' #v2: fixed an infinite loop if the logs did not contain any log message with Device UUID
' #v2.1: Added option to filter the logs for iOS, Android or both

Dim wb As Workbook
Dim myWorksheet As Worksheet, DataProcessSheet As Worksheet
Dim RowHeaderRange As Range, EnvInfHeader As Range, EnvInfColumn As Range, HeaderStart As Range, OSColumn As Range, RawDataRegion As Range
Dim DataWrite As Range
Dim DeviceInfo_Headers(3) As String
Dim i As Integer, m As Integer, j As Integer

'Gets rid of annoying message whenever you close the Excel file'
Application.ScreenUpdating = False
Application.DisplayAlerts = False

'This call is to guarantee header uniformization
Call AllMacros.FormatLogs

Set wb = ActiveWorkbook
Set myWorksheet = ActiveWorkbook.ActiveSheet
Set RowHeaderRange = myWorksheet.Cells.Range(Range("A1"), Range("A1").End(xlToRight))

'Some logs have '_' characaters on the headers. The For Each loop replaces them with spaces
    For Each cell In RowHeaderRange.Cells
        cell.Value = LCase(Replace(cell.Value, "_", ""))
    Next cell

'Finds "Environment Information Header" and defines its column
On Error GoTo EnvInf_ErrorHandler:
Set EnvInfHeader = RowHeaderRange.Cells.Find("environmentinformation", LookAt:=xlPart, MatchCase:=False)

' Filter for Mobile errors looking for DeviceUUID and/or iOS/Android
If iOS_cb = True And Android_cb = True Then
    RowHeaderRange.AutoFilter Field:=EnvInfHeader.Column, Criteria1:="=*" & "DeviceUUID" & "*"
ElseIf iOS_cb = False And Android_cb = True Then
    RowHeaderRange.AutoFilter Field:=EnvInfHeader.Column, Criteria1:=Array("=*" & "DeviceUUID" & "*", "=*" & "Android" & "*"), Operator:=xlAnd
ElseIf iOS_cb = True And Android_cb = False Then
    RowHeaderRange.AutoFilter Field:=EnvInfHeader.Column, Criteria1:=Array("=*" & "DeviceUUID" & "*", "=*" & "iOS" & "*"), Operator:=xlAnd
End If


On Error Resume Next
Worksheets("DeviceInformation").Delete
'Adds a new workshet
Set DataProcessSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count), Type:=xlWorksheet)
DataProcessSheet.Name = "DeviceInformation"

On Error GoTo General_ErrorHandler:

'Defining the range of the environment information column
myWorksheet.Activate

RowHeaderRange.Cells.Find("environmentinformation", LookAt:=xlPart, MatchCase:=False).Activate
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
    If Not LCase(cell.Value) = "environmentinformation" Then
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

Call CreatePVTable("Mobile", DataProcessSheet)

Exit Sub

EnvInf_ErrorHandler:
    MsgBox ("Couldn't find the header 'Environment Information' on the Logs"), vbCritical


General_ErrorHandler:
MsgBox Err.Description, vbCritical
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'                                                                       *** Pivot Table and Charts Creation ***
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Private Sub CreatePVTable(PivotTableType As String, Datasheet As Worksheet, Optional PVNumber As Integer)

Dim wb As Workbook
Dim ws As Worksheet, ws2 As Worksheet, SourceDataSheet As Worksheet
Dim pc As PivotCache, pc2 As PivotCache
Dim pt As PivotTable, pt2 As PivotTable
Dim pf As PivotField, pf2 As PivotField
Dim LastRow As Long, LasCol As Long
Dim PRange As Range, instantField As Range, cell As Range
Dim ch As Chart, ch2 As Chart
Dim chartName As String

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Set wb = ActiveWorkbook
Set SourceDataSheet = Datasheet
chartName = "OS Chart"

'Delete a sheet call "PivotTable" to avoid conflict when creating it
For Each ws In wb.Worksheets
    If ws.Name Like ("*" & "PivotTable" & "*") Then ws.Delete
Next ws

On Error GoTo General_ErrorHandler:

'PT NOT WORKING;
'pt = PT_Table(SourceDataSheet, 1)
'Creating Pivot Cache
Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        sourceData:=Datasheet.Name & "!" & Datasheet.Range("A1").CurrentRegion.Address, _
        Version:=xlPivotTableVersion15)
    
'Creating Pivot Table
    'TableDestination must be empty because the macro is on the Personal.xlsb and is not dynamic.
        ' More info: https://support.microsoft.com/en-us/help/940166/error-message-when-you-play-a-recorded-macro-to-create-a-pivottable-in
Set pt = pc.CreatePivotTable( _
    TableDestination:="", _
    TableName:="MyPivotTable")

Set ws = wb.ActiveSheet
    ws.Name = "PivotTable"
    
If PVNumber = 2 Then
    Set pc2 = ThisWorkbook.PivotCaches.Create( _
            SourceType:=xlDatabase, _
            sourceData:=Datasheet.Name & "!" & Datasheet.Range("A1").CurrentRegion.Address, _
            Version:=xlPivotTableVersion15)
    
    Datasheet.Activate
    Set pt2 = pc2.CreatePivotTable( _
        TableDestination:="", _
        TableName:="MyPivotTable2")
    
    Set ws2 = wb.ActiveSheet
        ws2.Name = "PivotTable2"
End If

'Checking if the Chart already exists
Application.DisplayAlerts = False
For Each ch In wb.Charts
    If ch.Name Like ("*" & chartName & "*") Then ch.Delete
Next ch

' Mobile
If PivotTableType = "Mobile" Then
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
    
    'Adding a Chart
    Set ch = Charts.Add
        ch.Name = chartName
        ch.SetSourceData pt.TableRange1
        ch.ChartType = xlColumnStacked

' SLOWSQL
ElseIf PivotTableType = "SLOW" Then
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

    pt.AddDataField pt.PivotFields("Execution Time"), , xlAverage
    ws.Activate
    'Adding a Chart
    Set ch = Charts.Add
        ch.Name = chartName + " Performance Analysis"
        ch.SetSourceData Source:=pt.TableRange1
        ch.ChartType = xlLineMarkers
       
    'Row Field
    Set pf2 = pt2.PivotFields("Instant")
            pf2.Orientation = xlRowField
            pf2.Position = 1
            
   'Grouping the Instant by day
   For Each cell In ws2.UsedRange.Columns("A").Cells
        If IsDate(cell.Value) Then
            cell.Group _
                Start:=True, End:=True, Periods:=Array(False, False, False, True, False, False, False)
                Exit For
        End If
    Next cell
        
     
    pt2.AddDataField pt2.PivotFields("Query"), , xlCount
    
    'Column Field
    Set pf2 = pt2.PivotFields("Query")
            pf2.Orientation = xlColumnField
            pf2.Position = 1
    
    'Adding a Chart
    ws2.Activate
    Set ch2 = Charts.Add
        ch2.Name = chartName + " Count"
        ch2.SetSourceData Source:=pt2.TableRange1
        ch2.ChartType = xlColumnClustered
    
' Integration
ElseIf PivotTableType = "Integration" Then
    'Row Fields
    Set pf = pt.PivotFields("instant")
            pf.Orientation = xlRowField
            pf.Position = 1
    
    'Column Fields
    Set pf = pt.PivotFields("action")
            pf.Orientation = xlColumnField
            pf.Position = 1
            
    'Filter Fields
    Set pf = pt.PivotFields("type")
            pf.Orientation = xlPageField
            pf.Position = 1
    
    Set pf = pt.PivotFields("espacename")
            pf.Orientation = xlPageField
            pf.Position = 2
    
    Set pf = pt.PivotFields("source")
            pf.Orientation = xlPageField
            pf.Position = 3
    
    Set pf = pt.PivotFields("endpoint")
            pf.Orientation = xlPageField
            pf.Position = 4

    pt.AddDataField pt.PivotFields("duration"), , xlAverage
    
    'Adding a Chart
    ws.Activate
    Set ch = Charts.Add
        ch.Name = chartName + " Performance Analysis"
        ch.SetSourceData pt.TableRange1
        ch.ChartType = xlLineMarkers
        
    'Row Field
    Set pf2 = pt2.PivotFields("instant")
            pf2.Orientation = xlRowField
            pf2.Position = 1
            
   'Grouping the Instant by day
   For Each cell In ws2.UsedRange.Columns("A").Cells
        If IsDate(cell.Value) Then
            cell.Group _
                Start:=True, End:=True, Periods:=Array(False, False, False, True, False, False, False)
                Exit For
        End If
    Next cell
        
    pt2.AddDataField pt2.PivotFields("action"), , xlCount
    
    'Column Field
    Set pf2 = pt2.PivotFields("action")
            pf2.Orientation = xlColumnField
            pf2.Position = 1
    
    'Adding a Chart
    ws2.Activate
    Set ch2 = Charts.Add
        ch2.Name = chartName + " Count"
        ch2.SetSourceData Source:=pt2.TableRange1
        ch2.ChartType = xlColumnClustered
        
ElseIf PivotTableType = "Screens" Then
    'Row Fields
    Set pf = pt.PivotFields("instant")
            pf.Orientation = xlRowField
            pf.Position = 1
    
    'Column Fields
    Set pf = pt.PivotFields("screen")
            pf.Orientation = xlColumnField
            pf.Position = 1
            
    Set pf = pt.PivotFields("espacename")
            pf.Orientation = xlColumnField
            pf.Position = 2
            
    pt.AddDataField pt.PivotFields("duration"), , xlAverage
    
    'Adding a Chart
    ws.Activate
    Set ch = Charts.Add
        ch.Name = chartName + " Performance Analysis"
        ch.SetSourceData pt.TableRange1
        ch.ChartType = xlLineMarkers
        
    
    'Row Field
    Set pf2 = pt2.PivotFields("instant")
            pf2.Orientation = xlRowField
            pf2.Position = 1
    
    'Column Fields
    Set pf2 = pt2.PivotFields("espacename")
            pf2.Orientation = xlColumnField
            pf2.Position = 1
            
    'Grouping the Instant by day
   For Each cell In ws2.UsedRange.Columns("A").Cells
        If IsDate(cell.Value) Then
            cell.Group _
                Start:=True, End:=True, Periods:=Array(False, False, False, True, False, False, False)
                Exit For
        End If
    Next cell
    
    pt2.AddDataField pt2.PivotFields("screen"), , xlCount
    
    'Column Field
    Set pf2 = pt2.PivotFields("screen")
            pf2.Orientation = xlColumnField
            pf2.Position = 1
    
    'Adding a Chart
    ws2.Activate
    Set ch2 = Charts.Add
        ch2.Name = chartName + " Count"
        ch2.SetSourceData Source:=pt2.TableRange1
        ch2.ChartType = xlColumnClustered
        
ElseIf PivotTableType = "Timers" Then

    'Row Fields
    Set pf = pt.PivotFields("instant")
            pf.Orientation = xlRowField
            pf.Position = 1
    
    'Column Fields
    Set pf = pt.PivotFields("cyclicjobname")
            pf.Orientation = xlColumnField
            pf.Position = 1
            
    Set pf = pt.PivotFields("espacename")
            pf.Orientation = xlColumnField
            pf.Position = 2
            
    pt.AddDataField pt.PivotFields("duration"), , xlAverage
    
    'Adding a Chart
    ws.Activate
    Set ch = Charts.Add
        ch.Name = chartName + " Performance Analysis"
        ch.SetSourceData pt.TableRange1
        ch.ChartType = xlLineMarkers
        
    'Row Field
    Set pf2 = pt2.PivotFields("instant")
            pf2.Orientation = xlRowField
            pf2.Position = 1
    
    'Column Fields
    Set pf2 = pt2.PivotFields("espacename")
            pf2.Orientation = xlColumnField
            pf2.Position = 1
            
    'Grouping the Instant by day
   For Each cell In ws2.UsedRange.Columns("A").Cells
        If IsDate(cell.Value) Then
            cell.Group _
                Start:=True, End:=True, Periods:=Array(False, False, False, True, False, False, False)
                Exit For
        End If
    Next cell
    
    pt2.AddDataField pt2.PivotFields("cyclicjobname"), , xlCount
    
    'Column Field
    Set pf2 = pt2.PivotFields("cyclicjobname")
            pf2.Orientation = xlColumnField
            pf2.Position = 1
    
    'Adding a Chart
    ws2.Activate
    Set ch2 = Charts.Add
        ch2.Name = chartName + " Count"
        ch2.SetSourceData Source:=pt2.TableRange1
        ch2.ChartType = xlColumnClustered

End If

Exit Sub

General_ErrorHandler:
MsgBox Err.Description, vbCritical


End Sub

Public Function PT_Table(Datasheet As Worksheet, PVNumber As Integer) As Variant

Dim pc As PivotCache
Dim pt As Variant

'Creating Pivot Cache
Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        sourceData:=Datasheet.Name & "!" & Datasheet.Range("A1").CurrentRegion.Address, _
        Version:=xlPivotTableVersion15)
        
'Debug.Print Datasheet.Name

Datasheet.Activate
Set pt = pc.CreatePivotTable( _
    TableDestination:="", _
    TableName:="MyPivotTable" + PVNumber)
    
Set PT_Table = pt

End Function


'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'                                                                       *** Format Logs ***
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Sub FormatLogs()

' #V2:
    '# Fix: headers with _ characters are replaced with blank spaces
    '# Fix: Some logs were missing headers. Added a resume next activity to avoid the trigger of exception
' #V2.1:
    '# Fix: set worksheet to the active (opened) worksheet instead of the hardcoded name "Sheet1"
' #V3:
    '# Fix an infinite loop when the headers are not on the first line
    '# One macro for all SC logs
    '# If Autofilter is already applied, then it's not removed
' #V3.1: Adapted to Platform 11 headers
' #v3.2: Delete headers space

    Dim RowHeaderRange As Range, nameHeader As Range, instantHeader As Range, messageHeader As Range, StackHeader As Range
    Dim moduleNameHeader As Range, RequestKeyHeader As Range, EspaceIdHeader As Range, ActionNameHeader As Range
    Dim endPointHeader As Range, actionHeader As Range, durationHeader As Range, ScreenHeader As Range, cell As Range
    Dim myWorksheet As Worksheet
    Dim Random As Variant
    
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
        cell.Value = LCase(Replace(cell.Value, "_", ""))
        cell.Value = LCase(Replace(cell.Value, " ", ""))
    Next cell
    
    'On error, it resumes next in case it does not find one or more of the headers below'
    On Error Resume Next
    'Finds Headers by exact match'
    'Common Headers
    Set instantHeader = RowHeaderRange.Cells.Find("instant", LookAt:=xlWhole, MatchCase:=False)
    instantHeader.EntireColumn.ColumnWidth = 20
    
    Set RequestKeyHeader = RowHeaderRange.Cells.Find("requestkey", LookAt:=xlWhole, MatchCase:=False)
    RequestKeyHeader.EntireColumn.ColumnWidth = 35
    
    Set nameHeader = RowHeaderRange.Cells.Find("name", LookAt:=xlWhole, MatchCase:=False)
    If nameHeader Is Nothing Then
        Set nameHeader = RowHeaderRange.Cells.Find("espacename", LookAt:=xlWhole, MatchCase:=False)
    End If
    nameHeader.EntireColumn.ColumnWidth = 20
    
    'Other Headers (General and Error mostly)
    Set ActionNameHeader = RowHeaderRange.Cells.Find("actionname", LookAt:=xlWhole, MatchCase:=False)
    ActionNameHeader.EntireColumn.ColumnWidth = 18

    Set messageHeader = RowHeaderRange.Cells.Find("message", LookAt:=xlWhole, MatchCase:=False)
    messageHeader.EntireColumn.ColumnWidth = 80
    
    Set StackHeader = RowHeaderRange.Cells.Find("stack", LookAt:=xlWhole, MatchCase:=False)
    StackHeader.EntireColumn.ColumnWidth = 40
    
    Set moduleNameHeader = RowHeaderRange.Cells.Find("modulename", LookAt:=xlWhole, MatchCase:=False)
    moduleNameHeader.EntireColumn.ColumnWidth = 20
    
    'Integration Headers
    Set endPointHeader = RowHeaderRange.Cells.Find("endpoint", LookAt:=xlWhole, MatchCase:=False)
    endPointHeader.EntireColumn.ColumnWidth = 90
    
    Set actionHeader = RowHeaderRange.Cells.Find("action", LookAt:=xlWhole, MatchCase:=False)
    actionHeader.EntireColumn.ColumnWidth = 90
    
    Set durationHeader = RowHeaderRange.Cells.Find("duration", LookAt:=xlWhole, MatchCase:=False)
    durationHeader.EntireColumn.ColumnWidth = 10
    
    'Screen and Mobile Headers
    Set ScreenHeader = RowHeaderRange.Cells.Find("screen", LookAt:=xlWhole, MatchCase:=False)
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


'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'                                                                       *** Integration Macro ***
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Sub Integration()

Dim wb As Workbook
Dim myWorksheet As Worksheet, DataProcessSheet As Worksheet
Dim RowHeaderRange As Range, durationHeader As Range, endPointHeader As Range, DataWrite As Range, actionHeader As Range, typeHeader As Range, sourceHeader As Range
Dim instantHeader As Range, DataWriteInstant As Range, nameHeader As Range
Dim InstantColumn As Range, cell As Range
Dim i As Integer, j As Integer
Dim execTimeValue As Double
Dim queryName As String

'# v1.1: Added Source, eSpaceName, Endpoint and Type to Filter Fields
'# v1.2: Adapted to Platform 11 headers
'# v1.2: Refactored for performance

Application.DisplayAlerts = False

'This call is to guarantee header uniformization
Call FormatLogs

Set wb = ActiveWorkbook

Set myWorksheet = ActiveWorkbook.ActiveSheet
Set RowHeaderRange = myWorksheet.Cells.Range(Range("A1"), Range("A1").End(xlToRight))

Set durationHeader = RowHeaderRange.Cells.Find("duration")
Set sourceHeader = RowHeaderRange.Cells.Find("source", LookAt:=xlWhole, MatchCase:=False)
Set endPointHeader = RowHeaderRange.Cells.Find("endpoint", LookAt:=xlWhole, MatchCase:=False)
Set actionHeader = RowHeaderRange.Cells.Find("action", LookAt:=xlWhole, MatchCase:=False)
Set typeHeader = RowHeaderRange.Cells.Find("type", LookAt:=xlWhole, MatchCase:=False)
Set instantHeader = RowHeaderRange.Cells.Find("instant", LookAt:=xlWhole, MatchCase:=False)

'Validation added for Platform 11
Set nameHeader = RowHeaderRange.Cells.Find("name", LookAt:=xlWhole, MatchCase:=False)
If Not nameHeader Is Nothing Then
    nameHeader.Value = "espacename"
ElseIf nameHeader Is Nothing Then
    Set nameHeader = RowHeaderRange.Cells.Find("espacename", LookAt:=xlWhole, MatchCase:=False)
End If

Set DataProcessSheet = myWorksheet

On Error GoTo General_ErrorHandler:

myWorksheet.Activate

Call CreatePVTable("Integration", DataProcessSheet, 2)

Exit Sub

General_ErrorHandler:
MsgBox Err.Description, vbCritical

End Sub


'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'                                                                       *** SLOWSQL ***
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Sub SlowSql(SlowSql_cb As Boolean, SlowExtension_cb As Boolean)

Dim wb As Workbook
Dim myWorksheet As Worksheet, DataProcessSheet As Worksheet
Dim messageHeader As Range, RowHeaderRange As Range, moduleNameHeader As Range, queryHeader As Range, execTimeHeader As Range, messageColumn As Range, DataWrite As Range
Dim instantHeader As Range, DataWriteInstant As Range, instantData As Range, nameHeader As Range, nameHeadProc As Range, modNameProc As Range, cell As Range
Dim i As Integer, j As Integer, k As Integer
Dim execTimeValue As Double
Dim queryName As String, eSpaceName As String, moduleName As String

'# v1.1: Added Slow Extensions and eSpace Names
'# v1.2: Validated if the Maximum Number of log entries was exceeded
'# v1.3: Added option to filter the logs for SlowSql, SlowExtension or both
'# v1.4: Adapted to Platform 11 headers;

Application.DisplayAlerts = False

'This call is to guarantee header uniformization
Call FormatLogs

Set wb = ActiveWorkbook

Set myWorksheet = ActiveWorkbook.ActiveSheet
Set RowHeaderRange = myWorksheet.Cells.Range(Range("A1"), Range("A1").End(xlToRight))

'Finding headers and their locations
Set messageHeader = RowHeaderRange.Cells.Find("message", LookAt:=xlWhole, MatchCase:=False)
Set moduleNameHeader = RowHeaderRange.Cells.Find("modulename", LookAt:=xlWhole, MatchCase:=False)
Set instantHeader = RowHeaderRange.Cells.Find("instant", LookAt:=xlWhole, MatchCase:=False)

Set nameHeader = RowHeaderRange.Cells.Find("name", LookAt:=xlWhole)
If nameHeader Is Nothing Then
    Set nameHeader = RowHeaderRange.Cells.Find("espacename", LookAt:=xlWhole, MatchCase:=False)
End If
    
' Autofilter for SlowSql and SlowExtension
If SlowSql_cb = True And SlowExtension_cb = True Then
    RowHeaderRange.AutoFilter Field:=moduleNameHeader.Column, Criteria1:=Array("SLOWSQL", "SLOWEXTENSION"), Operator:=xlFilterValues
ElseIf SlowSql_cb = True And SlowExtension_cb = False Then
    RowHeaderRange.AutoFilter Field:=moduleNameHeader.Column, Criteria1:="SLOWSQL", Operator:=xlFilterValues
ElseIf SlowSql_cb = False And SlowExtension_cb = True Then
    RowHeaderRange.AutoFilter Field:=moduleNameHeader.Column, Criteria1:="SLOWEXTENSION", Operator:=xlFilterValues
End If


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
    If Not (LCase(cell.Value) = "message" Or InStr(1, cell.Value, "The maximum number") > 0) Then
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

Call CreatePVTable("SLOW", DataProcessSheet, 2)

Exit Sub

General_ErrorHandler:
MsgBox Err.Description, vbCritical


End Sub
