Attribute VB_Name = "DeviceUUID"
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
