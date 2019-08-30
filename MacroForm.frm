VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MacroForm 
   Caption         =   "Please select the Macro to run"
   ClientHeight    =   8670
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8610
   OleObjectBlob   =   "MacroForm.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "MacroForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub FormatLogs_Click()
    Call AllMacros.FormatLogs
    Call UnloadForm
End Sub

Private Sub Integration_Click()
    Call AllMacros.Integration
    Call UnloadForm
End Sub


Private Sub MobileErrors_Click()
    If iOS_cb = False And Android_cb = False Then
        Call InvalidCall
        Exit Sub
    End If
    
    Call AllMacros.DeviceUUID(iOS_cb.Value, Android_cb.Value)
    Call UnloadForm
End Sub

Private Sub Screen_Click()
    Call AllMacros.Screens
    Call UnloadForm
End Sub

Private Sub SlowSql_Click()
    If SlowSql_cb = False And SlowExtension_cb = False Then
        Call InvalidCall
        Exit Sub
    End If
    
    Call AllMacros.SlowSql(SlowSql_cb.Value, SlowExtension_cb.Value)
    Call UnloadForm
    
End Sub

Private Sub UnloadForm()
    Unload Me
End Sub

Private Sub InvalidCall()
    MsgBox "Select at least one of the checkboxes", vbInformation
End Sub

Private Sub Timers_Click()
    Call AllMacros.Timers
    Call UnloadForm
End Sub
