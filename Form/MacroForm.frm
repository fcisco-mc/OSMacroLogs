VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MacroForm 
   Caption         =   "Please select the Macro to run"
   ClientHeight    =   2325
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6915
   OleObjectBlob   =   "MacroForm.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "MacroForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FormatLogs_Click()
    Call AllMacros.FormatErrorLogs
    Call UnloadForm
End Sub

Private Sub Integration_Click()
    Call AllMacros.Integration
    Call UnloadForm
End Sub

Private Sub MobileErrors_Click()
    Call AllMacros.DeviceUUID
    Call UnloadForm
End Sub


Private Sub SlowSql_Click()
    Call AllMacros.SlowSql
    Call UnloadForm
End Sub

Private Sub UnloadForm()
    Unload Me
End Sub
