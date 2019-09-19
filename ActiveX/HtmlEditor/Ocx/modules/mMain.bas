Attribute VB_Name = "mMain"
Option Explicit

Public AppPath As String

Sub Main()

    '------------------------------------------------------
    'Set the application path
    AppPath = App.Path
    If Right(AppPath, 1) <> "\" Then AppPath = AppPath & "\"
    '------------------------------------------------------
End Sub


