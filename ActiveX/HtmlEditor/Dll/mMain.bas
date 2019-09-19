Attribute VB_Name = "mMain"
'==========================================================
'           Copyright Information
'==========================================================
'Program Name: Mewsoft Visual Html Editor
'Program Author   : Elsheshtawy, A. A.
'Home Page        : http://www.mewsoft.com
'Copyrights © 2006 Mewsoft Corporation. All rights reserved.
'==========================================================
'==========================================================
Option Explicit
Public ClEditor As EDITOR
Public sDocFileName As String
Public AppPath As String
Public CurrentMode As Long

Public Const AppName = "CAHtm"

Sub Main()

    '------------------------------------------------------
    'Set the application path
    AppPath = App.Path
    If Right(AppPath, 1) <> "\" Then AppPath = AppPath & "\"
    '------------------------------------------------------
    Set ClEditor = New EDITOR
    ClEditor.Show
End Sub


