Attribute VB_Name = "Declara"
Option Explicit
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function SHAutoComplete Lib "Shlwapi.dll" (ByVal hWndEdit As Long, ByVal dwFlags As Long) As Long

Public Const SHACF_DEFAULT  As Long = &H0 ' Currently (SHACF_FILESYSTEM | SHACF_URLALL)
Public Const SHACF_FILESYSTEM As Long = &H1 ' This includes the File System as well as the rest of the shell (Desktop\My Computer\Control Panel\)
Public Const SHACF_URLHISTORY As Long = &H2 ' URLs in the User's History
Public Const SHACF_URLMRU As Long = &H4 ' URLs in the User's Recently Used list.
Public Const SHACF_URLALL As Long = (SHACF_URLHISTORY Or SHACF_URLMRU) ' Both File System and URLs in the User's History


