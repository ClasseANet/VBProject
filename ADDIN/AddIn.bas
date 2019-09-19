Attribute VB_Name = "AddIn"
Option Explicit
Global DB      As New DS_BANCO
Global Sys     As New SETTING
Global AddIn   As New ADDINS
Global glbProj As VBProject

Global VBInstance   As VBIDE.VBE       'instance of VB IDE
'Global gWindow   As VBIDE.Window    'used to make sure we only run one instance
Global gWindow   As Form    'used to make sure we only run one instance

Declare Function WritePrivateProfileString& Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FILENAME$)
Private Declare Sub PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd&, ByVal msg&, ByVal wp&, ByVal lp&)
Private Declare Sub SetFocus Lib "user32" (ByVal hwnd&)
Private Declare Function GetParent Lib "user32" (ByVal hwnd&) As Long
Global Const WM_SYSKEYDOWN = &H104
Global Const WM_SYSKEYUP = &H105
Global Const WM_SYSCHAR = &H106
Global Const VK_F = 70  ' VK_A thru VK_Z are the same as their ASCII equivalents: 'A' thru 'Z'
Global Const APP_CATEGORY = "Microsoft Visual Basic AddIns"
Dim hwndMenu       As Long           'needed to pass the menu keystrokes to VB
'********************************************************************
'* This sub should be executed from the Immediate window in         *
'* order to get this app added to the VBADDIN.INI  file you         *
'* you must change the name in the 2nd argument to reflecty         *
'* the correct name of your project                                 *
'********************************************************************
Public Sub AddToINI()
    Dim ErrCode As Long
    ErrCode = WritePrivateProfileString("Add-Ins32", "Construtor.Connect", "0", "vbaddin.ini")
End Sub
Function InRunMode(VBInst As VBIDE.VBE) As Boolean
  InRunMode = (VBInst.CommandBars("File").Controls(1).Enabled = False)
End Function

Sub HandleKeyDown(ud As Object, KeyCode As Integer, Shift As Integer)
  If Shift <> 4 Then Exit Sub
  If KeyCode < 65 Or KeyCode > 90 Then Exit Sub
  If VBInstance.DisplayModel = vbext_dm_SDI Then Exit Sub
  
  If hwndMenu = 0 Then hwndMenu = FindHwndMenu(ud.hwnd)
  PostMessage hwndMenu, WM_SYSKEYDOWN, KeyCode, &H20000000
  KeyCode = 0
  SetFocus hwndMenu
End Sub

Function FindHwndMenu&(ByVal hwnd&)
  Dim h As Long
  
Loop2:
  h = GetParent(hwnd)
  If h = 0 Then FindHwndMenu = hwnd: Exit Function
  hwnd = h
  GoTo Loop2
End Function



Public Function PalavraReservada(pTxt As String) As Boolean
   pTxt = Trim(UCase(pTxt))
   PalavraReservada = True
   Select Case True
      Case InArray(pTxt, Array("IF", "THEN", "ELSE", "END", "SELECT", "CASE", "DO", "LOOP", "UNTIL"))
      Case InArray(pTxt, Array("FOR", "EACH", "IN", "TO", "NEXT", "WHILE", "WEND", "WITH", "EXIT"))
      Case InArray(pTxt, Array("GLOBAL", "PUBLIC", "PRIVATE", "LOCAL", "DIM", "SUB", "FUNCTION"))
      Case InArray(pTxt, Array("STATIC", "DIM", "REDIM", "CONST", "ENUM", "PRESERV"))
      Case InArray(pTxt, Array("BOOLEAN", "STRING", "LONG", "INTEGER", "VARIANT", "DATE"))
      Case InArray(pTxt, Array("SINGLE", "BYTE", "CURRENCY", "DOUBLE", "OBJECT"))
      Case InArray(pTxt, Array("TRUE", "FALSE", "AND", "OR", "NOT"))
      Case InArray(pTxt, Array("SET", "NEW", "NOTHING", "EMPTY"))
      Case InArray(pTxt, Array("AS", "ON", "STEP", "GOTO", "ERROR", "NOTHING", "EMPTY"))
      Case InArray(pTxt, Array("PRINT", "INPUT", "OUTPUT", "OPEN", "CLOSE", "LINE"))
      Case Else: PalavraReservada = False
   End Select
End Function



