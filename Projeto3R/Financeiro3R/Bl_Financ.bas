Attribute VB_Name = "Bl_Financ"
Option Explicit
''Windows API Types
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Const DFC_BUTTON = 4
Public Const DFCS_BUTTONPUSH = &H10      'Push button
Public Const DFCS_PUSHED = &H200         'Button is pushed
Public Declare Function DrawFrameControl Lib "user32" (ByVal hdc&, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Boolean
Public Declare Function InflateRect Lib "user32.dll" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function OffsetRect Lib "user32.dll" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Function TranslateColor(ByVal clr As Long) As Long
    If OleTranslateColor(clr, 0, TranslateColor) Then
         TranslateColor = -1
    End If
End Function



