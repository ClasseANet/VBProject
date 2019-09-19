Attribute VB_Name = "CalSubClass"
Option Private Module
Option Explicit
Public NextProcs As Long
Public Nodef As Boolean
 

Public Type NMHDR
    hwndFrom As Long
    idfrom As Long
    code As Long
End Type

Public Const WM_SIZE = &H5
'Public Const WM_NOTIFY = &H4E
'Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal uMgs As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Any) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal wndrpcPrev As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const GWL_USERDATA = (-21)
Public Const GWL_WNDPROC = -4
Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, _
                           ByVal wParam As Long, ByVal lParam As Long) As Long
   On Error Resume Next
   Select Case hWnd
      Case FrmMonth.hWnd
         FrmMonth.ProcMsg hWnd, uMsg, wParam, lParam, 0&  ', 0&
   End Select
   If Nodef = True Then
      WindowProc = CallWindowProc(NextProcs, hWnd, uMsg, wParam, ByVal lParam)
   Else
      Nodef = True
   End If
End Function


