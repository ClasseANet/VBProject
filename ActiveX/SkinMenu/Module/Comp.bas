Attribute VB_Name = "basComp"

'-------------------------------'
' R Development Library 2.0 '
'-------------------------------'
'    R Interface Components '
'                   Version 1.0 '
'-------------------------------'
'          Core Routines Module '
'-------------------------------'
'Copyright © 1998-9 by R Software. All Rights Reserved

'Date Created:
'Last Updated:

Option Explicit
DefInt A-Z

Public UC As Object
Public ctlID As Long

'PlaySoundA Constants
Public Const SND_ASYNC = &H1             '  play asynchronously
Public Const SND_NODEFAULT = &H2         '  silence not default, if sound not found
Public Const SND_MEMORY = &H4            '  lpszSoundName points to a memory file

' GetMapMode returns
Public Const MM_ANISOTROPIC = 8
Public Const MM_HIENGLISH = 5
Public Const MM_HIMETRIC = 3
Public Const MM_ISOTROPIC = 7
Public Const MM_LOENGLISH = 4
Public Const MM_LOMETRIC = 2
Public Const MM_TEXT = 1
Public Const MM_TWIPS = 6


Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function GetMapMode Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function PlaySoundData Lib "WINMM.DLL" Alias "PlaySoundA" (lpData As Any, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Declare Function ReleaseCapture& Lib "user32" ()
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetCapture& Lib "user32" (ByVal hwnd As Long)
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long

Public Const WM_SETREDRAW = &HB

Public Const SW_SHOWNOACTIVATE = 4

Private Const HWND_TOP& = 0
Private Const SWP_NOMOVE& = &H2
Private Const SWP_NOACTIVATE& = &H10
Private Const SWP_NOSIZE& = &H1
Private Const SWP_SHOWWINDOW& = &H40

Public PE As ascPaintEffects

Public CtlCount As Long
Public CtlHandle As Long

Public Const ASMAIL$ = "support@R.globalnet.co.uk"
Public Const ASURL$ = "http://www.users.globalnet.co.uk/~R/"
Public Const ASURL2$ = "http://www.R.tsx.org/"

Public Const INTERR$ = "An unexpected application error has occured!"
Public Const ERRTEXT$ = "If this problem continues, please contact R technical support, at " + ASMAIL$ + ", quoting the above information."


Public Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128
End Type

Public Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128      '  Maintenance string for PSS usage
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
End Type


Public Declare Function GetVersion Lib "kernel32" Alias "GetVersionExA" _
    (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
    (lpVersionInformation As OSVERSIONINFOEX) As Long

Function GetOS() As String
    
    Dim vi As OSVERSIONINFO
    Dim vix As OSVERSIONINFOEX
    Dim s As String
    Dim Sm As Long, osMajor As Long, osMinor As Long

    Const VER_PLATFORM_WIN32s = 0
    Const VER_PLATFORM_WIN32_WINDOWS = 1
    Const VER_PLATFORM_WIN32_NT = 2


    vi.dwOSVersionInfoSize = Len(vi)
    GetVersion vi

    osMajor = vi.dwMajorVersion
    osMinor = vi.dwMinorVersion
    
    If vi.dwPlatformId = VER_PLATFORM_WIN32_NT Then
        Select Case vi.dwMajorVersion
            Case Is >= 5:
                If vi.dwMinorVersion = 0 Then
                    s = "2000"
                Else
                    s = "XP"
                End If
                vix.dwOSVersionInfoSize = Len(vix)
                GetVersionEx vix
            Case 4:
                s = "NT4"
            Case Is = 3:
                s = "NT3.5x"
            Case Else:
                s = "NT"
        End Select
    Else
        Select Case vi.dwMajorVersion
            Case 3:
                s = "3." & CStr(vi.dwMinorVersion)
            Case 4:
                Select Case vi.dwMinorVersion
                    Case 0:
                        s = "95"
                    Case 10:
                        s = "98"
                    Case 90:
                        s = "ME"
                    Case Else:
                        s = "9x"
                End Select
            Case 5:
                s = "XP"
            Case Else:
                s = "Unknown"
        End Select
    End If
    
    GetOS = s

End Function


Function CompareBoxes(Box1 As RECT, Box2 As RECT) As Integer
  
  ' Verifica il tipo di relazione tra i due rettangoli
  ' Restituisce True se i due box si sovrappongono in qualche modo
  
  Dim c As Boolean
    
    
  c = Box2.Right < Box1.Left Or Box2.Bottom < Box1.Top Or _
                 Box2.Left > Box1.Right Or Box2.Top > Box1.Bottom
    
  ' c risulta True se nessun lato si interseca, quindi, il test è
  ' positivo solo quando c=False
  
  If c Then
     CompareBoxes = 0 ' Box2 è completamente esterno all'altro
  Else
     c = Box2.Right < Box1.Right And Box2.Bottom < Box1.Bottom And _
                 Box2.Left > Box1.Left Or Box2.Top > Box1.Top
     If c Then
        CompareBoxes = 1 ' Box2 è interamente contenuto in Box1
     Else
        CompareBoxes = 2 ' I due Box intersecano
     End If
  End If
  
End Function




Function LINE_DISTANCE(PX As Long, PY As Long, x1 As Long, y1 As Long, X2 As Long, Y2 As Long) As Long
 

' Restituisce la distanza di un punto da una Linea
' P = Punto
' x1,y1 Primo vertice della linea
' x2,y2 Secondo Vertice della Linea

 Dim x0 As Double, y0 As Double, Dx As Double, Dy As Double

 x0 = PX
 y0 = PY
 
 If (x1 = X2) Then
  
  LINE_DISTANCE = Abs(x1 - x0)
 
 ElseIf y1 = Y2 Then
  
  LINE_DISTANCE = Abs(y1 - y0)
 
 Else
 
  Dx = X2 - x1
  Dy = Y2 - y1
  
  LINE_DISTANCE = Abs(Dy * x0 - Dx * y0 + X2 * y1 - x1 * Y2) / Sqr(Dx * Dx + Dy * Dy)
 
 End If

End Function


'-------------------------------
'Name        : ShowPopupMenu
'Created     : 27/08/1999 14:39
'-------------------------------
'Author      : Richard Moss
'Organisation: R Software
'-------------------------------
'Returns     : Nothing
'
'-------------------------------
'Updates     :
'
'-------------------------------
'---------AS-PROCBUILD 1.00.0024
Public Sub ShowPopupMenu(hWndClient As Long, PopupMenu As Menu, PopupParent As Form)
 Dim WinRect As RECT
 Dim WinPoint As POINTAPI
 Dim X As Single, Y As Single
 Dim ScaleMode As ScaleModeConstants
 ClientToScreen PopupParent.hwnd, WinPoint
 GetWindowRect hWndClient, WinRect
 If TypeOf PopupParent Is MDIForm Then
  ScaleMode = vbTwips
 Else
  ScaleMode = PopupParent.ScaleMode
 End If
 X = PopupParent.ScaleX(WinRect.Left - WinPoint.X, vbPixels, ScaleMode)
 Y = PopupParent.ScaleY(WinRect.Bottom - WinPoint.Y, vbPixels, ScaleMode)
 PopupParent.PopupMenu PopupMenu, , X, Y
End Sub '(Public) Sub ShowPopupMenu ()

'----------------------------------------------------------------------
'Name        : Highlight
'Created     : 21/08/1999 23:07
'Modified    :
'Modified By :
'----------------------------------------------------------------------
'Author      : Richard James Moss
'Organisation: R Software
'----------------------------------------------------------------------
Public Sub HighLight(c As Control)
 With c
  .SelStart = 0
  .SelLength = Len(.Text)
 End With
End Sub '(Public) Sub Highlight ()

'----------------------------------------------------------------------
'Name        : InitPaintEffects
'Created     : 12/07/1999 14:51
'Modified    :
'Modified By :
'----------------------------------------------------------------------
'Author      : Richard James Moss
'Organisation: R Software
'----------------------------------------------------------------------
Public Sub InitPaintEffects()
 If PE Is Nothing Then
  Set PE = New ascPaintEffects
 End If
End Sub '(Public) Sub InitPaintEffects ()


'----------------------------------------------------------------------
'Name        : Main
'Created     : 12/07/1999 14:40
'Modified    :
'Modified By :
'----------------------------------------------------------------------
'Author      : Richard James Moss
'Organisation: R Software
'----------------------------------------------------------------------
Public Sub Main()
 Set PE = New ascPaintEffects
End Sub '(Public) Sub Main ()

Function StartDocError$(R As Long)
 Dim M$
 If R >= 0 Then
  Select Case R
   Case 0: M$ = "System was out of memory or executable file was corrupt."
   Case 2: M$ = "The file was not found."
   Case 3: M$ = "The path was not found."
   Case 5: M$ = "Attempt was made to link to a task dynamically, or there was a sharing or network-protection error."
   Case 6: M$ = "Library required separate data segments for each task."
   Case 8: M$ = "There was insufficient memory to start the application."
   Case 10: M$ = "The Windows version was incorrect."
   Case 11: M$ = "The executable file was invalid. Either it was not a Windows-based application or there was an error in the .EXE image."
   Case 12: M$ = "Application was designed for a different operating system."
   Case 13: M$ = "Application was designed for MS-DOS version 4.0."
   Case 14: M$ = "Type of executable file was unknown."
   Case 15: M$ = "Attempt was made to load a real-mode application that was developed for an earlier version of Windows."
   Case 16: M$ = "Attempt was made to load a second instance of an executable file containing multiple data segments not marked read-only."
   Case 19: M$ = "Attempt was made to load a compressed executable file. The file must be decompressed before it can be loaded."
   Case 20: M$ = "Dynamic-link library (DLL) file was invalid. One of the DLLs required to run this application was corrupt."
   Case 21: M$ = "Application requires Microsoft Windows 32-bit extensions."
   Case 31: M$ = "No application has been associated for use with specified document."
   Case Else: M$ = "Unknown Error."
  End Select
 Else
  M$ = "Unknown error."
 End If
 StartDocError$ = M$ + Chr$(10) + Chr$(10) + "(Error Code: " + CStr(R) + ")"
End Function

Function IsUsingLargeFonts() As Boolean
 Dim hWndDesk As Long, hDCDesk As Long, logPix As Long, R As Long
 hWndDesk = GetDesktopWindow()
 hDCDesk = GetDC(hWndDesk)
 logPix = GetDeviceCaps(hDCDesk, 88)
 R = ReleaseDC(hWndDesk, hDCDesk)
 If logPix > 96 Then IsUsingLargeFonts = -1
End Function

Function DegreeToRad(Deg As Integer) As Single
 DegreeToRad = Deg / 57.295779513
End Function

Public Function RemoveExtension$(F$)
 Dim R$(), E$
 Dim i
 If InStr(F$, ".") Then
  R$ = Split(F$, ".")
  For i = 0 To UBound(R$) - 1
   E$ = E$ + R$(i) + "."
  Next
  RemoveExtension$ = Left$(E$, Len(E$) - 1)
 Else
  RemoveExtension$ = F$
 End If
End Function

Function IsInControl(ByVal hwnd As Long) As Boolean
 Dim P As POINTAPI
 GetCursorPos P
 If hwnd = WindowFromPoint(P.X, P.Y) Then IsInControl = -1
End Function

Public Function GetFile$(FP$)
 Dim R$()
 If Len(FP$) Then
  R$() = Split(FP$, "\")
  GetFile$ = R$(UBound(R$))
 End If
End Function

Sub PlaySnd(SndName$, m_PlaySounds As Boolean)
 Dim bySound() As Byte
 On Error Resume Next
  If m_PlaySounds Then
   bySound = LoadResData(SndName$, 100)
   If Err = 0 And UBound(bySound) > 0 Then
    PlaySoundData bySound(0), 0, SND_MEMORY + SND_ASYNC + SND_NODEFAULT
   End If
  End If
 On Error GoTo 0
End Sub

'Public Function ShowTip(ByVal Tip$, ByVal hwnd As Long, Optional ByVal Font As StdFont) As Boolean
' Const Dx = -2   ' Offset from the mouse position.
' Const Dy = 18
' Dim X As Long, Y As Long
' Dim pt As POINTAPI
' On Error Resume Next
'  GetCursorPos pt
'  X = pt.X
'  Y = pt.Y
'  HideTip
'  With frmTooltip
'   If Not Font Is Nothing Then
'    Set .lblTip.Font = Font
'    Set .Font = Font
'   End If
'   .lblTip.Width = .TextWidth(Tip$)
'   .lblTip.Caption = Tip$
'   .lblTip.Refresh
'   .CtlHWnd = hwnd
'   .Move (X + Dx) * Screen.TwipsPerPixelX, (Y + Dy) * Screen.TwipsPerPixelY, .lblTip.Width + (8 * Screen.TwipsPerPixelX), .lblTip.Height + (5 * Screen.TwipsPerPixelY)
'   .tmrTip.Enabled = 0
'   .tmrTip.Enabled = -1
'   If .Left + .Width > Screen.Width Then .Left = Screen.Width - .Width
'   If .Top + .Height > Screen.Height Then .Top = Screen.Height - .Height
'   SetWindowPos .hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
'  End With
'  ShowTip = -1
' On Error GoTo 0
'End Function

Function DefineAccessKey$(Caption$)
 Dim P, n
 Dim c$
 n = 1
 Do
  P = InStr(n, Caption$, "&")
  If P Then
   c$ = Mid$(Caption$, P + 1, 1)
   If c$ <> "&" Then DefineAccessKey$ = DefineAccessKey$ + c$
   n = P + 1
  End If
 Loop Until P = 0
End Function


Public Sub HideTip()
 On Error Resume Next
  'Unload frmTooltip
 On Error GoTo 0
End Sub


Public Sub Pointer(V)
 Screen.MousePointer = V
End Sub



Public Function UltimateParent(Ctl As Object) As Object
 Dim o As Object, T As Object
 On Error Resume Next
  Set T = Ctl.Parent
  Set UltimateParent = T
  Do
   Set o = T.Parent
   If Not o Is Nothing Then
    Set T = o
    Set UltimateParent = o
   End If
  Loop Until o Is Nothing
 On Error GoTo 0
End Function

