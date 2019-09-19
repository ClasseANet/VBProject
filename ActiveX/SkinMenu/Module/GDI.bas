Attribute VB_Name = "basGDI"

Option Explicit
DefInt A-Z

Type RECT
 Left       As Long
 Top        As Long
 Right      As Long
 Bottom     As Long
End Type

Type POINTAPI
 X As Long
 Y As Long
End Type

Type LPSIZE
 cx As Long
 cy As Long
End Type


Public Type XFORM
        eM11 As Double
        eM12 As Double
        eM21 As Double
        eM22 As Double
        eDx As Double
        eDy As Double
End Type


Public Type PointXY
  X As Long
  Y As Long
End Type

Public Const RGN_AND = 1
Public Const RGN_COPY = 5
Public Const RGN_DIFF = 4
Public Const RGN_OR = 2
Public Const RGN_XOR = 3

Public Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As PointXY, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long



Public Declare Function GetWorldTransform Lib "gdi32" (ByVal hdc As Long, lpXform As XFORM) As Long

Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetCurrentPositionEx Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetWindowExtEx Lib "gdi32" (ByVal hdc As Long, LPSIZE As LPSIZE) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function CreatePen& Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long)
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function DeleteObject& Lib "gdi32" (ByVal hObject As Long)
Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Declare Function DrawFocusRect& Lib "user32" (ByVal hdc As Long, lpRect As RECT)
Declare Function DrawFrameControl Lib "user32" (ByVal hdc&, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Boolean
Declare Function DrawText& Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long)
Declare Function FillRect& Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long)
Declare Function GetBkColor& Lib "gdi32" (ByVal hdc As Long)
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Declare Function GetTextColor& Lib "gdi32" (ByVal hdc As Long)
Declare Function LineTo& Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long)
Declare Function MoveToEx& Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI)
Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Declare Function SelectObject& Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long)
Declare Function SetTextColor& Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long)
Declare Function SetTextJustification Lib "gdi32" (ByVal hdc As Long, ByVal nBreakExtra As Long, ByVal nBreakCount As Long) As Long
Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Declare Function UpdateWindow& Lib "user32" (ByVal hwnd As Long)
Public Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long

'  flags for DrawFrameControl
Public Const DFC_CAPTION = 1 'Title bar
Public Const DFC_MENU = 2   'Menu
Public Const DFC_SCROLL = 3 'Scroll bar
Public Const DFC_BUTTON = 4 'Standard button

Public Const DFCS_CAPTIONCLOSE = &H0    'Close button
Public Const DFCS_CAPTIONMIN = &H1 'Minimize button
Public Const DFCS_CAPTIONMAX = &H2 'Maximize button
Public Const DFCS_CAPTIONRESTORE = &H3  'Restore button
Public Const DFCS_CAPTIONHELP = &H4     'Windows 95 only: Help button

Public Const DFCS_MENUARROW = &H0 'Submenu arrow
Public Const DFCS_MENUCHECK = &H1 'Check mark
Public Const DFCS_MENUBULLET = &H2 'Bullet
Public Const DFCS_MENUARROWRIGHT = &H4

Public Const DFCS_SCROLLUP = &H0   'Up arrow of scroll bar
Public Const DFCS_SCROLLDOWN = &H1 'Down arrow of scroll bar
Public Const DFCS_SCROLLLEFT = &H2 'Left arrow of scroll bar
Public Const DFCS_SCROLLRIGHT = &H3 'Right arrow of scroll bar

Public Const DFCS_SCROLLCOMBOBOX = &H5   'Combo box scroll bar
Public Const DFCS_SCROLLSIZEGRIP = &H8   'Size grip
Public Const DFCS_SCROLLSIZEGRIPRIGHT = &H10   'Size grip in bottom-right corner of window

Public Const DFCS_BUTTONCHECK = &H0 'Check box
Public Const DFCS_BUTTONRADIO = &H4 'Radio button
Public Const DFCS_BUTTON3STATE = &H8 'Three-state button
Public Const DFCS_BUTTONPUSH = &H10 'Push button
Public Const DFCS_INACTIVE = &H100 'Button is inactive (grayed)
Public Const DFCS_PUSHED = &H200  'Button is pushed
Public Const DFCS_CHECKED = &H400 'Button is checked
Public Const DFCS_ADJUSTRECT = &H2000   'Bounding rectangle is adjusted to exclude the surrounding edge of the push button
Public Const DFCS_FLAT = &H4000   'Button has a flat border
Public Const DFCS_MONO = &H8000   'Button has a monochrome border

Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENOUTER = &H2
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_OUTER = &H3
Public Const BDR_INNER = &HC
Public Const BDR_RAISED = &H5
Public Const BDR_SUNKEN = &HA

Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)

Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8
Public Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Public Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Public Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Public Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_DIAGONAL = &H10

' For diagonal lines, the BF_RECT flags specify the end point of
' the vector bounded by the rectangle parameter.
Public Const BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
Public Const BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
Public Const BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
Public Const BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)

Public Const BF_MIDDLE = &H800    ' Fill in the middle.
Public Const BF_SOFT = &H1000     ' Use for softer buttons.
Public Const BF_ADJUST = &H2000   ' Calculate the space left over.
Public Const BF_FLAT = &H4000     ' For flat rather than 3-D borders.
Public Const BF_MONO = &H8000     ' For monochrome borders.

'DrawText Constants
Public Const DT_BOTTOM = &H8
Public Const DT_CALCRECT = &H400
Public Const DT_CENTER = &H1
Public Const DT_LEFT = &H0
Public Const DT_NOCLIP = &H100
Public Const DT_NOPREFIX = &H800
Public Const DT_RIGHT = &H2
Public Const DT_SINGLELINE = &H20
Public Const DT_TOP = &H0
Public Const DT_VCENTER = &H4
Public Const DT_WORDBREAK = &H10

Public pt As POINTAPI

Sub GetRoundedRect(pt() As POINTAPI, x1 As Single, y1 As Single, X2 As Single, Y2 As Single, n As Long)
   Erase pt: n = 0
   AddPt pt, x1 + 2, y1
   AddPt pt, X2 - 2, y1
   AddPt pt, X2 - 1, y1 + 1
   AddPt pt, X2, y1 + 2
   AddPt pt, X2, Y2 - 2
   AddPt pt, X2 - 1, Y2 - 1
   AddPt pt, X2 - 2, Y2
   AddPt pt, x1 + 2, Y2
   AddPt pt, x1 + 1, Y2 - 1
   AddPt pt, x1, Y2 - 2
   AddPt pt, x1, y1 + 2
   AddPt pt, x1 + 1, y1 + 1
   AddPt pt, x1 + 2, y1
      
   n = UBound(pt)

End Sub


Private Sub AddPt(pt() As POINTAPI, X As Single, Y As Single)
    Dim n As Long
    On Error Resume Next
    n = UBound(pt)
    n = n + 1
    ReDim Preserve pt(n)
    pt(n).X = X
    pt(n).Y = Y
End Sub


Sub GradateFill(Obj As Object, x1 As Long, y1 As Long, X2 As Long, Y2 As Long, FromColor As Long, ToColor As Long, Optional Vertical As Long = 0, Optional DoubleFade As Long = 1)
   Dim c() As Long, Y As Long, n As Long, X As Long
   
If Vertical = 0 Then ' gradient orizzontale
   n = X2 - x1
   ReDim c(n)
   
   CreateGradateColors FromColor, ToColor, c
   
   For X = 0 To n
       Obj.Line (x1 + X, y1)-(x1 + X, Y2), c(X)
   Next
Else ' verticale a tubo
   If DoubleFade > 0 Then n = (Y2 - y1) * 0.5 Else n = Y2 - y1
   ReDim c(n)
 '  If Vertical = 2 Then x = n Else x = 0
   CreateGradateColors FromColor, ToColor, c
   
   For Y = 0 To n
       Obj.Line (x1, y1 + Y)-(X2, y1 + Y), c(Y)
   Next
  If DoubleFade > 0 Then
   For Y = 0 To n
       Obj.Line (x1, Y2 - Y)-(X2, Y2 - Y), c(Y)
   Next
  End If
   If Vertical = 2 Then
        Obj.FillStyle = 0
        Obj.FillColor = FromColor
        X = (Y2 - y1) * 0.5
        Obj.Circle (x1 - 6, y1 + X), X, FromColor
        Obj.FillColor = ToColor
        Obj.Circle (X2 + 6, y1 + X), X, ToColor
        Obj.FillStyle = 1
   End If
   If Vertical = 3 Then
        Obj.FillStyle = 0
        Obj.FillColor = FromColor
        X = (Y2 - y1) * 0.5
              Obj.DrawWidth = 10
              Obj.PSet (x1, y1), Obj.BackColor
              Obj.PSet (X2, y1), Obj.BackColor
              Obj.PSet (x1, Y2), Obj.BackColor
              Obj.PSet (X2, Y2), Obj.BackColor
              Obj.DrawWidth = 1
        
       ' Obj.Circle (x1 - 6, y1 + x), x, FromColor
        Obj.FillColor = ToColor
       ' Obj.Circle (x2 + 6, y1 + x), x, ToColor
        Obj.FillStyle = 1
   End If

End If
   
End Sub

Public Sub DrawCtlEdge(hdc As Long, X As Single, Y As Single, w As Single, H As Single, Optional Style As Long = EDGE_RAISED, Optional Flags As Long = BF_RECT)
 Dim R As RECT
 With R
  .Left = X
  .Top = Y
  .Right = X + w
  .Bottom = Y + H
 End With
 DrawEdge hdc, R, Style, Flags
End Sub

Public Function DrawControl(ByVal hdc As Long, ByVal X As Single, ByVal Y As Single, ByVal w As Single, ByVal H As Single, ByVal CtlType As Long, ByVal Flags As Long)
 Dim R As RECT
 With R
  .Left = X
  .Top = Y
  .Right = X + w
  .Bottom = Y + H
 End With
 DrawControl = DrawFrameControl(hdc, R, CtlType, Flags)
End Function

Function TranslateColor(ByVal clr As OLE_COLOR, Optional hPal As Long = 0) As Long
 If OleTranslateColor(clr, hPal, TranslateColor) Then TranslateColor = -1
End Function
Sub GetRGB(Rg As Long, Red As Long, Green As Long, Blue As Long)
 
    Blue = Fix(Rg / 65536)
    Green = Fix((Rg - (65536 * Blue)) / 256)
    Red = Fix(Rg - ((Blue * 65536) + (Green * 256)))
 
End Sub


Public Sub CreateGradateColors(BackColor As Long, ForeColor As Long, ColorRange() As Long)

' definisce 256 gradazioni partendo da BackColor fino a ForeColor

Dim foo As Long
Dim dblG As Double, dblR As Double, dblB As Double
Dim addG As Double, addR As Double, addB As Double
Dim bckR As Double, bckG As Double, bckB As Double
dblR = CDbl(BackColor And &HFF)
dblG = CDbl(BackColor And &HFF00&) / 255
dblB = CDbl(BackColor And &HFF0000) / &HFF00&
bckR = CDbl(ForeColor And &HFF&)
bckG = CDbl(ForeColor And &HFF00&) / 255
bckB = CDbl(ForeColor And &HFF0000) / &HFF00&
addR = (bckR - dblR) / UBound(ColorRange)
addG = (bckG - dblG) / UBound(ColorRange)
addB = (bckB - dblB) / UBound(ColorRange)

For foo = 0 To UBound(ColorRange)
    dblR = dblR + addR
    dblG = dblG + addG
    dblB = dblB + addB
    If dblR > 255 Then dblR = 255
    If dblG > 255 Then dblG = 255
    If dblB > 255 Then dblB = 255
    If dblR < 0 Then dblR = 0
    If dblG < 0 Then dblG = 0
    If dblB < 0 Then dblB = 0
    ColorRange(foo) = RGB(dblR, dblG, dblB)
Next foo

End Sub




Public Function LineDC(ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, Optional Color As OLE_COLOR = -1) As Long
 Dim hPen As Long, hPenOld As Long
 Dim R
 hPen = CreatePen(0, 1, IIf(Color = -1, GetTextColor(hdc), TranslateColor(Color)))
 hPenOld = SelectObject(hdc, hPen)
 MoveToEx hdc, x1, y1, pt
 LineDC = LineTo(hdc, X2, Y2)
 SelectObject hdc, hPenOld
 DeleteObject hPen
 DeleteObject hPenOld
End Function

Public Sub Box3DDC(ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal w As Long, ByVal H As Long, Optional HighLight As OLE_COLOR = vb3DHighlight, Optional Shadow As OLE_COLOR = vb3DShadow, Optional Fill As OLE_COLOR = -1)
 Dim hPen As Long, hPenOld As Long
 'Fill
 If Fill <> -1 Then BoxSolidDC hdc, X, Y, w, H, Fill
 'Highlight
 hPen = CreatePen(0, 1, TranslateColor(HighLight))
 hPenOld = SelectObject(hdc, hPen)
 MoveToEx hdc, X + w - 1, Y, pt
 LineTo hdc, X, Y
 LineTo hdc, X, Y + H - 1
 SelectObject hdc, hPenOld
 DeleteObject hPen
 DeleteObject hPenOld
 'Shadow
 hPen = CreatePen(0, 1, TranslateColor(Shadow))
 hPenOld = SelectObject(hdc, hPen)
 LineTo hdc, X + w - 1, Y + H - 1
 LineTo hdc, X + w - 1, Y
 SelectObject hdc, hPenOld
 DeleteObject hPen
 DeleteObject hPenOld
End Sub
Public Sub BoxDC(ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal w As Long, ByVal H As Long, Optional Color As OLE_COLOR = vbButtonFace, Optional Fill As OLE_COLOR = -1)
 Dim hPen As Long, hPenOld As Long
 'Fill
 If Fill <> -1 Then BoxSolidDC hdc, X, Y, w, H, Fill
 'Box
 hPen = CreatePen(0, 1, TranslateColor(Color))
 hPenOld = SelectObject(hdc, hPen)
 MoveToEx hdc, X + w - 1, Y, pt
 LineTo hdc, X, Y
 LineTo hdc, X, Y + H - 1
 LineTo hdc, X + w - 1, Y + H - 1
 LineTo hdc, X + w - 1, Y
 SelectObject hdc, hPenOld
 DeleteObject hPen
 DeleteObject hPenOld
End Sub

Public Function BoxSolidDC(ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal w As Long, ByVal H As Long, Optional ByVal Fill As OLE_COLOR = vbButtonFace)
 Dim hBrush As Long
 Dim R As RECT
 hBrush = CreateSolidBrush(TranslateColor(Fill))
 With R
  .Left = X
  .Top = Y
  .Right = X + w - 1
  .Bottom = Y + H - 1
 End With
 FillRect hdc, R, hBrush
 DeleteObject hBrush
End Function

Public Sub BoxRect3DDC(ByVal hdc As Long, R As RECT, Optional HighLight As OLE_COLOR = vb3DHighlight, Optional Shadow As OLE_COLOR = vb3DShadow, Optional Fill As OLE_COLOR = -1)
 Box3DDC hdc, R.Left, R.Top, R.Right - R.Left, R.Bottom - R.Top, HighLight, Shadow, Fill
End Sub

Public Sub PaintText(ByVal hdc As Long, ByVal Text$, ByVal X As Single, ByVal Y As Single, ByVal w As Single, ByVal H As Single, Optional ByVal Flags As Long = DT_LEFT)
 Dim R As RECT
 With R
  .Left = X
  .Top = Y
  .Right = X + w
  .Bottom = Y + H
 End With
 DrawText hdc, Text$, -1, R, Flags
End Sub


Public Sub DrawFocus(ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal w As Long, ByVal H As Long)
 Dim R As RECT
 With R
  .Left = X
  .Top = Y
  .Right = X + w
  .Bottom = Y + H
 End With
 DrawFocusRect hdc, R
End Sub

