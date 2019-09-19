VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ascPaintEffects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Provides methods for painting transparent and disabled looking images."
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'-----------------------------------------'
'            R Interface Components-1 '
'                             Version 2.2 '
'-----------------------------------------'
'              Paint Effects Class Module '
'                             Version 1.0 '
'-----------------------------------------'
'   Based on original sample by Microsoft '
'-----------------------------------------'
'Copyright © 1999 by R Software. All Rights Reserved.


'Created        : 06/30/1999
'Completed      :
'Last Updated   :

Option Explicit
DefInt A-Z

Private Type RECT
 Left       As Long
 Top        As Long
 Right      As Long
 Bottom     As Long
End Type

Private Type BITMAPINFOHEADER '40 bytes
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

Private Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Private Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors(1) As RGBQUAD
End Type

Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function CreateHalftonePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function SetDIBColorTable Lib "gdi32" (ByVal hdc As Long, ByVal un1 As Long, ByVal un2 As Long, pcRGBQuad As RGBQUAD) As Long
Private Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
Private Declare Function GetMapMode Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyHeight As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
'kdq 10/19/98 added for monochrome look on bitmap
' DrawState used for greyscale conversions
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long


'DrawIconEx Flags
Private Const DI_MASK = &H1
Private Const DI_IMAGE = &H2
Private Const DI_NORMAL = &H3
Private Const DI_COMPAT = &H4
Private Const DI_DEFAULTSIZE = &H8

'DIB Section constants
Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0 '  color table in RGBs


'Raster Operation Codes
Private Const DSna = &H220326 '0x00220326

'VB Errors
Private Const giINVALID_PICTURE As Integer = 481

'kdq 10/19/98 added for drawstate api
' DrawState constants
Private Const DSS_DISABLED = &H20
Private Const DSS_MONO = &H80
Private Const DSS_NORMAL = &H0
Private Const DSS_RIGHT = &H8000
Private Const DSS_UNION = &H10
Private Const DST_BITMAP = &H4
Private Const DST_COMPLEX = &H0
Private Const DST_ICON = &H3
Private Const DST_PREFIXTEXT = &H2
Private Const DST_TEXT = &H1

Private m_hpalHalftone As Long  'Halftone created for default palette use
Public Sub TileBitmapToHDC(ByVal lhDCDest As Long, _
                                    ByVal picSource As Picture, _
                                    ByVal lLeft As Long, _
                                    ByVal lTop As Long, _
                                    ByVal lWidth As Long, _
                                    ByVal lHeight As Long, _
                                    ByVal lDestLeft As Long, _
                                    ByVal lDestTop As Long, _
                                    ByVal lDestWidth As Long, _
                                    ByVal lDestHeight As Long, _
                                    Optional ByVal lhPal As Long)
    '-------------------------------------------------------------------------
    'Purpose:   Draws a Bitmap to an HDC without transparency
    'In:
    '   [lhdcDest]
    '           HDC of the memory device context to paint the picture on
    '   [picSource]
    '           Picture to paint
    '   [lLeft]
    '           X coordinate of the upper left corner of the area that the
    '           picture is to be painted on. (in pixels)
    '   [lTop]
    '           Y coordinate of the upper left corner of the area that the
    '           picture is to be painted on. (in pixels)
    '   [lWidth]
    '           Width of picture area to paint in pixels
    '   [lHeight]
    '           Height of picture area to paint in pixels
    '   [lhPal]
    '           Must be a valid HPALETTE
    '-------------------------------------------------------------------------
    Dim lhdcTemp As Long
    Dim lhPalOld As Long
    Dim hbmOld As Long
    Dim hDCScreen As Long
    Dim X As Long, Y As Long
    Dim w As Long, H As Long
    
hDCScreen = GetDC(0&)
    
    'Validate that a bitmap was passed in
    If picSource.Type <> vbPicTypeBitmap Then Error.Raise giINVALID_PICTURE
    'Create a DC to select bitmap into
    lhdcTemp = CreateCompatibleDC(hDCScreen)
    lhPalOld = SelectPalette(lhdcTemp, lhPal, True)
    RealizePalette lhdcTemp
    'Select bitmap into DC
    hbmOld = SelectObject(lhdcTemp, picSource.handle)
    'Spray out across destination
    For X = lDestLeft To lDestLeft + lDestWidth Step lWidth
     For Y = lDestTop To lDestTop + lDestHeight Step lHeight
      ' Check we're not going over the required area:
      If X + lWidth > (lDestLeft + lDestWidth) Then
       w = (lDestLeft + lDestWidth) - X
      Else
       w = lWidth
      End If
      If Y + lHeight > (lDestTop + lDestHeight) Then
       H = (lDestTop + lDestHeight) - Y
      Else
       H = lHeight
      End If
      BitBlt lhDCDest, X, Y, w, H, lhdcTemp, 0, 0, vbSrcCopy
     Next
    Next
    'Cleanup
    SelectObject lhdcTemp, hbmOld
    SelectPalette lhdcTemp, lhPalOld, True
    RealizePalette (lhdcTemp)
    DeleteDC lhdcTemp
  ReleaseDC 0&, hDCScreen
End Sub

Public Sub PaintDisabledPicture(ByVal hDCDest As Long, _
                                ByVal picSource As StdPicture, _
                                ByVal xDest As Long, _
                                ByVal yDest As Long, _
                                ByVal Width As Long, _
                                ByVal Height As Long, _
                                Optional ByVal xSrc As Long = 0, _
                                Optional ByVal ySrc As Long = 0, _
                                Optional ByVal clrMask As OLE_COLOR = 16711935, _
                                Optional ByVal hPal As Long = 0)
    Dim lbmTemp As Long
    Dim lbmTempOld As Long
    Dim lhdcTemp As Long
    Dim lhPalOld As Long
    Dim udtTempRect As RECT
    Dim lhbrWhite As Long
    Dim hDCScreen As Long
    
    On Error Resume Next
    If picSource Is Nothing Then
     'No Image!
    ElseIf picSource.Type = vbPicTypeNone Then
     'No Image!
    ElseIf picSource.Type = vbPicTypeIcon Then
     'Icon
     DrawState hDCDest, 0, 0, picSource.handle, 0, xDest, yDest, 0, 0, (DST_ICON Or DSS_DISABLED)
     If Err Then PaintDisabledPictureEx hDCDest, xDest, yDest, Width, Height, picSource, xSrc, ySrc, clrMask, vb3DHighlight, vb3DShadow, hPal
    Else
     hDCScreen = GetDC(0&)
     lhdcTemp = CreateCompatibleDC(hDCScreen)     'Create a temporary hDC compatible to the Destination hDC
     lbmTemp = CreateCompatibleBitmap(hDCScreen, Width, Height)
     lbmTempOld = SelectObject(lhdcTemp, lbmTemp)
     lhPalOld = SelectPalette(lhdcTemp, hPal, True)
     RealizePalette lhdcTemp
     With udtTempRect
      .Top = 0
      .Left = 0
      .Bottom = Height
      .Right = Width
     End With
     SetBkColor lhdcTemp, vbWhite
     lhbrWhite = CreateSolidBrush(vbWhite)
     FillRect lhdcTemp, udtTempRect, lhbrWhite
     PaintTransparentPicture lhdcTemp, picSource, 0, 0, Width, Height, xSrc, ySrc, clrMask
     SelectObject lhdcTemp, lbmTempOld
     Err.Clear
     DrawState hDCDest, 0, 0, lbmTemp, 0, xDest, yDest, 0, 0, (DST_BITMAP Or DSS_DISABLED)
     If Err Then PaintDisabledPictureEx hDCDest, xDest, yDest, Width, Height, picSource, xSrc, ySrc, clrMask, vb3DHighlight, vb3DShadow, hPal
     'Cleanup
     DeleteObject lhbrWhite
     SelectPalette lhdcTemp, lhPalOld, True
     RealizePalette lhdcTemp
     DeleteObject SelectObject(lhdcTemp, lbmTempOld)
     DeleteObject lbmTempOld
     DeleteObject lbmTemp
     DeleteObject lhPalOld
     DeleteDC lhdcTemp
  ReleaseDC 0&, hDCScreen
 End If
End Sub


Public Sub PaintCheckedPattern(ByVal hDCDest As Long, ByVal hdcSrc As Long, ByVal xDest As Long, ByVal yDest As Long, _
                                     ByVal wRequired As Long, ByVal hRequired As Long, _
                                    ByVal clrRequired As OLE_COLOR, Optional ByVal hPal As Long = 0)
    Dim hdcMask As Long        'hDC of the created mask image
    Dim hdcColor As Long       'hDC of the created color image
    Dim hbmMask As Long        'Bitmap handle to the mask image
    Dim hbmColor As Long       'Bitmap handle to the color image
    Dim hbmColorOld As Long
    Dim hbmMaskOld As Long
    Dim hPalOld As Long
    Dim hPalBufferOld As Long
    Dim lRequiredColor As Long
    Dim X As Long, Y As Long
    Dim bltWidth As Integer, bltHeight As Integer
    Dim hDCScreen As Long
    
    On Error Resume Next
hDCScreen = GetDC(0&)
    ' Validate palette
    If hPal = 0 Then hPal = m_hpalHalftone
    
    OleTranslateColor clrRequired, hPal, lRequiredColor
    
    'Create a (color) bitmap for the cover (can't use CompatibleBitmap with
    'hDCSrc, because this will create a DIB section if the original bitmap
    'is a DIB section)
    hbmColor = CreateCompatibleBitmap(hDCScreen, 8, 8)
    'Now create a monochrome bitmap for the mask
    hbmMask = CreateBitmap(8, 8, 1, 1, ByVal 0&)
    'First, blt the source bitmap onto the cover.  We do this first
    'and then use it instead of the source bitmap
    'because the source bitmap may be
    'a DIB section, which behaves differently than a bitmap.
    '(Specifically, copying from a DIB section to a monochrome bitmap
    'does a nearest-color selection rather than painting based on the
    'backcolor and forecolor.
    hdcColor = CreateCompatibleDC(hDCScreen)
    hbmColorOld = SelectObject(hdcColor, hbmColor)
    hPalOld = SelectPalette(hdcColor, hPal, True)
    RealizePalette hdcColor
    'In case hDCSrc contains a monochrome bitmap, we must set the destination
    'foreground/background colors according to those currently set in hDCSrc
    '(because Windows will associate these colors with the two monochrome colors)
    SetBkColor hdcColor, GetBkColor(hdcSrc)
    SetTextColor hdcColor, GetTextColor(hdcSrc)
    BitBlt hdcColor, 0, 0, 8, 8, hdcSrc, 0, 0, vbSrcCopy
    'Paint the mask.  What we want is white at the transparent color
    'from the source, and black everywhere else.
    hdcMask = CreateCompatibleDC(hDCScreen)
    hbmMaskOld = SelectObject(hdcMask, hbmMask)

    'When bitblt'ing from color to monochrome, Windows sets to 1
    'all pixels that match the background color of the source DC.  All
    'other bits are set to 0.
    SetBkColor hdcColor, vbBlack
    SetTextColor hdcColor, vbWhite
    BitBlt hdcMask, 0, 0, 8, 8, hdcColor, 0, 0, vbSrcCopy
    'Paint the rest of the cover bitmap.
    '
    'What we want here is black at the transparent color, and
    'the original colors everywhere else.  To do this, we first
    'paint the original onto the cover (which we already did), then we
    'AND the inverse of the mask onto that using the DSna ternary raster
    'operation (0x00220326 - see Win32 SDK reference, Appendix, "Raster
    'Operation Codes", "Ternary Raster Operations", or search in MSDN
    'for 00220326).  DSna [reverse polish] means "(not SRC) and DEST".
    '
    'When bitblt'ing from monochrome to color, Windows transforms all white
    'bits (1) to the background color of the destination hDC.  All black (0)
    'bits are transformed to the foreground color.
    SetTextColor hDCDest, lRequiredColor
    
    ' Tile image:
    For X = xDest To xDest + wRequired Step 8
        For Y = yDest To yDest + hRequired Step 8
            ' Check we're not going over the required area:
            If X + 8 > (xDest + wRequired) Then
                bltWidth = (xDest + wRequired) - X
            Else
                bltWidth = 8
            End If
            If Y + 8 > (yDest + hRequired) Then
                bltHeight = (yDest + hRequired) - Y
            Else
                bltHeight = 8
            End If
            BitBlt hDCDest, X, Y, bltWidth, bltHeight, hdcMask, 0, 0, vbSrcCopy
        Next Y
    Next X
    ' All done!
    DeleteObject SelectObject(hdcColor, hbmColorOld)
    SelectPalette hdcColor, hPalOld, True
    RealizePalette hdcColor
    DeleteDC hdcColor
    DeleteObject SelectObject(hdcMask, hbmMaskOld)
    DeleteDC hdcMask
    ReleaseDC 0&, hDCScreen
End Sub

'-------------------------------------------------------------------------
'Purpose:   Creates a disabled-appearing (grayed) bitmap, given any format of
'           input bitmap.
'In:
'   [hdcDest]
'           Device context to paint the picture on
'   [xDest]
'           X coordinate of the upper left corner of the area that the
'           picture is to be painted on. (in pixels)
'   [yDest]
'           Y coordinate of the upper left corner of the area that the
'           picture is to be painted on. (in pixels)
'   [Width]
'           Width of picture area to paint in pixels.  Note: If this value
'           is outrageous (i.e.: you passed a forms ScaleWidth in twips
'           instead of the pictures' width in pixels), this procedure will
'           attempt to create bitmaps that require outrageous
'           amounts of memory.
'   [Height]
'           Height of picture area to paint in pixels.  Note: If this
'           value is outrageous (i.e.: you passed a forms ScaleHeight in
'           twips instead of the pictures' height in pixels), this
'           procedure will attempt to create bitmaps that require
'           outrageous amounts of memory.
'   [picSource]
'           Standard Picture object to be used as the image source
'   [xSrc]
'           X coordinate of the upper left corner of the area in the picture
'           to use as the source. (in pixels)
'           Ignored if picSource is an Icon.
'   [ySrc]
'           Y coordinate of the upper left corner of the area in the picture
'           to use as the source. (in pixels)
'           Ignored if picSource is an Icon.
'   [clrMask]
'           Color of pixels to be masked out
'   [clrHighlight]
'           Color to be used as outline highlight
'   [clrShadow]
'           Color to be used as outline shadow
'   [hPal]
'           Handle of palette to select into the memory DC's used to create
'           the painting effect.
'           If not provided, a HalfTone palette is used.
'-------------------------------------------------------------------------
Public Sub PaintDisabledPictureEx(ByVal hDCDest As Long, _
                                ByVal xDest As Long, _
                                ByVal yDest As Long, _
                                ByVal Width As Long, _
                                ByVal Height As Long, _
                                ByVal picSource As StdPicture, _
                                Optional ByVal xSrc As Long = 0, _
                                Optional ByVal ySrc As Long = 0, _
                                Optional ByVal clrMask As OLE_COLOR = vbWhite, _
                                Optional ByVal clrHighlight As OLE_COLOR = vb3DHighlight, _
                                Optional ByVal clrShadow As OLE_COLOR = vb3DShadow, _
                                Optional ByVal hPal As Long = 0)
Attribute PaintDisabledPictureEx.VB_Description = "Paints a disabled appearing image (embossed) given a source picture object."
    Dim hdcSrc As Long         'hDC that the source bitmap is selected into
    Dim hbmMemSrcOld As Long
    Dim hbmMemSrc As Long
    Dim udtRect As RECT
    Dim hbrMask As Long
    Dim lMaskColor As Long
    Dim hDCScreen As Long
    Dim hPalOld As Long
    
    'Verify that the passed picture is not nothing
    If picSource Is Nothing Then GoTo PaintDisabledDC_InvalidParam
    Select Case picSource.Type
        Case vbPicTypeBitmap
            'Select passed picture into an hDC
            hDCScreen = GetDC(0&)
            'Validate palette
            If hPal = 0 Then
                hPal = m_hpalHalftone
            End If
            hdcSrc = CreateCompatibleDC(hDCScreen)
            hbmMemSrcOld = SelectObject(hdcSrc, picSource.handle)
            hPalOld = SelectPalette(hdcSrc, hPal, True)
            RealizePalette hdcSrc
            
            'Draw the bitmap
            PaintDisabledDC hDCDest, xDest, yDest, Width, Height, hdcSrc, xSrc, ySrc, clrMask, clrHighlight, clrShadow, hPal
            
            SelectObject hdcSrc, hbmMemSrcOld
            SelectPalette hdcSrc, hPalOld, True
            RealizePalette hdcSrc
            DeleteDC hdcSrc
            ReleaseDC 0&, hDCScreen
        Case vbPicTypeIcon
            'Create a bitmap and select it into a DC
            hDCScreen = GetDC(0&)
            'Validate palette
            If hPal = 0 Then
                hPal = m_hpalHalftone
            End If
            On Error Resume Next
             DrawState hDCDest, 0, 0, picSource.handle, 0, xDest, yDest, 0, 0, (DST_ICON Or DSS_DISABLED)
             If Err Then
              hdcSrc = CreateCompatibleDC(hDCScreen)
              hbmMemSrc = CreateCompatibleBitmap(hDCScreen, Width, Height)
              hbmMemSrcOld = SelectObject(hdcSrc, hbmMemSrc)
              hPalOld = SelectPalette(hdcSrc, hPal, True)
              RealizePalette hdcSrc
              'Draw Icon onto DC
              udtRect.Bottom = Height
              udtRect.Right = Width
              OleTranslateColor clrMask, 0&, lMaskColor
              SetBkColor hdcSrc, lMaskColor
              hbrMask = CreateSolidBrush(lMaskColor)
              FillRect hdcSrc, udtRect, hbrMask
              DeleteObject hbrMask
              DrawIcon hdcSrc, 0, 0, picSource.handle
              'Draw Disabled image
              PaintDisabledDC hDCDest, xDest, yDest, Width, Height, hdcSrc, 0&, 0&, clrMask, clrHighlight, clrShadow, hPal
              'Clean up
              SelectPalette hdcSrc, hPalOld, True
              RealizePalette hdcSrc
              DeleteObject SelectObject(hdcSrc, hbmMemSrcOld)
              DeleteDC hdcSrc
             End If
            On Error GoTo 0
           ReleaseDC 0&, hDCScreen
        Case Else
            GoTo PaintDisabledDC_InvalidParam
    End Select
    Exit Sub
PaintDisabledDC_InvalidParam:
    'Error.Raise giINVALID_PICTURE
    Exit Sub
End Sub

'-------------------------------------------------------------------------
'Purpose:   Creates a disabled-appearing (grayed) bitmap, given any format of
'           input bitmap.
'In:
'   [hdcDest]
'           Device context to paint the picture on
'   [xDest]
'           X coordinate of the upper left corner of the area that the
'           picture is to be painted on. (in pixels)
'   [yDest]
'           Y coordinate of the upper left corner of the area that the
'           picture is to be painted on. (in pixels)
'   [Width]
'           Width of picture area to paint in pixels.  Note: If this value
'           is outrageous (i.e.: you passed a forms ScaleWidth in twips
'           instead of the pictures' width in pixels), this procedure will
'           attempt to create bitmaps that require outrageous
'           amounts of memory.
'   [Height]
'           Height of picture area to paint in pixels.  Note: If this
'           value is outrageous (i.e.: you passed a forms ScaleHeight in
'           twips instead of the pictures' height in pixels), this
'           procedure will attempt to create bitmaps that require
'           outrageous amounts of memory.
'   [hdcSrc]
'           Device context that contains the source picture
'   [xSrc]
'           X coordinate of the upper left corner of the area in the picture
'           to use as the source. (in pixels)
'   [ySrc]
'           Y coordinate of the upper left corner of the area in the picture
'           to use as the source. (in pixels)
'   [clrMask]
'           Color of pixels to be masked out
'   [clrHighlight]
'           Color to be used as outline highlight
'   [clrShadow]
'           Color to be used as outline shadow
'   [hPal]
'           Handle of palette to select into the memory DC's used to create
'           the painting effect.
'           If not provided, a HalfTone palette is used.
'-------------------------------------------------------------------------
Public Sub PaintDisabledDC(ByVal hDCDest As Long, _
                                ByVal xDest As Long, _
                                ByVal yDest As Long, _
                                ByVal Width As Long, _
                                ByVal Height As Long, _
                                ByVal hdcSrc As Long, _
                                Optional ByVal xSrc As Long = 0, _
                                Optional ByVal ySrc As Long = 0, _
                                Optional ByVal clrMask As OLE_COLOR = vbWhite, _
                                Optional ByVal clrHighlight As OLE_COLOR = vb3DHighlight, _
                                Optional ByVal clrShadow As OLE_COLOR = vb3DShadow, _
                                Optional ByVal hPal As Long = 0)
Attribute PaintDisabledDC.VB_Description = "Paints a disabled appearing image (embossed) given a source hDC."
    Dim hDCScreen As Long
    Dim hbmMonoSection As Long
    Dim hbmMonoSectionSav As Long
    Dim hdcMonoSection As Long
    Dim hdcColor As Long
    Dim hdcDisabled As Long
    Dim hbmDisabledSav As Long
    Dim lpbi As BITMAPINFO
    Dim hbmMono As Long
    Dim hdcMono As Long
    Dim hbmMonoSav As Long
    Dim lMaskColor As Long
    Dim lMaskColorCompare As Long
    Dim hdcMaskedSource As Long
    Dim hbmMasked As Long
    Dim hbmMaskedOld As Long
    Dim hpalMaskedOld As Long
    Dim hpalDisabledOld As Long
    Dim hpalMonoOld As Long
    Dim rgbBlack As RGBQUAD
    Dim rgbWhite As RGBQUAD
    Dim dwSys3dShadow As Long
    Dim dwSys3dHighlight As Long
    Dim pvBits As Long
    Dim rgbnew(1) As RGBQUAD
    Dim hbmDisabled As Long
    Dim lMonoBkGrnd As Long
    Dim lMonoBkGrndChoices(2) As Long
    Dim lIndex As Long  'For ... Next index
    Dim hbrWhite As Long
    Dim udtRect As RECT
    
    'TODO: handle pictures with dark masks
    If hPal = 0 Then
        hPal = m_hpalHalftone
    End If
  ' Define some colors
    OleTranslateColor clrShadow, hPal, dwSys3dShadow
    OleTranslateColor clrHighlight, hPal, dwSys3dHighlight
    
    hDCScreen = GetDC(0&)
    With rgbBlack
        .rgbBlue = 0
        .rgbGreen = 0
        .rgbRed = 0
        .rgbReserved = 0
    End With
    With rgbWhite
        .rgbBlue = 255
        .rgbGreen = 255
        .rgbRed = 255
        .rgbReserved = 255
    End With

    ' The first step is to create a monochrome bitmap with two colors:
    ' white where colors in the original are light, and black
    ' where the original is dark.  We can't simply bitblt to a bitmap.
    ' Instead, we create a monochrome (bichrome?) DIB section and bitblt
    ' to that.  Windows will do the conversion automatically based on the
    ' DIB section's palette.  (I.e. using a DIB section, Windows knows how
    ' to map "light" colors and "dark" colors to white/black, respectively.
    With lpbi.bmiHeader
        .biSize = LenB(lpbi.bmiHeader)
        .biWidth = Width
        .biHeight = -Height
        .biPlanes = 1
        .biBitCount = 1         ' monochrome
        .biCompression = BI_RGB
        .biSizeImage = 0
        .biXPelsPerMeter = 0
        .biYPelsPerMeter = 0
        .biClrUsed = 0          ' max colors used (2^1 = 2)
        .biClrImportant = 0     ' all (both :-]) colors are important
    End With
    With lpbi
        .bmiColors(0) = rgbBlack
        .bmiColors(1) = rgbWhite
    End With

    hbmMonoSection = CreateDIBSection(hDCScreen, lpbi, DIB_RGB_COLORS, pvBits, 0&, 0)
    
    hdcMonoSection = CreateCompatibleDC(hDCScreen)
    hbmMonoSectionSav = SelectObject(hdcMonoSection, hbmMonoSection)
    
    'Bitblt to the Monochrome DIB section
    'If a mask color is provided, create a new bitmap and copy the source
    'to it transparently.  If we don't do this, a dark mask color will be
    'turned into the outline part of the monochrome DIB section
    'Convert mask color and white before comparing
    'because the Mask color might be a system color that would be evaluated
    'to white.
    OleTranslateColor vbWhite, hPal, lMaskColorCompare
    OleTranslateColor clrMask, hPal, lMaskColor
    If lMaskColor = lMaskColorCompare Then
        BitBlt hdcMonoSection, 0, 0, Width, Height, hdcSrc, xSrc, ySrc, vbSrcCopy
    Else
        hbmMasked = CreateCompatibleBitmap(hDCScreen, Width, Height)
        hdcMaskedSource = CreateCompatibleDC(hDCScreen)
        hbmMaskedOld = SelectObject(hdcMaskedSource, hbmMasked)
        hpalMaskedOld = SelectPalette(hdcMaskedSource, hPal, True)
        RealizePalette hdcMaskedSource
        'Fill the bitmap with white
        With udtRect
            .Left = 0
            .Top = 0
            .Right = Width
            .Bottom = Height
        End With
        hbrWhite = CreateSolidBrush(vbWhite)
        FillRect hdcMaskedSource, udtRect, hbrWhite
        DeleteObject hbrWhite
        'Do the transparent paint
        PaintTransparentDC hdcMaskedSource, 0, 0, Width, Height, hdcSrc, xSrc, ySrc, lMaskColor, hPal
        'BitBlt to the Mono DIB section.  The mask color has been turned to white.
        BitBlt hdcMonoSection, 0, 0, Width, Height, hdcMaskedSource, 0, 0, vbSrcCopy
        'Clean up
        SelectPalette hdcMaskedSource, hpalMaskedOld, True
        RealizePalette hdcMaskedSource
        DeleteObject SelectObject(hdcMaskedSource, hbmMaskedOld)
        DeleteDC hdcMaskedSource
    End If
      
    ' Okay, we've got our B&W DIB section.
    ' Now that we have our monochrome bitmap, the final appearance that we
    ' want is this:  First, think of the black portion of the monochrome
    ' bitmap as our new version of the original bitmap.  We want to have a dark
    ' gray version of this with a light version underneath it, shifted down and
    ' to the right.  The light acts as a highlight, and it looks like the original
    ' image is a gray inset.
    
    ' First, create a copy of the destination.  Draw the light gray transparently,
    ' and then draw the dark gray transparently
    
    hbmDisabled = CreateCompatibleBitmap(hDCScreen, Width, Height)
    
    hdcDisabled = CreateCompatibleDC(hDCScreen)
    hbmDisabledSav = SelectObject(hdcDisabled, hbmDisabled)
    hpalDisabledOld = SelectPalette(hdcDisabled, hPal, True)
    RealizePalette hdcDisabled
    'We used to fill the background with gray, instead copy the
    'destination to memory DC.  This will allow a disabled image
    'to be drawn over a background image.
    BitBlt hdcDisabled, 0, 0, Width, Height, hDCDest, xDest, yDest, vbSrcCopy
    
    'When painting the monochrome bitmaps transparently onto the background
    'we need a background color that is not the light color of the dark color
    'Provide three choices to ensure a unique color is picked.
    OleTranslateColor vbBlack, hPal, lMonoBkGrndChoices(0)
    OleTranslateColor vbRed, hPal, lMonoBkGrndChoices(1)
    OleTranslateColor vbBlue, hPal, lMonoBkGrndChoices(2)
    
    'Pick a background color choice that doesn't match
    'the shadow or highlight color
    For lIndex = 0 To 2
        If lMonoBkGrndChoices(lIndex) <> dwSys3dHighlight And _
                lMonoBkGrndChoices(lIndex) <> dwSys3dShadow Then
            'This color can be used for a mask
            lMonoBkGrnd = lMonoBkGrndChoices(lIndex)
            Exit For
        End If
    Next

    ' Now paint a the light color shifted and transparent over the background
    ' It is not necessary to change the DIB section's color table
    ' to equal the highlight color and mask color.  In fact, setting
    ' the color table to anything besides black and white causes unpredictable
    ' results (seen in win95 with IE4, using 256 colors).
    ' Setting the Back and Text colors of the Monochrome bitmap, ensure
    ' that the desired colors are produced.
    With rgbnew(0)
        .rgbRed = (vbWhite \ 2 ^ 16) And &HFF
        .rgbGreen = (vbWhite \ 2 ^ 8) And &HFF
        .rgbBlue = vbWhite And &HFF
    End With
    With rgbnew(1)
        .rgbRed = (vbBlack \ 2 ^ 16) And &HFF
        .rgbGreen = (vbBlack \ 2 ^ 8) And &HFF
        .rgbBlue = vbBlack And &HFF
    End With
        
    SetDIBColorTable hdcMonoSection, 0, 2, rgbnew(0)
    
    '...We can't pass a DIBSection to PaintTransparentDC(), so we need to
    ' make a copy of our mono DIBSection.  Notice that we only need a monochrome
    ' bitmap, but we must set its back/fore colors to the monochrome colors we
    ' want (light gray and black), and PaintTransparentDC() will honor them.
    hbmMono = CreateBitmap(Width, Height, 1, 1, ByVal 0&)
    hdcMono = CreateCompatibleDC(hDCScreen)
    hbmMonoSav = SelectObject(hdcMono, hbmMono)
    SetMapMode hdcMono, GetMapMode(hdcSrc)
    SetBkColor hdcMono, dwSys3dHighlight
    SetTextColor hdcMono, lMonoBkGrnd
    hpalMonoOld = SelectPalette(hdcMono, hPal, True)
    RealizePalette hdcMono
    BitBlt hdcMono, 0, 0, Width, Height, hdcMonoSection, 0, 0, vbSrcCopy

    '...We can go ahead and call PaintTransparentDC with our monochrome
    ' copy
    ' Draw this transparently over the disabled bitmap
    '...Don't forget to shift right and left....
    PaintTransparentDC hdcDisabled, 1, 1, Width, Height, hdcMono, 0, 0, lMonoBkGrnd, hPal
    
    ' Now draw a transparent copy, using dark gray where the monochrome had
    ' black, and transparent elsewhere.  We'll use a transparent color of black.

    '...We can't pass a DIBSection to PaintTransparentDC(), so we need to
    ' make a copy of our mono DIBSection.  Notice that we only need a monochrome
    ' bitmap, but we must set its back/fore colors to the monochrome colors we
    ' want (dark gray and black), and PaintTransparentDC() will honor them.
    ' Use hbmMono and hdcMono; already created for first color
    SetBkColor hdcMono, dwSys3dShadow
    SetTextColor hdcMono, lMonoBkGrnd
    BitBlt hdcMono, 0, 0, Width, Height, hdcMonoSection, 0, 0, vbSrcCopy

    '...We can go ahead and call PaintTransparentDC with our monochrome
    ' copy
    ' Draw this transparently over the disabled bitmap
    PaintTransparentDC hdcDisabled, 0, 0, Width, Height, hdcMono, 0, 0, lMonoBkGrnd, hPal
    BitBlt hDCDest, xDest, yDest, Width, Height, hdcDisabled, 0, 0, vbSrcCopy
    ' Okay, we're done!
    SelectPalette hdcDisabled, hpalDisabledOld, True
    RealizePalette hdcDisabled
    DeleteObject SelectObject(hdcMonoSection, hbmMonoSectionSav)
    DeleteDC hdcMonoSection
    DeleteObject SelectObject(hdcDisabled, hbmDisabledSav)
    DeleteDC hdcDisabled
    DeleteObject SelectObject(hdcMono, hbmMonoSav)
    SelectPalette hdcMono, hpalMonoOld, True
    RealizePalette hdcMono
    DeleteDC hdcMono
    ReleaseDC 0&, hDCScreen
End Sub

'-------------------------------------------------------------------------
'Purpose:   Draws a transparent bitmap to a DC.  The pixels of the passed
'           bitmap that match the passed mask color will not be painted
'           to the destination DC
'In:
'   [hdcDest]
'           Device context to paint the picture on
'   [xDest]
'           X coordinate of the upper left corner of the area that the
'           picture is to be painted on. (in pixels)
'   [yDest]
'           Y coordinate of the upper left corner of the area that the
'           picture is to be painted on. (in pixels)
'   [Width]
'           Width of picture area to paint in pixels.  Note: If this value
'           is outrageous (i.e.: you passed a forms ScaleWidth in twips
'           instead of the pictures' width in pixels), this procedure will
'           attempt to create bitmaps that require outrageous
'           amounts of memory.
'   [Height]
'           Height of picture area to paint in pixels.  Note: If this
'           value is outrageous (i.e.: you passed a forms ScaleHeight in
'           twips instead of the pictures' height in pixels), this
'           procedure will attempt to create bitmaps that require
'           outrageous amounts of memory.
'   [hdcSrc]
'           Device context that contains the source picture
'   [xSrc]
'           X coordinate of the upper left corner of the area in the picture
'           to use as the source. (in pixels)
'   [ySrc]
'           Y coordinate of the upper left corner of the area in the picture
'           to use as the source. (in pixels)
'   [clrMask]
'           Color of pixels to be masked out
'   [hPal]
'           Handle of palette to select into the memory DC's used to create
'           the painting effect.
'           If not provided, a HalfTone palette is used.
'-------------------------------------------------------------------------
Public Sub PaintTransparentDC(ByVal hDCDest As Long, _
                                    ByVal xDest As Long, _
                                    ByVal yDest As Long, _
                                    ByVal Width As Long, _
                                    ByVal Height As Long, _
                                    ByVal hdcSrc As Long, _
                                    Optional ByVal xSrc As Long = 0, _
                                    Optional ByVal ySrc As Long = 0, _
                                    Optional ByVal clrMask As OLE_COLOR = 16711935, _
                                    Optional ByVal hPal As Long = 0)
Attribute PaintTransparentDC.VB_Description = "Paints an image with transparent pixels defined by the mask color.  Accepts an hDC as its image source."
    Dim hdcMask As Long        'hDC of the created mask image
    Dim hdcColor As Long       'hDC of the created color image
    Dim hbmMask As Long        'Bitmap handle to the mask image
    Dim hbmColor As Long       'Bitmap handle to the color image
    Dim hbmColorOld As Long
    Dim hbmMaskOld As Long
    Dim hPalOld As Long
    Dim hDCScreen As Long
    Dim hdcScnBuffer As Long         'Buffer to do all work on
    Dim hbmScnBuffer As Long
    Dim hbmScnBufferOld As Long
    Dim hPalBufferOld As Long
    Dim lMaskColor As Long
    
    hDCScreen = GetDC(0&)
    'Validate palette
    If hPal = 0 Then
        hPal = m_hpalHalftone
    End If
    OleTranslateColor clrMask, hPal, lMaskColor
    
    'Create a color bitmap to server as a copy of the destination
    'Do all work on this bitmap and then copy it back over the destination
    'when it's done.
    hbmScnBuffer = CreateCompatibleBitmap(hDCScreen, Width, Height)
    'Create DC for screen buffer
    hdcScnBuffer = CreateCompatibleDC(hDCScreen)
    hbmScnBufferOld = SelectObject(hdcScnBuffer, hbmScnBuffer)
    hPalBufferOld = SelectPalette(hdcScnBuffer, hPal, True)
    RealizePalette hdcScnBuffer
    'Copy the destination to the screen buffer
    BitBlt hdcScnBuffer, 0, 0, Width, Height, hDCDest, xDest, yDest, vbSrcCopy
    
    'Create a (color) bitmap for the cover (can't use CompatibleBitmap with
    'hdcSrc, because this will create a DIB section if the original bitmap
    'is a DIB section)
    hbmColor = CreateCompatibleBitmap(hDCScreen, Width, Height)
    'Now create a monochrome bitmap for the mask
    hbmMask = CreateBitmap(Width, Height, 1, 1, ByVal 0&)
    'First, blt the source bitmap onto the cover.  We do this first
    'and then use it instead of the source bitmap
    'because the source bitmap may be
    'a DIB section, which behaves differently than a bitmap.
    '(Specifically, copying from a DIB section to a monochrome bitmap
    'does a nearest-color selection rather than painting based on the
    'backcolor and forecolor.
    hdcColor = CreateCompatibleDC(hDCScreen)
    hbmColorOld = SelectObject(hdcColor, hbmColor)
    hPalOld = SelectPalette(hdcColor, hPal, True)
    RealizePalette hdcColor
    'In case hdcSrc contains a monochrome bitmap, we must set the destination
    'foreground/background colors according to those currently set in hdcSrc
    '(because Windows will associate these colors with the two monochrome colors)
    SetBkColor hdcColor, GetBkColor(hdcSrc)
    SetTextColor hdcColor, GetTextColor(hdcSrc)
    BitBlt hdcColor, 0, 0, Width, Height, hdcSrc, xSrc, ySrc, vbSrcCopy
    'Paint the mask.  What we want is white at the transparent color
    'from the source, and black everywhere else.
    hdcMask = CreateCompatibleDC(hDCScreen)
    hbmMaskOld = SelectObject(hdcMask, hbmMask)

    'When bitblt'ing from color to monochrome, Windows sets to 1
    'all pixels that match the background color of the source DC.  All
    'other bits are set to 0.
    SetBkColor hdcColor, lMaskColor
    SetTextColor hdcColor, vbWhite
    BitBlt hdcMask, 0, 0, Width, Height, hdcColor, 0, 0, vbSrcCopy
    'Paint the rest of the cover bitmap.
    '
    'What we want here is black at the transparent color, and
    'the original colors everywhere else.  To do this, we first
    'paint the original onto the cover (which we already did), then we
    'AND the inverse of the mask onto that using the DSna ternary raster
    'operation (0x00220326 - see Win32 SDK reference, Appendix, "Raster
    'Operation Codes", "Ternary Raster Operations", or search in MSDN
    'for 00220326).  DSna [reverse polish] means "(not SRC) and DEST".
    '
    'When bitblt'ing from monochrome to color, Windows transforms all white
    'bits (1) to the background color of the destination hDC.  All black (0)
    'bits are transformed to the foreground color.
    SetTextColor hdcColor, vbBlack
    SetBkColor hdcColor, vbWhite
    BitBlt hdcColor, 0, 0, Width, Height, hdcMask, 0, 0, DSna
    'Paint the Mask to the Screen buffer
    BitBlt hdcScnBuffer, 0, 0, Width, Height, hdcMask, 0, 0, vbSrcAnd
    'Paint the Color to the Screen buffer
    BitBlt hdcScnBuffer, 0, 0, Width, Height, hdcColor, 0, 0, vbSrcPaint
    'Copy the screen buffer to the screen
    BitBlt hDCDest, xDest, yDest, Width, Height, hdcScnBuffer, 0, 0, vbSrcCopy
    'All done!
    DeleteObject SelectObject(hdcColor, hbmColorOld)
    SelectPalette hdcColor, hPalOld, True
    RealizePalette hdcColor
    DeleteDC hdcColor
    DeleteObject SelectObject(hdcScnBuffer, hbmScnBufferOld)
    SelectPalette hdcScnBuffer, hPalBufferOld, True
    RealizePalette hdcScnBuffer
    DeleteDC hdcScnBuffer
    
    DeleteObject SelectObject(hdcMask, hbmMaskOld)
    DeleteDC hdcMask
    ReleaseDC 0&, hDCScreen
End Sub

'-------------------------------------------------------------------------
'Purpose:   Draws a transparent bitmap to a DC.  The pixels of the passed
'           bitmap that match the passed mask color will not be painted
'           to the destination DC
'In:
'   [hdcDest]
'           Device context to paint the picture on
'   [xDest]
'           X coordinate of the upper left corner of the area that the
'           picture is to be painted on. (in pixels)
'   [yDest]
'           Y coordinate of the upper left corner of the area that the
'           picture is to be painted on. (in pixels)
'   [Width]
'           Width of picture area to paint in pixels.  Note: If this value
'           is outrageous (i.e.: you passed a forms ScaleWidth in twips
'           instead of the pictures' width in pixels), this procedure will
'           attempt to create bitmaps that require outrageous
'           amounts of memory.
'   [Height]
'           Height of picture area to paint in pixels.  Note: If this
'           value is outrageous (i.e.: you passed a forms ScaleHeight in
'           twips instead of the pictures' height in pixels), this
'           procedure will attempt to create bitmaps that require
'           outrageous amounts of memory.
'   [picSource]
'           Standard Picture object to be used as the image source
'   [xSrc]
'           X coordinate of the upper left corner of the area in the picture
'           to use as the source. (in pixels)
'           Ignored if picSource is an Icon.
'   [ySrc]
'           Y coordinate of the upper left corner of the area in the picture
'           to use as the source. (in pixels)
'           Ignored if picSource is an Icon.
'   [clrMask]
'           Color of pixels to be masked out
'   [hPal]
'           Handle of palette to select into the memory DC's used to create
'           the painting effect.
'           If not provided, a HalfTone palette is used.
'-------------------------------------------------------------------------
Public Sub PaintTransparentPicture(ByVal hDCDest As Long, _
                                    ByVal picSource As Picture, _
                                    ByVal xDest As Long, _
                                    ByVal yDest As Long, _
                                    ByVal Width As Long, _
                                    ByVal Height As Long, _
                                    Optional ByVal xSrc As Long = 0, _
                                    Optional ByVal ySrc As Long = 0, _
                                    Optional ByVal clrMask As OLE_COLOR = 16711935, _
                                    Optional ByVal hPal As Long = 0)
Attribute PaintTransparentPicture.VB_Description = "Paints an image with transparent pixels defined by the mask color.  Accepts a picture object as its image source."
    Dim hdcSrc As Long         'hDC that the source bitmap is selected into
    Dim hbmMemSrcOld As Long
    Dim hbmMemSrc As Long
    Dim udtRect As RECT
    Dim hbrMask As Long
    Dim lMaskColor As Long
    Dim hDCScreen As Long
    Dim hPalOld As Long
    'Verify that the passed picture is a Bitmap
    If picSource Is Nothing Then GoTo PaintTransparentStdPic_InvalidParam
    
    Select Case picSource.Type
        Case vbPicTypeBitmap
            hDCScreen = GetDC(0&)
            'Validate palette
            If hPal = 0 Then
                hPal = m_hpalHalftone
            End If
            'Select passed picture into an hDC
            hdcSrc = CreateCompatibleDC(hDCScreen)
            hbmMemSrcOld = SelectObject(hdcSrc, picSource.handle)
            hPalOld = SelectPalette(hdcSrc, hPal, True)
            RealizePalette hdcSrc
            'Draw the bitmap
            PaintTransparentDC hDCDest, xDest, yDest, Width, Height, hdcSrc, xSrc, ySrc, clrMask, hPal
            
            SelectObject hdcSrc, hbmMemSrcOld
            SelectPalette hdcSrc, hPalOld, True
            RealizePalette hdcSrc
            DeleteDC hdcSrc
            ReleaseDC 0&, hDCScreen
        Case vbPicTypeIcon
            'Create a bitmap and select it into an DC
            hDCScreen = GetDC(0&)
            'Validate palette
            If hPal = 0 Then
                hPal = m_hpalHalftone
            End If
            hdcSrc = CreateCompatibleDC(hDCScreen)
            hbmMemSrc = CreateCompatibleBitmap(hDCScreen, Width, Height)
            hbmMemSrcOld = SelectObject(hdcSrc, hbmMemSrc)
            hPalOld = SelectPalette(hdcSrc, hPal, True)
            RealizePalette hdcSrc
            'Draw Icon onto DC
            udtRect.Bottom = Height
            udtRect.Right = Width
            OleTranslateColor clrMask, 0&, lMaskColor
            hbrMask = CreateSolidBrush(lMaskColor)
            FillRect hdcSrc, udtRect, hbrMask
            DeleteObject hbrMask
            DrawIcon hdcSrc, 0, 0, picSource.handle
            'Draw Transparent image
            PaintTransparentDC hDCDest, xDest, yDest, Width, Height, hdcSrc, xSrc, ySrc, lMaskColor, hPal
            'Clean up
            DeleteObject SelectObject(hdcSrc, hbmMemSrcOld)
            SelectPalette hdcSrc, hPalOld, True
            RealizePalette hdcSrc
            DeleteDC hdcSrc
            ReleaseDC 0&, hDCScreen
        Case Else
            GoTo PaintTransparentStdPic_InvalidParam
    End Select
    Exit Sub
PaintTransparentStdPic_InvalidParam:
'    Err.Raise giINVALID_PICTURE
    Exit Sub
End Sub

'-------------------------------------------------------------------------
'Purpose:   Draws a standard picture object to a DC
'In:
'   [hdcDest]
'           Handle of the device context to paint the picture on
'   [xDest]
'           X coordinate of the upper left corner of the area that the
'           picture is to be painted on. (in pixels)
'   [yDest]
'           Y coordinate of the upper left corner of the area that the
'           picture is to be painted on. (in pixels)
'   [Width]
'           Width of picture area to paint in pixels.  Note: If this value
'           is outrageous (i.e.: you passed a forms ScaleWidth in twips
'           instead of the pictures' width in pixels), this procedure will
'           attempt to create bitmaps that require outrageous
'           amounts of memory.
'   [Height]
'           Height of picture area to paint in pixels.  Note: If this
'           value is outrageous (i.e.: you passed a forms ScaleHeight in
'           twips instead of the pictures' height in pixels), this
'           procedure will attempt to create bitmaps that require
'           outrageous amounts of memory.
'   [picSource]
'           Standard Picture object to be used as the image source
'   [xSrc]
'           X coordinate of the upper left corner of the area in the picture
'           to use as the source. (in pixels)
'           Ignored if picSource is an Icon.
'   [ySrc]
'           Y coordinate of the upper left corner of the area in the picture
'           to use as the source. (in pixels)
'           Ignored if picSource is an Icon.
'   [hPal]
'           Handle of palette to select into the memory DC's used to create
'           the painting effect.
'           If not provided, a HalfTone palette is used.
'-------------------------------------------------------------------------
Public Sub PaintStandardPicture(ByVal hDCDest As Long, _
                                    ByVal picSource As Picture, _
                                    ByVal xDest As Long, _
                                    ByVal yDest As Long, _
                                    ByVal Width As Long, _
                                    ByVal Height As Long, _
                                    Optional ByVal xSrc As Long = 0, _
                                    Optional ByVal ySrc As Long = 0, _
                                    Optional ByVal hPal As Long = 0)
Attribute PaintStandardPicture.VB_Description = "Paints an image provided by a picture object to an hDC with no effects."
    Dim hdcTemp As Long
    Dim hPalOld As Long
    Dim hbmMemSrcOld As Long
    Dim hDCScreen As Long
    Dim hbmMemSrc As Long
    'Validate that a bitmap was passed in
    If picSource Is Nothing Then GoTo PaintNormalStdPic_InvalidParam
    Select Case picSource.Type
        Case vbPicTypeBitmap
            If hPal = 0 Then
                hPal = m_hpalHalftone
            End If
            hDCScreen = GetDC(0&)
            'Create a DC to select bitmap into
            hdcTemp = CreateCompatibleDC(hDCScreen)
            hPalOld = SelectPalette(hdcTemp, hPal, True)
            RealizePalette hdcTemp
            'Select bitmap into DC
            hbmMemSrcOld = SelectObject(hdcTemp, picSource.handle)
            'Copy to destination DC
            BitBlt hDCDest, xDest, yDest, Width, Height, hdcTemp, xSrc, ySrc, vbSrcCopy
            'Cleanup
            SelectObject hdcTemp, hbmMemSrcOld
            SelectPalette hdcTemp, hPalOld, True
            RealizePalette hdcTemp
            DeleteDC hdcTemp
            ReleaseDC 0&, hDCScreen
        Case vbPicTypeIcon
            'Create a bitmap and select it into an DC
            'Draw Icon onto DC
            DrawIconEx hDCDest, xDest, yDest, picSource.handle, Width, Height, 0&, 0&, DI_NORMAL
        Case Else
            GoTo PaintNormalStdPic_InvalidParam
    End Select
    Exit Sub
PaintNormalStdPic_InvalidParam:
    'Err.Raise giINVALID_PICTURE
End Sub

Private Sub Class_Initialize()
    Dim hDCScreen As Long
    'Create halftone palette
    hDCScreen = GetDC(0&)
    m_hpalHalftone = CreateHalftonePalette(hDCScreen)
    ReleaseDC 0&, hDCScreen
End Sub

Private Sub Class_Terminate()
    DeleteObject m_hpalHalftone
End Sub
