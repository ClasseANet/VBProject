VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.UserControl HtmlEditor 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   2490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2490
   Palette         =   "HtmlEditor.ctx":0000
   ScaleHeight     =   2490
   ScaleWidth      =   2490
   ToolboxBitmap   =   "HtmlEditor.ctx":0342
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   1275
      Left            =   240
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   960
      Width           =   1935
      ExtentX         =   3413
      ExtentY         =   2249
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.PictureBox PicVerSizer 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   180
      ScaleHeight     =   255
      ScaleWidth      =   855
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   855
      Begin VB.Line LineVerSizer 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         X1              =   120
         X2              =   720
         Y1              =   120
         Y2              =   120
      End
   End
   Begin VB.PictureBox PicHorSizer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   1740
      ScaleHeight     =   735
      ScaleWidth      =   255
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   60
      Visible         =   0   'False
      Width           =   255
      Begin VB.Line LineHorSizer 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         X1              =   120
         X2              =   120
         Y1              =   0
         Y2              =   720
      End
   End
End
Attribute VB_Name = "HtmlEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Private Type GUID   'UUID
     Data1 As Long
     Data2 As Long
     Data3 As Long
     Data4(0 To 7) As Byte
End Type

Private Declare Function CLSIDFromString Lib "ole32" (lpsz As Byte, pGuid As GUID) As Long
Private Declare Function StringFromGUID2 Lib "ole32" (pGuid As GUID, ByVal szGuid As String, ByVal cchMax As Long) As Long
Private Declare Function IsEqualGUID Lib "ole32" (pGuid1 As Any, pGuid2 As Any) As Long
'====================================================================
Private WithEvents m_oWebBrowser As SHDocVw.WebBrowser               ' WebBrowser control
Attribute m_oWebBrowser.VB_VarHelpID = -1
Private WithEvents m_oDocument As HTMLDocument
Attribute m_oDocument.VB_VarHelpID = -1

Private Const strGUID_HTML = "{DE4BA900-59CA-11CF-9592-444553540000}" 'CGIDSTR_HTML
Dim GUID_HTML As GUID
Dim CommandTarget As IOleCommandTarget       ' WebBrowser's IOleCommandTarget interface
'====================================================================
'           Events
'====================================================================
Event StatusTextChange(ByVal Text As String)
Event UpdatePageStatus(ByVal pDisp As Object, nPage As Variant, fDone As Variant)
Event NavigateComplete2(ByVal pDisp As Object, URL As Variant)
Event CommandStateChange(ByVal Command As Long, ByVal Enable As Boolean)
Event DocumentMouseMove(Element As IHTMLElement, oEvent As IHTMLEventObj)
Event DocumentMouseDown(Element As IHTMLElement, oEvent As IHTMLEventObj)
Event DocumentMouseUp(Element As IHTMLElement, oEvent As IHTMLEventObj)
Event DocumentMouseOver(Element As IHTMLElement, oEvent As IHTMLEventObj)
Event DocumentMouseOut(Element As IHTMLElement, oEvent As IHTMLEventObj)
Event DocumentOnClick(Element As IHTMLElement, oEvent As IHTMLEventObj)
'====================================================================
'====================================================================
Private m_bDesignMode As Boolean                          ' BrowseMode
Private m_bReadyState As Boolean
Private m_bLiveResize As Boolean
Private m_bMultipleSelection As Boolean
Private m_bShowBorders  As Boolean
Private m_bShowDetails As Boolean
Private m_bVisible As Boolean
Private m_sContent As String
Private m_bProtectMetaTags As Boolean
Private m_bSelectTables As Boolean

Private BookmarkIndex As String
Private sStr As String

Private Const PropName_ShowGrid = "ShowGrid"
Private Const PropName_SnapToGrid = "SnapToGrid"
Private Const PropName_GridSize = "GridSize"
Private Const PropName_GridBrush = "GridBrush"
Private Const PropName_GridBrushBMP = "GridBrushBMP"

Private SelectedCell As IHTMLElement
Private m_bSizingCursor As Boolean
Private m_bSizingCell As Boolean
Private m_bSizingCellWidth As Boolean
Private m_bSizingCellHeight As Boolean
Private m_lSizingCellStartX As Long
Private m_lSizingCellStartY As Long
Private m_lLastHorSizeX As Long
Private m_lLastVerSizeY As Long

Private TableSelectedEvent As IHTMLEventObj
Private TableSelectedElement As IHTMLElement
Private SelectedHTMLElemenRECT As RECT

Private m_bTableSelectStarted As Boolean
Private m_TableSelectedCells As Collection 'mshtml.HTMLElementCollection

Public Function m_oDocument_onclick() As Boolean

    Dim e As IHTMLEventObj
    Dim el As IHTMLElement
    
    Set e = m_oDocument.parentWindow.event
    Set el = e.srcElement ' e.toElement
    RaiseEvent DocumentOnClick(el, e)
   
    'Set el = m_oDocument.elementFromPoint(e.X, e.Y)
    'Debug.Print "El: "; el.tagName
    
    'Debug.Print "OnClick: "; e.clientX; e.clientY
End Function

Public Sub m_oDocument_onmouseover()
    Dim e As IHTMLEventObj
    Dim el As IHTMLElement
    
    If m_oDocument.ReadyState <> "complete" Then Exit Sub
    Set e = m_oDocument.parentWindow.event
    Set el = e.srcElement ' e.toElement
    RaiseEvent DocumentMouseOver(el, e)
End Sub

Public Sub m_oDocument_onmouseout()
    Dim e As IHTMLEventObj
    Dim el As IHTMLElement
    
    If m_oDocument.ReadyState <> "complete" Then Exit Sub
    Set e = m_oDocument.parentWindow.event
    Set el = e.srcElement ' e.toElement
    RaiseEvent DocumentMouseOut(el, e)
End Sub

Public Sub GetActualElemenRECT()
    
    Dim aEvent As IHTMLEventObj
    Dim aElement As IHTMLElement2
    Dim P As PointAPI
    Dim aElement2 As IHTMLElement2

    '> //We handle the click ourselves - cancel the browsers click handling
    
    Set aEvent = m_oDocument.parentWindow.event
    'IHTMLWindow2 as    'IHTMLEventObj;
    '> aevent.returnValue := True;
    
    ' //Calculate component's screen pos (upper left corner)
    P.X = aEvent.screenX - aEvent.offsetX
    P.y = aEvent.screenY - aEvent.offsetY
    '> //Calculate object rect on main window
    SelectedHTMLElemenRECT.Left = P.X
    SelectedHTMLElemenRECT.Top = P.y
    
    Set aElement2 = aEvent.srcElement ' as IHTMLElement2;
    SelectedHTMLElemenRECT.Right = P.X + aElement2.clientWidth
    SelectedHTMLElemenRECT.Bottom = P.y + aElement2.clientHeight
    'Debug.Print "Rect: "; P.X, P.Y; SelectedHTMLElemenRECT.Right; SelectedHTMLElemenRECT.Bottom

End Sub

Public Sub m_oDocument_onmouseup()

    Dim e As IHTMLEventObj
    Dim el As IHTMLElement
    Dim elW As Long
    Dim elH As Long
    Dim elTop As Long
    Dim elLeft As Long
    Dim Table As IHTMLTable
    Dim CurX As Long
    Dim CurY As Long
    Dim td As IHTMLTableCell
    Dim tr As IHTMLTableRow
    Dim TR2 As IHTMLTableRow2
    Dim Cell As IHTMLTableCell
    Dim NewHeight As Long, NewWidth As Long
    Dim CellCollection As IHTMLElementCollection
    Dim RowCollection As IHTMLElementCollection
    
    If m_oDocument.ReadyState <> "complete" Then Exit Sub
    
    Set e = m_oDocument.parentWindow.event
    Set el = e.srcElement
    
    Dim el2 As IHTMLElement2
    Set el2 = e.srcElement
        
    RaiseEvent DocumentMouseUp(el, e)
    
    'Debug.Print "m_oDocument_onmouseup"
    
    'If m_bSizingCell And e.Button = vbLeftButton Then
    If m_bSizingCell And (Not TableSelectedElement Is Nothing) Then
        m_bSizingCell = False
        Screen.MousePointer = vbDefault 'vbIbeam
        PicHorSizer.Visible = False
        PicVerSizer.Visible = False
        
        elW = el.offsetWidth
        elH = el.offsetHeight
        elTop = e.screenY - e.offsetY
        elLeft = e.screenX - e.offsetX
        
        'If el.tagName = "TD" And e.Button = vbLeftButton Then
        'If el.tagName = "TD" And e.Button = vbLeftButton Then
            'Set td = e.srcElement
            '--------------------------------------------------------
            'resize cell width
            If m_bSizingCellWidth Then
                NewWidth = e.screenX - SelectedHTMLElemenRECT.Left
                If Not TableSelectedElement Is Nothing Then
                    If NewWidth > 1 Then
                        'TableSelectedElement
                        
                        'td.Width = e.offsetX
                        ' If SnapToGrid Then
                        ' I = NewWidth Mod SnapToGridX
                        ' NewWidth = (NewWidth \ SnapToGridX) * SnapToGridX
                        ' If I > (SnapToGridX \ 2) Then
                        '    NewWidth = NewWidth + (SnapToGridX \ 2)
                        ' End If
                        
                        Dim cellIndex As Integer
                        Set Table = TableSelectedElement.offsetParent ' parentElement
                        Set td = TableSelectedElement
                        cellIndex = td.cellIndex
                        Set TR2 = TableSelectedElement.parentElement
                        For Each tr In Table.Rows
                            Set CellCollection = tr.cells
                            Set td = CellCollection.Item(vbNull, cellIndex)
                            td.Width = NewWidth
                        Next tr
                            'TR2.Height = NewHeight
                            'For Each cell In tr.cells
                                'cell.Height = e.offsetY
                            'Next cell
                    End If
                End If
            '--------------------------------------------------------
            'resize row height
            ElseIf m_bSizingCellHeight Then
                NewHeight = e.screenY - SelectedHTMLElemenRECT.Top
                'TableSelectedEvent
                'TableSelectedElement

'                 If SnapToGrid Then
'                 I = NewHeight Mod SnapToGridY
'                 NewHeight = (NewHeight \ SnapToGridY) * SnapToGridY
'                 If I > (SnapToGridY \ 2) Then
'                    NewHeight = NewHeight + (SnapToGridY \ 2)
'                End If
                
                If NewHeight > 1 Then
                    If Not TableSelectedElement Is Nothing Then
                        Set TR2 = TableSelectedElement.parentElement
                        TR2.Height = NewHeight
                        'For Each cell In tr.cells
                            'cell.Height = e.offsetY
                        'Next cell
                    End If
                End If
                'If ((IsValid) And (SnapToGrid) And (GridSize > 2)) Then
                'DragRc.Left = Round((DragRc.Left / GridSize), 0) * GridSize - 1
            '--------------------------------------------------------
            End If
        'End If
    End If
'                       SnapToGrid
'http://groups.yahoo.com/group/delphi-dhtmledit/message/1305
'> if SnapToGrid
'> then begin
'> I := NewWidth mod SnapToGridX;
'> NewWidth := (NewWidth div SnapToGridX) * SnapToGridX;
'> if I > (SnapToGridX div 2)
'> then NewWidth := NewWidth + (SnapToGridX div 2);
'> end;
'>
    '----------------------------------------------------------------
    'Table selection end
    '----------------------------------------------------------------
    If m_bTableSelectStarted Then
        m_bTableSelectStarted = False
        'Debug.Print "End Table select"
    End If
    '----------------------------------------------------------------
End Sub

Public Sub m_oDocument_onmousedown()
    
    Dim e As IHTMLEventObj
    Dim el As IHTMLElement
    Dim elW As Long
    Dim elH As Long
    Dim elTop As Long
    Dim elLeft As Long
    Dim tbl As IHTMLTable
    Dim CurX As Long
    Dim CurY As Long
    
    If m_oDocument.ReadyState <> "complete" Then Exit Sub
    Set e = m_oDocument.parentWindow.event
    Set el = e.srcElement
    
    Dim el2 As IHTMLElement2
    Set el2 = e.srcElement
        
    'Debug.Print "Element: " & el.tagName; e.X; e.y

    RaiseEvent DocumentMouseDown(el, e)
    
    elW = el.offsetWidth
    elH = el.offsetHeight
    elTop = e.screenY - e.offsetY
    elLeft = e.screenX - e.offsetX
    
    If el.tagName = "TD" And e.Button = vbLeftButton Then
        m_lSizingCellStartX = e.screenX
        m_lSizingCellStartY = e.screenY
        'If Not m_bSizingCell Then
        'ClearSelection
        If e.screenX > (elLeft + elW - 6) And e.screenX < (elLeft + elW) Then
            m_bSizingCell = True
            m_bSizingCellWidth = True
            m_bSizingCellHeight = False
            SetCursorSizeWE
            m_bSizingCursor = True
            PicHorSizer.Visible = True
            PicHorSizer.ZOrder
            'keep the postition and the element of the table select start
            Set TableSelectedEvent = m_oDocument.parentWindow.event
            Set TableSelectedElement = e.srcElement
            GetActualElemenRECT
        ElseIf e.screenY > (elTop + elH - 6) And e.screenY < (elTop + elH) Then
            m_bSizingCell = True
            m_bSizingCellHeight = True
            m_bSizingCellWidth = False
            SetCursorSizeNS
            m_bSizingCursor = True
            PicVerSizer.Visible = True
            PicVerSizer.ZOrder
            'keep the postition and the element of the table select start
            Set TableSelectedEvent = m_oDocument.parentWindow.event
            Set TableSelectedElement = e.srcElement
            GetActualElemenRECT
        End If
        'End If
    End If
    '----------------------------------------------------------------
    ' Table selection start
    '----------------------------------------------------------------
    If el.tagName = "TD" And e.Button = vbLeftButton Then
        If Not m_bTableSelectStarted Then
            m_bTableSelectStarted = True
            Set m_TableSelectedCells = Nothing
            'Debug.Print "Start Table select"
            'm_TableSelectedCells.Add el, el.tagName
        End If
        'ClearSelection
    End If
    '----------------------------------------------------------------
End Sub

Public Sub m_oDocument_onmousemove()
    
    Dim e As IHTMLEventObj
    Dim el As IHTMLElement
    Dim td As IHTMLTableCell
    Dim elW As Long
    Dim elH As Long
    Dim elTop As Long
    Dim elLeft As Long
    Dim tbl As IHTMLTable
    Dim CurX As Long
    Dim CurY As Long
    
    If m_oDocument.ReadyState <> "complete" Then Exit Sub
    Set e = m_oDocument.parentWindow.event
    Set el = e.srcElement
    
    Dim el2 As IHTMLElement2
    Set el2 = e.srcElement
        
    RaiseEvent DocumentMouseMove(el, e)
    
    elW = el.offsetWidth
    elH = el.offsetHeight
    elTop = e.screenY - e.offsetY
    elLeft = e.screenX - e.offsetX
    
    '----------------------------------------------------------------
    'Display sizing cursor near cells borders
    If el.tagName = "TD" And e.Button = 0 Then
        'If Not m_bSizingCell Then
        'ClearSelection
        If (e.screenX > (elLeft + elW - 6) And e.screenX < (elLeft + elW)) Then
            SetCursorSizeWE
            m_bSizingCursor = True
        ElseIf (e.screenY > (elTop + elH - 6) And e.screenY < (elTop + elH)) Then
            SetCursorSizeNS
            m_bSizingCursor = True
        End If
        'End If
    End If
    '----------------------------------------------------------------
    
    'm_bSizingCellWidth = True
    'm_bSizingCellHeight = False
    
    'If m_bSizingCell And el.tagName = "TD" Then
    If m_bSizingCell Then
        ClearSelection
        'Set td = e.srcElement
        'Set SelectedCell = el
        'Set tbl = el.offsetParent
        'Debug.Print "L:"; elLeft; "T:"; elTop; "W:"; elW; "H:"; elH; "X:"; e.X; "Y:"; e.Y; "offsetX"; e.offsetX; "offsetY"; e.offsetY
        'Debug.Print "El: "; e.screenX - e.offsetX; e.screenY - e.offsetY
        'Debug.Print "tbl "; tbl.Width
        'If m_bSizingCursor = False Then
            'm_bSizingCellHeight = True
        
        'If (e.screenX > (elLeft + elW - 6) And e.screenX < (elLeft + elW)) Then
'        If m_bSizingCellWidth Then
'            SetCursorSizeWE
'            m_bSizingCursor = True
'        'ElseIf (e.screenY > (elTop + elH - 6) And e.screenY < (elTop + elH)) Then
'        ElseIf m_bSizingCellHeight Then
'            SetCursorSizeNS
'            m_bSizingCursor = True
'        End If
        
        If m_bSizingCellWidth Then
            SetCursorSizeWE
            DrawHorSizingLine e.screenX
            m_bSizingCursor = True
        ElseIf m_bSizingCellHeight Then
            'td.Height = e.offsetY
            SetCursorSizeNS
            DrawVerSizingLine e.screenY
            m_bSizingCursor = True
        End If
    End If
    
    '----------------------------------------------------------------
    'Table selection moving
    If m_bTableSelectStarted Then
        ClearSelection
        HighlightSelectedCell
    End If
    '----------------------------------------------------------------
End Sub

Public Sub HighlightSelectedCell()
    
    Dim e As IHTMLEventObj
    Dim el As IHTMLElement
    Dim td As IHTMLTableCell
    Dim Style As IHTMLStyle
    
    If Not m_bSelectTables Then Exit Sub
    
    Set e = m_oDocument.parentWindow.event
    'Set el = e.srcElement ' e.toElement
    Set el = m_oDocument.elementFromPoint(e.X, e.y)
    
    'Set el = GetElementUnderCaret
    If el Is Nothing Then Exit Sub
    If Not (TypeOf el Is MSHTML.IHTMLTableCell) Then Exit Sub
    
    Set td = el
    Set Style = td.Style
    Style.BorderColor = vbRed
    Style.BorderWidth = 2
    Style.backgroundColor = RGB(0, 0, 255)
    'Debug.Print "HighlightSelectedCell: "; td.cellIndex
    'td.bgColor = CStr(Val(td.bgColor) Xor &HCECECE)
    
    
End Sub

Public Sub DrawHorSizingLine(X As Long)

    Dim hPen As Long, hOldPen As Long
    Dim DC As Long
    Dim P1 As PointAPI
    Dim P2 As PointAPI
    Dim LP As PointAPI
    
    If m_lLastHorSizeX = X Then Exit Sub
    m_lLastHorSizeX = X
    
    P1.X = X
    P1.y = 0
    'P2.X = X2
    'P2.Y = Y2
    
     'ClientToScreen UserControl.hwnd, P1
    ScreenToClient UserControl.hWnd, P1
    'ScreenToClient UserControl.hwnd, P2
    'DC = GetDC(UserControl.hwnd)
    'hPen = CreatePen(PS_SOLID, 1, cColor)
    'hOldPen = SelectObject(DC, hPen)
    'SetROP2 DC, R2_NOTXORPEN
    'MoveToEx DC, P1.X, P1.Y, LP
    'LineTo DC, P2.X, P2.Y
    'Rectangle DC, P1.X, P1.Y, P2.X + 1, P2.Y
    'SelectObject DC, hOldPen
    'DeleteObject hPen
    'ReleaseDC UserControl.hwnd, DC
    
    PicHorSizer.Move (P1.X + 2) * Screen.TwipsPerPixelX, 1, 1, UserControl.ScaleHeight
    PicHorSizer.Refresh
    
End Sub

Public Sub PicHorSizer_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    m_oDocument_onmousedown
End Sub

Public Sub PicVerSizer_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    m_oDocument_onmousedown
End Sub

Public Sub PicHorSizer_Resize()
    LineHorSizer.X1 = 0
    LineHorSizer.Y1 = 0
    LineHorSizer.X2 = 0
    LineHorSizer.Y2 = UserControl.ScaleHeight
    LineHorSizer.Refresh
End Sub

Public Sub DrawVerSizingLine(y As Long)

    Dim P1 As PointAPI
    Dim P2 As PointAPI
    
    If m_lLastVerSizeY = y Then Exit Sub
    m_lLastVerSizeY = y
    
    P1.X = 0
    P1.y = y
    
    ScreenToClient UserControl.hWnd, P1
    PicVerSizer.Move 1, (P1.y + 2) * Screen.TwipsPerPixelY, UserControl.ScaleWidth, 1
    PicVerSizer.Refresh
    
End Sub

Public Sub PicVerSizer_Resize()
    LineVerSizer.X1 = 0
    LineVerSizer.Y1 = 0
    LineVerSizer.X2 = UserControl.ScaleWidth
    LineVerSizer.Y2 = 0
    LineVerSizer.Refresh
End Sub

Sub ColorToRGBByte(ByVal Colr As Long, ByRef R As Byte, ByRef G As Byte, ByRef B As Byte)
    R = Colr Mod 256
    G = Colr \ 256 Mod 256
    B = Colr \ 65536
End Sub

'====================================================================
'execCommand
'queryCommandEnabled
'queryCommandIndeterm
'queryCommandSupported
'queryCommandValue
'====================================================================
Sub ColorToRGB(ByVal Colr As Long, ByRef R As Long, ByRef G As Long, ByRef B As Long)
    R = Colr Mod 256
    G = Colr \ 256 Mod 256
    B = Colr \ 65536
End Sub

Public Sub SetCursorSizeWE()
    Dim hIcon As Long
    'hIcon = LoadCursor(ByVal 0&, IDC_SIZENWSE)
    'hIcon = LoadCursor(ByVal 0&, IDC_SIZENS)
    'hIcon = LoadCursor(ByVal 0&, IDC_SIZENESW)
    hIcon = LoadCursor(ByVal 0&, IDC_SIZEWE)
    Call SetCursor(hIcon)
End Sub

Public Sub SetCursorSizeNS()
    Dim hIcon As Long
    hIcon = LoadCursor(ByVal 0&, IDC_SIZENS)
    Call SetCursor(hIcon)
End Sub

'Returns the element directly under the insertion point
Public Function GetElementUnderCaret() As IHTMLElement
    
    On Error Resume Next
    
    Dim rg As IHTMLTxtRange
    Dim ctlRg As IHTMLControlRange

    ' Branch on the type of selection and
    ' get the element under the caret or the site selected object
    ' and return it

    Select Case m_oDocument.selection.Type

    Case "None", "Text"

        Set rg = m_oDocument.selection.createRange

        ' Collapse the range so that the scope of the
        ' range of the selection is the the caret. That way the
        ' parentElement method will return the element directly
        ' under the caret. If you don't want to change the state of the
        ' selection, then duplicate the range and collapse it

        If Not rg Is Nothing Then
            rg.collapse
            Set GetElementUnderCaret = rg.parentElement
        End If

    Case "Control"
        ' An element is site selected
        Set ctlRg = m_oDocument.selection.createRange
 
        ' There can only be one site selected element at a time so the
        ' commonParentElement will return the site selected element
        Set GetElementUnderCaret = ctlRg.commonParentElement
        
    End Select

End Function

'====================================================================
Public Function CommandStatus(ByVal CMDID As WBIDM, Optional Name As String) As Long
    
    Dim uOLECMD As OLECMD
    Dim uCMDTEXT As OLECMDTEXT
    'Dim oCommandTarget As IOleCommandTarget      ' WebBrowser's IOleCommandTarget interface
    
    'Set oCommandTarget = EditorDoc
    '----------------------------------------------------------------
    ' Initialize the UDTs
    
    uOLECMD.CMDID = CMDID ' A command identifier; taken from the OLECMDID enumeration
    uOLECMD.cmdf = 0 ' should be set to zero
    
    'uOLECMD.cmdf = OLECMDF_ENABLED 'Flags associated with cmdID; taken from the OLECMDF enumeration
    'OLECMDF:     OLECMDF_SUPPORTED    = 1,     OLECMDF_ENABLED      = 2,
    'OLECMDF_LATCHED      = 4,    OLECMDF_NINCHED = 8
    
    uCMDTEXT.cmdtextf = OLECMDTEXTF_NAME
    uCMDTEXT.cwBuf = 260
    '----------------------------------------------------------------
    ' Query the status
    CommandTarget.QueryStatus GUID_HTML, 1, uOLECMD, uCMDTEXT
    '----------------------------------------------------------------
    ' Return the status
    CommandStatus = uOLECMD.cmdf
    
    ' Return the name
    Name = uCMDTEXT.rgwz
    Name = Left$(Name, InStr(Name, vbNullChar) - 1)
    'Debug.Print "cmdtextf: " & uCMDTEXT.cmdtextf
    'Debug.Print "cmdtextf: " & uCMDTEXT.cwBuf
End Function

Public Sub CommandExec(ByVal CMDID As WBIDM, Optional ByVal CMDOPT As OLECMDEXECOPT = OLECMDEXECOPT_DODEFAULT, Optional VarIn As Variant, Optional varOut As Variant)
    
    'Dim oCommandTarget As olelib.IOleCommandTarget      ' WebBrowser's IOleCommandTarget interface
    'Set oCommandTarget = EditorDoc
    ' Execute the command
    CommandTarget.Exec GUID_HTML, CMDID, CMDOPT, VarIn, varOut

End Sub

Public Function CmdExec(ByVal CMDID As String, Optional ByVal ShowUI As Boolean = False, Optional ByVal vValue As Variant) As Boolean
    CmdExec = m_oDocument.execCommand(CMDID, ShowUI, vValue)
End Function

'Returns a Boolean value that indicates the current state of the command.
'true The given command has been executed on the object.
'false The given command has not been executed on the object.
Public Function CmdState(ByVal CMDID As String) As Boolean
    CmdState = m_oDocument.queryCommandState(ByVal CMDID)
End Function

Public Function CmdValue(ByVal CMDID As String) As Variant
    CmdValue = m_oDocument.queryCommandValue(CMDID)
End Function

Public Function CmdEnabled(ByVal CMDID As String) As Boolean
    CmdEnabled = m_oDocument.queryCommandEnabled(CMDID)
End Function

Public Function CmdSupported(ByVal CMDID As String) As Boolean
    CmdSupported = m_oDocument.queryCommandSupported(CMDID)
End Function

Public Function ReadyState() As Boolean
    ReadyState = m_bReadyState
End Function

'====================================================================
Public Sub UserControl_Initialize()
    
    m_bReadyState = False
    m_bSizingCursor = False
    
    WebBrowser1.Navigate2 "about:blank"
    
    'WebBrowser1.Navigate2 "http://mewsoft.com"
    Do While WebBrowser1.ReadyState <> READYSTATE_COMPLETE: DoEvents: Loop
    Do While Not m_bReadyState: DoEvents: Loop
    
    Set m_oWebBrowser = WebBrowser1
    
    Set m_oDocument = WebBrowser1.Document

    Set CommandTarget = WebBrowser1.Document ' IOleCommandTarget
    '----------------------------------------------------------------
    '----------------------------------------------------------------
    'DE4BA900-59CA-11CF-9592-444553540000
    Dim suLib() As Byte
    suLib = strGUID_HTML & vbNullChar
    CLSIDFromString suLib(0), GUID_HTML
    '----------------------------------------------------------------
    m_oDocument.body.innerHTML = m_sContent
    
    Do While m_oDocument.ReadyState <> "complete": DoEvents: Loop
    
    SetDirty False
    '----------------------------------------------------------------
    'CommandExec IDM_PROTECTMETATAGS, OLECMDEXECOPT_DODEFAULT, True, vbNull
    ProtectMetaTags True
    '----------------------------------------------------------------
    BookmarkIndex = 0
    '----------------------------------------------------------------
    'WebBrowser handle
    m_lWebBrowserhWnd = GetWebBrowserHandle
    
    'If Not HookedWebBrowser Then Call HookWebBrowser
    'Call RefreshContainer(m_lWebBrowserhWnd)
    '----------------------------------------------------------------
  
End Sub

Public Sub UserControl_InitProperties()
    
    'm_bReadyState = False
    'm_oWebBrowser.Navigate2 "about:" & Ambient.DisplayName
    
    'WebBrowser1.Navigate2 "about:blank"
    'Do While WebBrowser1.readyState <> READYSTATE_COMPLETE: DoEvents: Loop
    'Set m_oDocument = WebBrowser1.Document
    'Do While m_oDocument.ReadyState <> "complete": DoEvents: Loop
    
    'm_oDocument.body.innerHTML = m_sContent
    Do While m_oDocument.ReadyState <> "complete": DoEvents: Loop

End Sub

'====================================================================
Public Sub UserControl_ReadProperties(PropBag As PropertyBag)

    ' Read properties
    m_bDesignMode = PropBag.ReadProperty("DesignMode", True)
    m_bLiveResize = PropBag.ReadProperty("LiveResize", True)
    m_bMultipleSelection = PropBag.ReadProperty("MultipleSelection", True)
    m_bShowBorders = PropBag.ReadProperty("ShowBorders", True)
    m_bShowDetails = PropBag.ReadProperty("ShowDetails", False)
    m_bVisible = PropBag.ReadProperty("Visible", True)
    m_bSelectTables = PropBag.ReadProperty("SelectTables", False)
    'm_sContent = PropBag.ReadProperty("Content", "Html Document Content")
    'Me.Visible = m_bVisible
    Me.DesignMode = m_bDesignMode
    Me.LiveResize = m_bLiveResize
    Me.MultipleSelection = m_bMultipleSelection
    Me.ShowBorders = m_bShowBorders
    Me.ShowDetails = m_bShowDetails
    
End Sub

Public Sub UserControl_Terminate()
    
    UnHookWebBrowser
    
End Sub

'====================================================================
Public Sub UserControl_WriteProperties(PropBag As PropertyBag)

    PropBag.WriteProperty "DesignMode", m_bDesignMode, True
    PropBag.WriteProperty "LiveResize", m_bLiveResize, True
    PropBag.WriteProperty "MultipleSelection", m_bMultipleSelection, True
    PropBag.WriteProperty "ShowBorders", m_bShowBorders, True
    PropBag.WriteProperty "ShowDetails", m_bShowDetails, False
    PropBag.WriteProperty "Visible", m_bShowDetails, False
    PropBag.WriteProperty "Visible", m_bVisible, True
    PropBag.WriteProperty "SelectTables", m_bSelectTables, False
    
End Sub

'====================================================================
'====================================================================
'====================================================================
'Private WithEvents m_oWebBrowser As SHDocVw.WebBrowser               ' WebBrowser control
'====================================================================
'====================================================================

Public Sub UserControl_Resize()
    
    WebBrowser1.Move 0, 0, UserControl.Width, UserControl.Height
    
End Sub

Public Sub WebBrowser1_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
     
     m_bReadyState = False
     'Only hook the window procedure once
    'If Not HookedWebBrowser Then Call HookWebBrowser
    'Debug.Print "UserControl.hwnd: "; UserControl.hwnd
     
End Sub

Public Sub WebBrowser1_CommandStateChange(ByVal Command As Long, ByVal Enable As Boolean)
    RaiseEvent CommandStateChange(Command, Enable)
End Sub

Public Sub WebBrowser1_DownloadBegin()
    m_bReadyState = False
End Sub

Public Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    
    m_bReadyState = True
    RaiseEvent NavigateComplete2(pDisp, URL)
    
End Sub

'====================================================================
Public Sub WebBrowser1_StatusTextChange(ByVal Text As String)

    'Set m_oWebBrowser = m_oWebBrowser.document
    RaiseEvent StatusTextChange(Text)
   ' Get the URL from the pointer
   'sURL = SysAllocString(pchURLIn)
End Sub
'====================================================================
Public Sub WebBrowser1_UpdatePageStatus(ByVal pDisp As Object, nPage As Variant, fDone As Variant)
    RaiseEvent UpdatePageStatus(pDisp, nPage, fDone)
End Sub

'====================================================================
Public Function WebBrowser() As Variant
    Set WebBrowser = m_oWebBrowser
End Function


'====================================================================
Public Property Get DesignMode() As Boolean
   DesignMode = m_bDesignMode
End Property

Public Property Let DesignMode(New_DesignMode As Boolean)

   m_bDesignMode = New_DesignMode

    If m_bDesignMode Then
        m_oDocument.DesignMode = "On"
    Else
        m_oDocument.DesignMode = "Off"
    End If
    
    PropertyChanged "DesignMode"
    
End Property

Public Property Get LiveResize() As Boolean
    LiveResize = m_bLiveResize
End Property

Public Property Let LiveResize(ByVal bNewValue As Boolean)
    
    m_bLiveResize = bNewValue
    CmdExec "LiveResize", False, m_bLiveResize
    PropertyChanged "LiveResize"
    
End Property

Public Property Get MultipleSelection() As Boolean
    MultipleSelection = m_bMultipleSelection
End Property

Public Property Let MultipleSelection(ByVal bNewValue As Boolean)
    
    m_bMultipleSelection = bNewValue
    CmdExec "MultipleSelection", False, m_bMultipleSelection
    PropertyChanged "MultipleSelection"
    
End Property

Public Property Get ShowBorders() As Boolean
        
    ShowBorders = m_bShowBorders

End Property

Public Property Let ShowBorders(ByVal bNewValue As Boolean)
    
    m_bShowBorders = bNewValue
    CommandExec IDM_SHOWZEROBORDERATDESIGNTIME, OLECMDEXECOPT_DONTPROMPTUSER, m_bShowBorders
    PropertyChanged "ShowBorders"
    
End Property

Public Property Get ShowDetails() As Boolean
    ShowDetails = m_bShowDetails
End Property

Public Property Let ShowDetails(ByVal bNewValue As Boolean)
    m_bShowDetails = bNewValue
    'CommandExec IDM_SHOWALLTAGS, OLECMDEXECOPT_DONTPROMPTUSER, m_bShowDetails
    ShowAllGlyph bNewValue
    PropertyChanged "ShowDetails"
End Property

Public Function Document() As Variant
    
    Set Document = WebBrowser1.Document
    
End Function

Public Property Get DocumentHtml() As String

    Do While m_oDocument.ReadyState <> "complete": DoEvents: Loop
    DocumentHtml = m_oDocument.body.innerHTML

End Property

Public Property Let DocumentHtml(ByVal sDocumentHtml As String)

    Do While m_oDocument.ReadyState <> "complete": DoEvents: Loop
    m_oDocument.body.innerHTML = sDocumentHtml

End Property

'====================================================================
'====================================================================
'2D-Position
'Allows absolutely positioned elements to be moved by dragging.
Public Sub Position2D()
Attribute Position2D.VB_Description = "Allows absolutely positioned elements to be moved by dragging."
    Dim State As Boolean
    State = IsPosition2D
    State = Not State
    CmdExec "2D-Position", False, State
End Sub

Public Function IsPosition2D() As Boolean
    IsPosition2D = CmdState("2D-Position")
End Function

'AbsolutePosition
'Sets an element's position property to "absolute."
Public Sub AbsolutePosition()
Attribute AbsolutePosition.VB_Description = "Sets an element's position property to ""absolute."""
    Dim State As Boolean
    State = IsAbsolutePosition
    State = Not State
    CmdExec "AbsolutePosition", False, State
End Sub

Public Function IsAbsolutePosition() As Boolean
    IsAbsolutePosition = CmdState("AbsolutePosition")
End Function

Public Function IsAbsolutePositionEnabled() As Boolean
    IsAbsolutePositionEnabled = CmdEnabled("AbsolutePosition")
End Function

'AbsolutePosition + 2D-Position
Public Sub AbsolutePosition2D()
    Call Position2D
    Call AbsolutePosition
End Sub

'BackColor
'Sets or retrieves the background color of the current selection.
Public Sub SetBackColor(Color As String)
Attribute SetBackColor.VB_Description = "Sets or retrieves the background color of the current selection."
     
    CmdExec "BackColor", False, Color

End Sub

Public Function GetBackColor() As String
    
    BackColor = CmdValue("BackColor")

End Function

'BlockDirLTR
'Not currently supported.
Public Sub BlockDirLTR()
    
    CmdExec "BlockDirLTR"

End Sub

'BlockDirRTL
'Not currently supported.
Public Sub BlockDirRTL()
    
    CmdExec "BlockDirRTL"

End Sub

'Bold
'Toggles the current selection between bold and nonbold.
Public Sub Bold()
    CmdExec "Bold"
End Sub

Public Function IsBold() As Boolean
    IsBold = CmdState("Bold")
End Function


'ClearAuthenticationCache
'Clears all authentication credentials from the cache. Applies only to execCommand.
Public Sub ClearAuthenticationCache()
    
    CmdExec "ClearAuthenticationCache"

End Sub

'Copy
'Copies the current selection to the clipboard.
Public Sub Copy()
    CmdExec "Copy"
End Sub

Public Function IsCopy() As Boolean
    IsCopy = CmdEnabled("Copy")
End Function

'CreateBookmark
'Creates a bookmark anchor or retrieves the name of a bookmark anchor for the current selection or insertion point.
Public Sub CreateBookmark(Optional sName As String)
    'Providing an empty string will cause the command to fail.
    BookmarkIndex = BookmarkIndex + 1
    
    If sName = "" Then
        sStr = "Bookmark" & CStr(BookmarkIndex)
        Else
        sStr = sName
    End If
    
    CmdExec "CreateBookmark", False, sStr

End Sub

'CreateLink
'Inserts a hyperlink on the current selection, or displays a dialog box enabling the user to specify a URL to insert as a hyperlink on the current selection.
Public Sub CreateLink()
    CmdExec "CreateLink"
End Sub

Public Function IsCreateLink() As Boolean
    IsCreateLink = CmdEnabled("CreateLink")
End Function

'Cut
'Copies the current selection to the clipboard and then deletes it.
Public Sub Cut()
    CmdExec "Cut"
End Sub

Public Function IsCut() As Boolean
    IsCut = CmdEnabled("Cut")
End Function

'Delete
'Deletes the current selection.
Public Sub Delete()
    CmdExec "Delete"
End Sub

Public Function IsDelete() As Boolean
    IsDelete = CmdEnabled("Delete")
End Function

'DirLTR
'Not currently supported.
Public Sub DirLTR()
    CmdExec "DirLTR"
End Sub

Public Function IsDirLTR() As Boolean
    IsDirLTR = CmdState("DirLTR")
End Function

'DirRTL
'Not currently supported.
Public Sub DirRTL()
    CmdExec "DirRTL"
End Sub

Public Function IsDirRTL() As Boolean
    IsDirRTL = CmdState("DirRTL")
End Function

'FontName
'Sets or retrieves the font for the current selection.
Public Sub SetFontName(sFontName As String)
    CmdExec "FontName", False, sFontName
End Sub

Public Function GetFontName() As String
    GetFontName = CmdValue("FontName")
End Function

'FontSize
'Sets or retrieves the font size for the current selection.
Public Sub SetFontSize(sFontSize As String)
    CmdExec "FontSize", False, sFontSize
End Sub

Public Function GetFontSize() As String
    GetFontSize = CmdValue("FontSize")
End Function

'ForeColor
'Sets or retrieves the foreground (text) color of the current selection.
Public Sub SetForeColor(sColor As String)
'String that specifies a color name or a six-digit hexadecimal RGB value,
'with or without a leading hash mark, as defined in the Color Table.
    CmdExec "ForeColor", False, sColor
End Sub

Public Function GetForeColor() As String
    GetForeColor = CmdValue("ForeColor")
End Function

'FormatBlock
'Sets the current block format tag.
Public Sub FormatBlock(sFormat As String)
    CmdExec "FormatBlock", False, sFormat
End Sub

Public Function GetFormatBlock() As String
    GetFormatBlock = CmdValue("FormatBlock")
End Function

'indent
'Increases the indent of the selected text by one indentation increment.
Public Sub Indent()
    CmdExec "Indent"
End Sub

Public Function IsIndent() As Boolean
    IsIndent = CmdState("Indent")
End Function


'InlineDirLTR
'Not currently supported.
Public Sub InlineDirLTR()
    CmdExec "InlineDirLTR"
End Sub

'InlineDirRTL
'Not currently supported.
Public Sub InlineDirRTL()
    CmdExec "InlineDirRTL"
End Sub

'====================================================================
'           Form functions
'====================================================================
Sub InsertForm()
    
    Dim strHTML As String
    
    FormIndex = FormIndex + 1
    
    strHTML = "<form name=""from" & CStr(FormIndex) & """" & " method=""POST"" action=""http://"">" & vbCrLf & _
            "   <br><br>" & vbCrLf & _
            "   <input type=""submit"" value=""Submit"" name=""B1"">&nbsp;&nbsp;" & vbCrLf & _
            "   <input type=""reset"" value=""Reset"" name=""B2"">" & vbCrLf & _
            "</form>"
    
    InsertHTMLCode strHTML
    
End Sub

'InsertButton
'Overwrites a button control on the text selection.
Public Sub InsertButton()
    CmdExec "InsertButton"
End Sub

'InsertFieldset
'Overwrites a box on the text selection.
Public Sub InsertFieldset()
    CmdExec "InsertFieldset"
End Sub

'InsertHorizontalRule
'Overwrites a horizontal line on the text selection.
Public Sub InsertHorizontalRule()
    CmdExec "InsertHorizontalRule"
End Sub

'InsertIFrame
'Overwrites an inline frame on the text selection.
Public Sub InsertIFrame()
    CmdExec "InsertIFrame"
End Sub

'InsertImage
'Overwrites an image on the text selection.
Public Sub InsertImage()
    CmdExec "InsertImage"
End Sub

Public Function IsInsertImage() As Boolean
    IsInsertImage = CmdEnabled("InsertImage")
End Function

'InsertInputButton
'Overwrites a button control on the text selection.
Public Sub InsertInputButton()
    CmdExec "InsertInputButton"
End Sub

Public Function IsInsertInputButton() As Boolean
    IsInsertInputButton = CmdEnabled("InsertInputButton")
End Function

'InsertInputCheckbox
'Overwrites a check box control on the text selection.
Public Sub InsertInputCheckbox()
    CmdExec "InsertInputCheckbox"
End Sub

Public Function IsInsertInputCheckbox() As Boolean
    IsInsertInputCheckbox = CmdEnabled("InsertInputCheckbox")
End Function

'InsertInputFileUpload
'Overwrites a file upload control on the text selection.
Public Sub InsertInputFileUpload()
    CmdExec "InsertInputFileUpload"
End Sub

Public Function IsInsertInputFileUpload() As Boolean
    IsInsertInputFileUpload = CmdEnabled("InsertInputFileUpload")
End Function

'InsertInputHidden
'Inserts a hidden control on the text selection.
Public Sub InsertInputHidden()
    CmdExec "InsertInputHidden"
End Sub

Public Function IsInsertInputHidden() As Boolean
    IsInsertInputHidden = CmdEnabled("InsertInputHidden")
End Function

'InsertInputImage
'Overwrites an image control on the text selection.
Public Sub InsertInputImage()
    CmdExec "InsertInputImage"
End Sub

Public Function IsInsertInputImage() As Boolean
    IsInsertInputImage = CmdEnabled("InsertInputImage")
End Function

'InsertInputPassword
'Overwrites a password control on the text selection.
Public Sub InsertInputPassword()
    CmdExec "InsertInputPassword"
End Sub

Public Function IsInsertInputPassword() As Boolean
    IsInsertInputPassword = CmdEnabled("InsertInputPassword")
End Function

'InsertInputRadio
'Overwrites a radio control on the text selection.
Public Sub InsertInputRadio()
    CmdExec "InsertInputRadio"
End Sub

Public Function IsInsertInputRadio() As Boolean
    IsInsertInputRadio = CmdEnabled("InsertInputRadio")
End Function

'InsertInputReset
'Overwrites a reset control on the text selection.
Public Sub InsertInputReset()
    CmdExec "InsertInputReset"
End Sub

Public Function IsInsertInputReset() As Boolean
    IsInsertInputReset = CmdEnabled("InsertInputReset")
End Function

'InsertInputSubmit
'Overwrites a submit control on the text selection.
Public Sub InsertInputSubmit()
    CmdExec "InsertInputSubmit"
End Sub

Public Function IsInsertInputSubmit() As Boolean
    IsInsertInputSubmit = CmdEnabled("InsertInputSubmit")
End Function

'InsertInputText
'Overwrites a text control on the text selection.
Public Sub InsertInputText()
    CmdExec "InsertInputText"
End Sub

Public Function IsInsertInputText() As Boolean
    IsInsertInputText = CmdEnabled("InsertInputText")
End Function

'InsertMarquee
'Overwrites an empty marquee on the text selection.
Public Sub InsertMarquee()
    CmdExec "InsertMarquee"
End Sub

Public Function IsInsertMarquee() As Boolean
    IsInsertMarquee = CmdEnabled("InsertMarquee")
End Function

'InsertOrderedList
'Toggles the text selection between an ordered list and a normal format block.
Public Sub InsertOrderedList()
    CmdExec "InsertOrderedList"
End Sub

Public Function IsInsertOrderedList() As Boolean
    IsInsertOrderedList = CmdEnabled("InsertOrderedList")
End Function

'InsertParagraph
'Overwrites a line break on the text selection.
Public Sub InsertParagraph()
    CmdExec "InsertParagraph"
End Sub

Public Function IsInsertParagraph() As Boolean
    IsInsertParagraph = CmdEnabled("InsertParagraph")
End Function

'InsertSelectDropdown
'Overwrites a drop-down selection control on the text selection.
Public Sub InsertSelectDropdown()
    CmdExec "InsertSelectDropdown"
End Sub

Public Function IsInsertSelectDropdown() As Boolean
    IsInsertSelectDropdown = CmdEnabled("InsertSelectDropdown")
End Function

'InsertSelectListbox
'Overwrites a list box selection control on the text selection.
Public Sub InsertSelectListbox()
    CmdExec "InsertSelectListbox"
End Sub

Public Function IsInsertSelectListbox() As Boolean
    IsInsertSelectListbox = CmdEnabled("InsertSelectListbox")
End Function

'InsertTextArea
'Overwrites a multiline text input control on the text selection.
Public Sub InsertTextArea()
    CmdExec "InsertTextArea"
End Sub

Public Function IsInsertTextArea() As Boolean
    IsInsertTextArea = CmdEnabled("InsertTextArea")
End Function

'InsertUnorderedList
'Toggles the text selection between an ordered list and a normal format block.
Public Sub InsertUnorderedList()
    CmdExec "InsertUnorderedList"
End Sub

Public Function IsInsertUnorderedList() As Boolean
    IsInsertUnorderedList = CmdEnabled("InsertUnorderedList")
End Function

'Italic
'Toggles the current selection between italic and nonitalic.
Public Sub Italic()
    CmdExec "Italic"
End Sub

Public Function IsItalic() As Boolean
    IsItalic = CmdState("Italic")
End Function

'JustifyCenter
'Centers the format block in which the current selection is located.
Public Sub JustifyCenter()
    CmdExec "JustifyCenter"
End Sub

Public Function IsJustifyCenter() As Boolean
    IsJustifyCenter = CmdState("JustifyCenter")
End Function

'JustifyFull
'Not currently supported.
Public Sub JustifyFull()
    CmdExec "JustifyFull"
End Sub

Public Function IsJustifyFull() As Boolean
    IsJustifyFull = CmdState("JustifyFull")
End Function

'JustifyLeft
'Left-justifies the format block in which the current selection is located.
Public Sub JustifyLeft()
    CmdExec "JustifyLeft"
End Sub

Public Function IsJustifyLeft() As Boolean
    IsJustifyLeft = CmdState("JustifyLeft")
End Function

'JustifyNone
'Not currently supported.
Public Sub JustifyNone()
    CmdExec "JustifyNone"
End Sub

Public Function IsJustifyNone() As Boolean
    IsJustifyNone = CmdState("JustifyNone")
End Function

'JustifyRight
'Right-justifies the format block in which the current selection is located.
Public Sub JustifyRight()
    CmdExec "JustifyRight"
End Sub

Public Function IsJustifyRight() As Boolean
    IsJustifyRight = CmdState("JustifyRight")
End Function

'Open
'Not currently supported.

'Outdent
'Decreases by one increment the indentation of the format block in which the current selection is located.
Public Sub Outdent()
    CmdExec "Outdent"
End Sub

Public Function IsOutdent() As Boolean
    IsOutdent = CmdState("Outdent")
End Function

'OverWrite
'Toggles the text-entry mode between insert and overwrite.
Public Sub OverWrite()
    CmdExec "OverWrite"
End Sub

Public Function IsOverWrite() As Boolean
    IsOverWrite = CmdState("OverWrite")
End Function

'Paste
'Overwrites the contents of the clipboard on the current selection.
Public Sub Paste()
    CmdExec "Paste"
End Sub

Public Function IsPaste() As Boolean
    IsPaste = CmdEnabled("Paste")
End Function

'PlayImage
'Not currently supported.
Public Sub PlayImage()
    CmdExec "PlayImage"
End Sub

Public Function IsPlayImage() As Boolean
    IsPlayImage = CmdState("PlayImage")
End Function

'Print
'Opens the print dialog box so the user can print the current page.
Public Sub PrintDocument()
    CmdExec "Print"
End Sub

'Redo
'Not currently supported.
Public Sub Redo()
    CmdExec "Redo"
End Sub

Public Function IsRedo() As Boolean
    IsRedo = CmdEnabled("Redo")
End Function

'Refresh
'Refreshes the current document.
Public Sub Refresh()
    CmdExec "Refresh"
End Sub

Public Function IsRefresh() As Boolean
    IsRefresh = CmdState("Refresh")
End Function

'RemoveFormat
'Removes the formatting tags from the current selection.
Public Sub RemoveFormat()
    CmdExec "RemoveFormat"
End Sub

Public Function IsRemoveFormat() As Boolean
    IsRemoveFormat = CmdState("RemoveFormat")
End Function

'RemoveParaFormat
'Not currently supported.
Public Sub RemoveParaFormat()
    CmdExec "RemoveParaFormat"
End Sub

Public Function IsRemoveParaFormat() As Boolean
    IsRemoveParaFormat = CmdState("RemoveParaFormat")
End Function

'SaveAs
'Saves the current Web page to a file.
Public Sub SaveAs()
    CmdExec "SaveAs"
End Sub

Public Function IsSaveAs() As Boolean
    IsSaveAs = CmdState("SaveAs")
End Function

'SelectAll
'Selects the entire document.
Public Sub SelectAll()
    CmdExec "SelectAll"
End Sub

Public Function IsSelectAll() As Boolean
    IsSelectAll = CmdState("SelectAll")
End Function

'SizeToControl
'Not currently supported.
Public Sub SizeToControl()
    CmdExec "SizeToControl"
End Sub

'SizeToControlHeight
'Not currently supported.
Public Sub SizeToControlHeight()
    CmdExec "SizeToControlHeight"
End Sub

'SizeToControlWidth
'Not currently supported.
Public Sub SizeToControlWidth()
    CmdExec "SizeToControlWidth"
End Sub

'Stop
'Not currently supported.
Public Sub Stop1()
    CmdExec "Stop"
End Sub

'StopImage
'Not currently supported.
Public Sub StopImage()
    CmdExec "StopImage"
End Sub

'Strikethrough
'Not currently supported.
Public Sub Strikethrough()
    CmdExec "Strikethrough"
End Sub

Public Function IsStrikethrough() As Boolean
    IsStrikethrough = CmdState("Strikethrough")
End Function

'subScript
'Not currently supported.
Public Sub SubScript()
    CmdExec "SubScript"
End Sub

Public Function IsSubScript() As Boolean
    IsSubScript = CmdState("SubScript")
End Function

'superScript
'Not currently supported.
Public Sub SuperScript()
    CmdExec "SuperScript"
End Sub

Public Function IsSuperScript() As Boolean
    IsSuperScript = CmdState("SuperScript")
End Function

'UnBookmark
'Removes any bookmark from the current selection.
Public Sub UnBookmark()
    CmdExec "UnBookmark"
End Sub

Public Function IsUnBookmark() As Boolean
    IsUnBookmark = CmdState("UnBookmark")
End Function

'Underline
'Toggles the current selection between underlined and not underlined.
Public Sub Underline()
    CmdExec "Underline"
End Sub

Public Function IsUnderline() As Boolean
    IsUnderline = CmdState("Underline")
End Function

'Undo
'Not currently supported.
Public Sub Undo()
    CmdExec "Undo"
End Sub

Public Function IsUndo() As Boolean
    IsUndo = CmdEnabled("Undo")
End Function

'Unlink
'Removes any hyperlink from the current selection.
Public Sub Unlink()
    CmdExec "Unlink"
End Sub

Public Function IsUnlink() As Boolean
    IsUnlink = CmdState("Unlink")
End Function

'Unselect
'Clears the current selection.
Public Sub Unselect()
    CmdExec "Unselect"
End Sub

Public Function IsUnselect() As Boolean
    IsUnselect = CmdState("Unselect")
End Function

Public Sub UserControl_Show()

    If (UserControl.Ambient.UserMode) Then
        'Debug.Print "Run Time"
        UserControl.BackStyle = 0 'Transparent
    Else
        'Debug.Print "Design Time"
        UserControl.BackStyle = 1 '(Default) Opaque
    End If
    
End Sub

Public Sub SetDirty(bValue As Boolean)
    Dim VarIn As Variant
    VarIn = bValue
    CommandExec IDM_SETDIRTY, OLECMDEXECOPT_DONTPROMPTUSER, VarIn
End Sub

Public Function GetDirty() As Boolean
    GetDirty = CommandStatus(IDM_SETDIRTY)
End Function

'Displays a glyph for all the br tags.
Public Sub ShowBreakGlyph(bValue As Boolean)
    Dim VarIn As Variant
    VarIn = bValue
    CommandExec IDM_SHOWWBRTAGS, OLECMDEXECOPT_DONTPROMPTUSER, VarIn
End Sub

'Displays glyphs to show the location of all tags in a document.
Public Sub ShowAllGlyph(bValue As Boolean)
    Dim VarIn As Variant
    VarIn = bValue
    CommandExec IDM_SHOWALLTAGS, OLECMDEXECOPT_DONTPROMPTUSER, VarIn
End Sub

'Displays a glyph for all the area tags.
Public Sub ShowAreaGlyph(bValue As Boolean)
    Dim VarIn As Variant
    VarIn = bValue
    CommandExec IDM_SHOWAREATAGS, OLECMDEXECOPT_DONTPROMPTUSER, VarIn
End Sub

'Displays a glyph for all the comment tags.
Public Sub ShowCommentGlyph(bValue As Boolean)
    Dim VarIn As Variant
    VarIn = bValue
    CommandExec IDM_SHOWCOMMENTTAGS, OLECMDEXECOPT_DONTPROMPTUSER, VarIn
End Sub

'Displays all the tags shown in Microsoft Internet Explorer 4.0.
Public Sub ShowMiscGlyph(bValue As Boolean)
    Dim VarIn As Variant
    VarIn = bValue
    CommandExec IDM_SHOWMISCTAGS, OLECMDEXECOPT_DONTPROMPTUSER, VarIn
End Sub

'Displays a glyph for all the script tags.
Public Sub ShowScriptGlyph(bValue As Boolean)
    Dim VarIn As Variant
    VarIn = bValue
    CommandExec IDM_SHOWSCRIPTTAGS, OLECMDEXECOPT_DONTPROMPTUSER, VarIn
End Sub

'Displays a glyph for all the style tags
Public Sub ShowStyleGlyph(bValue As Boolean)
    Dim VarIn As Variant
    VarIn = bValue
    CommandExec IDM_SHOWSTYLETAGS, OLECMDEXECOPT_DONTPROMPTUSER, VarIn
End Sub

'Displays a glyph for all the unknown tags.
Public Sub ShowUnknownGlyph(bValue As Boolean)
    Dim VarIn As Variant
    VarIn = bValue
    CommandExec IDM_SHOWUNKNOWNTAGS, OLECMDEXECOPT_DONTPROMPTUSER, VarIn
End Sub

'Const IDM_RESPECTVISIBILITY_INDESIGN = 2405
'When this feature is activated, any element that has a visibility set
'to "hidden" or display property set to "none" will not be shown in both
'design mode and browse mode.
Public Sub RespectVisibility(bValue As Boolean)
    Dim VarIn As Variant
    VarIn = bValue
    'CommandExec IDM_RESPECTVISIBILITY_INDESIGN, OLECMDEXECOPT_DONTPROMPTUSER, VarIn
    
    Dim doc As IHTMLDocument2
    Set doc = WebBrowser1.Document
    'doc.execCommand "RESPECTVISIBILITYINDESIGN", False, False
    doc.execCommand "respectvisibilityindesign", False, VarIn
    
End Sub

'CommandExec IDM_PROTECTMETATAGS, OLECMDEXECOPT_DODEFAULT, True, vbNull
'And yet another hack for VID to not aggressively overwrite some meta tags.
Public Sub ProtectMetaTags(ByVal bValue As Boolean)
    Dim VarIn As Variant
    VarIn = bValue
    CommandExec IDM_PROTECTMETATAGS, OLECMDEXECOPT_DONTPROMPTUSER, VarIn, vbNull
End Sub

'Using Editing Glyphs
'http://msdn.microsoft.com/library/default.asp?url=/workshop/browser/editing/usingeditingglyphs.asp#Glyph_Table_String_Format

'for disabling selection handles
Public Sub DisableEditFocus(ByVal bValue As Boolean)
    Dim VarIn As Variant
    VarIn = bValue
    CommandExec IDM_DISABLE_EDITFOCUS_UI, OLECMDEXECOPT_DONTPROMPTUSER, VarIn, vbNull
End Sub

'
Public Sub KeepSelection(ByVal bValue As Boolean)
    Dim VarIn As Variant
    VarIn = bValue
    Const IDM_KEEPSELECTION = 2410
    CommandExec IDM_KEEPSELECTION, OLECMDEXECOPT_DONTPROMPTUSER, VarIn, vbNull
End Sub

Public Sub OverrideCursor(ByVal bValue As Boolean)
    Dim VarIn As Variant
    VarIn = bValue
    Const IDM_OVERRIDE_CURSOR = 2420
    CommandExec IDM_OVERRIDE_CURSOR, OLECMDEXECOPT_DONTPROMPTUSER, VarIn, vbNull
End Sub

'
Public Sub InsertObject()
    CommandExec IDM_INSERTOBJECT, OLECMDEXECOPT_DONTPROMPTUSER ', VarIn, vbNull
End Sub

Public Sub SendBackward()
    Dim VarIn As Variant
    VarIn = True
    'CommandExec IDM_SENDTOBACK, OLECMDEXECOPT_DONTPROMPTUSER, VarIn, vbNull
End Sub

'IDM_SIZETOCONTROLWIDTH
'IDM_VERTSPACEDECREASE
'IDM_ADDFAVORITES
'IDM_CREATESHORTCUT

'BringForward
Public Sub BringForward()
    'CmdExec IDM_BRINGFORWARD, OLECMDEXECOPT_DONTPROMPTUSER, True
    'CmdExec IDM_SNAPTOGRID, OLECMDEXECOPT_DONTPROMPTUSER, True
    Dim VarIn As Variant
    Dim varOut As Variant
    Dim Cmd As Variant
    Dim X As Boolean
    X = True
    VarIn = X
    varOut = vbNull
    'CommandTarget.Exec GUID_HTML, 8, OLECMDEXECOPT_DONTPROMPTUSER, VarIn, VarOut
    'WebBrowser1.ExecWB IDM_COPY, OLECMDEXECOPT_DODEFAULT, VarIn, VarOut
    
End Sub

Public Function GetBlockFormats() As String()
    
    On Error Resume Next
    
    Dim Formats() As String
    Dim VarIn As Variant
    Dim varOut As Variant
    ReDim varOut(50) As String
    Dim X As Integer
    
    For X = 1 To 50
        varOut(X) = Space(100)
    Next X
    
    CommandExec IDM_GETBLOCKFMTS, OLECMDEXECOPT_DONTPROMPTUSER, VarIn, varOut
    
    Dim Elements() As String
    
    ParseOleVariantStrArray varOut, Elements
    ReDim Formats(LBound(Elements) To UBound(Elements))
    For X = LBound(Elements) To UBound(Elements) - 1
        Formats(X) = Elements(X)
    Next X
    
    Erase varOut
    
    GetBlockFormats = Formats()

End Function

'Sort string arrays
Sub SortArray(inpArray() As String)

    Dim intRet
    Dim intCompare
    Dim intLoopTimes
    Dim strTemp
    
    For intLoopTimes = 1 To UBound(inpArray)
        For intCompare = LBound(inpArray) To UBound(inpArray) - 1
            intRet = StrComp(inpArray(intCompare), _
                     inpArray(intCompare + 1), vbTextCompare)
    
            If intRet = 1 Then 'String1 is greater than String2
                strTemp = inpArray(intCompare)
                inpArray(intCompare) = inpArray(intCompare + 1)
                inpArray(intCompare + 1) = strTemp
            End If
        Next
    Next

End Sub

Public Sub ClearSelection()
    Dim VarIn As Variant
    VarIn = 1
    CommandExec IDM_CLEARSELECTION, OLECMDEXECOPT_DONTPROMPTUSER, VarIn, vbNull
End Sub

'Public Sub MergeCells()
'    Dim VarIn As Variant
'    VarIn = vbNull
'    On Error Resume Next
'    CommandExec 2205&, OLECMDEXECOPT_DONTPROMPTUSER, VarIn, VarIn
'End Sub

Public Sub ShowGrid()
    Dim VarIn As Variant
    VarIn = vbNull
    On Error Resume Next
    CommandExec IDM_SHOWGRID, OLECMDEXECOPT_DONTPROMPTUSER, VarIn, VarIn
End Sub

Public Function GetSelectedElements() As Collection
    
    'IHTMLElement
    On Error Resume Next
    
    Dim Elements As New Collection
    Dim rg As IHTMLTxtRange
    Dim ctlRg As IHTMLControlRange

    ' Branch on the type of selection and
    ' get the element under the caret or the site selected object
    ' and return it
    Set Elements = Nothing
    
    Select Case m_oDocument.selection.Type

    Case "None", "Text"

        Set rg = m_oDocument.selection.createRange

        ' Collapse the range so that the scope of the
        ' range of the selection is the the caret. That way the
        ' parentElement method will return the element directly
        ' under the caret. If you don't want to change the state of the
        ' selection, then duplicate the range and collapse it

        If Not rg Is Nothing Then
            'rg.collapse
            'Set GetElementsUnderCaret = rg.parentElement
        End If

    Case "Control"
        ' An element is site selected
        Set ctlRg = m_oDocument.selection.createRange
        Dim Count As Long
        
        For Count = 1 To ctlRg.length
            Elements.Add ctlRg.Item(Count - 1)
        Next Count
 
    End Select
    
    Set GetSelectedElements = Elements
   
End Function

Public Sub UnderlineSelected()

    UnderlineElement GetElementUnderCaret
    
End Sub

Public Sub UnderlineElement(ByVal elem As IHTMLElement)

'   http://itwriting.co.uk/phorum/read.php?f=3&i=1615&t=1615

    Dim doc4 As MSHTML.IHTMLDocument4
    
    Dim ids As MSHTML.IDisplayServices
    Dim ims As MSHTML.IMarkupServices
    Dim imc As MSHTML.IMarkupContainer
    
    Dim irs As MSHTML.IHTMLRenderStyle
    
    Dim impStart As MSHTML.IMarkupPointer
    Dim impEnd As MSHTML.IMarkupPointer
    
    Dim idpStart As MSHTML.IDisplayPointer
    Dim idpEnd As MSHTML.IDisplayPointer
    
    Dim ihrs As MSHTML.IHighlightRenderingServices
    Dim ihs As MSHTML.IHighlightSegment
    
    'Dim isp As IServiceProvider
    'isp.QueryService
    '----------------------------------------------------------------
    Set doc4 = m_oDocument
    Set ids = doc4
    Set ims = doc4
    ' Get the markup container
    Set imc = doc4
    '----------------------------------------------------------------
    '---------------------------------------------
    ' Create the start markup pointer and position
    ' it after the beginning of the element
    '---------------------------------------------
    ims.CreateMarkupPointer impStart
    
    'impStart.MoveAdjacentToElement elem, mshtml._ELEMENT_ADJACENCY.ELEM_ADJ_AfterBegin
    impStart.MoveAdjacentToElement elem, ELEM_ADJ_AfterBegin
    
    ' Create a display pointer from the markup pointer
    
    ids.CreateDisplayPointer idpStart
    idpStart.MoveToMarkupPointer impStart, Nothing
    '---------------------------------------------
    ' Create the end markup pointer and position
    ' it before the end of the element
    '---------------------------------------------
    ims.CreateMarkupPointer impEnd
    impEnd.MoveAdjacentToElement elem, ELEM_ADJ_BeforeEnd
    '----------------------------------------------------------------
    ' Create a display pointer from the markup pointer
    
    ids.CreateDisplayPointer idpEnd
    idpEnd.MoveToMarkupPointer impEnd, Nothing
    'Dim ilineinf As MSHTML.ILineInfo
    'idpEnd.GetLineInfo ilineinf
    'Debug.Print ilineinf.x
    '----------------------------------------------------------------
    ' Create a render style
    Set irs = doc4.createRenderStyle(vbNull)
        
    ' Must set this, despite it supposedly being the default setting!
    irs.defaultTextSelection = "false"
    irs.textBackgroundColor = "White"
    irs.textColor = "Black"
    irs.textDecoration = "underline"
    irs.textDecorationColor = "red"
    irs.textUnderlineStyle = "wave"
    '----------------------------------------------------------------
    ' Add the segment
    
    'ihrs = DirectCast(doc4, MSHTML.IHighlightRenderingServices)
    Set ihrs = doc4
    ihrs.AddSegment idpStart, idpEnd, irs, ihs
    '----------------------------------------------------------------
End Sub

'paraElement.setAttribute("className","mystyle", 0);
'====================================================================
Function GetWebBrowserHandle() As Long

    Dim l As Long, sl As Long
    Dim strin As String * 256
    
    l = FindWindowEx(UserControl.hWnd, ByVal 0&, "Shell Embedding", vbNullString)
    l = FindWindowEx(l, ByVal 0&, "Shell DocObject View", vbNullString)
    l = FindWindowEx(l, ByVal 0&, "Internet Explorer_Server", vbNullString)
    sl = GetClassName(l, strin, 256) '/// l is the hWnd of the browser.
    'Caption = strin '// set the form's caption to the classname of the browser
    '///IF YOU WAIT TILL THE BROWSER HAS NAVIGATED AND THEN USE THE ABOVE _
    '///YOU WILL RECEIVE THE HANDLE OF THE WEBBROWSER. _
    '///OR IF YOU DONT WANT IT IN THIS AREA, FIRST ALLOW THE BROWSER TO NAVIGATE _
    '///THEN ADD IN A SUB ( EG: COMMAND_CLICK )
    GetWebBrowserHandle = l
End Function
'====================================================================
Public Sub BringToFront()
    
    Dim e As IHTMLEventObj
    Dim el As IHTMLElement
    Dim s As IHTMLStyle
    
    Set e = m_oDocument.parentWindow.event
    Set el = GetElementUnderCaret
    Set s = el.Style
    'Debug.Print "s.zIndex: "; s.zIndex
     
    s.zIndex = s.zIndex - 1
   
    'HTMLTableCell.Style.removeAttribut("YourCssClassName")

    'Dim st As IHTMLStyle2
    'Set st = el.Style
    'st.layoutGrid = "15 5"
    'Debug.Print "st.layoutGrid"; st.layoutGrid
    'st.layoutGridLine = 15
    'Dim fp As FPHTMLStyle
    
'http://groups.yahoo.com/group/delphi-dhtmledit/message/1305
'> if SnapToGrid
'> then begin
'> I := NewWidth mod SnapToGridX;
'> NewWidth := (NewWidth div SnapToGridX) * SnapToGridX;
'> if I > (SnapToGridX div 2)
'> then NewWidth := NewWidth + (SnapToGridX div 2);
'> end;
'>
    
End Sub

Public Sub SendToBack()
    
    Dim e As IHTMLEventObj
    Dim el As IHTMLElement
    Dim s As IHTMLStyle
    
    Set e = m_oDocument.parentWindow.event
    Set el = GetElementUnderCaret
    Set s = el.Style
    
    'Debug.Print "s.zIndex: "; s.zIndex
    s.zIndex = s.zIndex + 1
    
End Sub
'====================================================================
'====================================================================
'To get the focus of the web browser or any control
'1.   Giving the WBC the focus  :
'    WebBrowser1.SetFocus
'2. Calling GetFocus WinApi function ( after defining it) and store the result in a variable:
'   WBHwnd = GetFocus

'   Debug.Print 1 / 0
'   If Err <> 0 Then MsgBox "PLEASE READ THE NOTES ON THE DECLARATIONS SECTION OF THE FORM BEFORE RUNNING THIS SAMPLE"

'Public Sub IHTMLEditHost_SnapRect(ByVal pIElement As MSHTML.IHTMLElement, ByRef prcNew As MSHTML.tagRECT, ByVal eHandle As Long)
'
'    'http://itwriting.com/phorum/read.php?f=3&i=125&t=125
'    'Implements mshtml.IHTMLEditHost.SnapRect
'    Debug.Print "IHTMLEditHost_SnapRect"
'    'prcNew.bottom
'
'    'http://www.codeproject.com/internet/snaptogrid.asp#xx839830xx
'    Dim lWidth As Long
'    Dim lHeight As Long
'
'    lWidth = prcNew.Right - prcNew.Left
'    lHeight = prcNew.Bottom - prcNew.Top
'
'    Dim m_iSnap As Long
'    m_iSnap = 15
'
'    Select Case eHandle
'        Case ELEMENT_CORNER_NONE:
'            prcNew.Top = ((prcNew.Top + (m_iSnap / 2)) / m_iSnap) * m_iSnap
'            prcNew.Left = ((prcNew.Left + (m_iSnap / 2)) / m_iSnap) * m_iSnap
'            prcNew.Bottom = prcNew.Top + lHeight
'            prcNew.Right = prcNew.Left + lWidth
'        '    ELEMENT_CORNER_NONE = 0
'        '    ELEMENT_CORNER_TOP = 1
'        '    ELEMENT_CORNER_LEFT = 2
'        '    ELEMENT_CORNER_BOTTOM = 3
'        '    ELEMENT_CORNER_RIGHT = 4
'        '    ELEMENT_CORNER_TOPLEFT = 5
'        '    ELEMENT_CORNER_TOPRIGHT = 6
'        '    ELEMENT_CORNER_BOTTOMLEFT = 7
'        '    ELEMENT_CORNER_BOTTOMRIGHT = 8
'        Case Else:
'    End Select
'
'End Sub

' pvAddRefMe
'
' Increments the reference count of this control
'
'Public Sub pvAddRefMe()
'    Dim oUnk As olelib.IUnknown
'
'   Set oUnk = Me
'   oUnk.AddRef
'
'End Sub

'Public Sub IServiceProvider_QueryService(guidService As Guid, riid As Guid, ppvObject As Long)

'    Debug.Print "IServiceProvider_QueryService"
'
'    'if (guidService == SID_SHTMLEditHost && riid == IID_IHTMLEditHost)
'   If IsEqualGUID(guidService, SID_SHTMLEditHost) Then
'
'      Dim oISM As IHTMLEditHost
'
'      ' Increment the reference count
'      pvAddRefMe
'
'      Set oISM = Me
'
'      ' Return this object
'      MoveMemory ppvObject, oISM, 4&
'
'   Else
'
'      ' The service or interface is
'      ' not supported
'      Err.Raise E_NOINTERFACE
'
'   End If
'
'End Sub

'http://msdn.microsoft.com/workshop/browser/editing/impedithost.asp?frame=true
'STDMETHODIMP CBrowserHost::QueryService(REFGUID guidService,
'                                      REFIID riid,
'                                      void **ppv)
'{
'    HRESULT hr = E_NOINTERFACE;
'
'    if (guidService == SID_SHTMLEditHost && riid == IID_IHTMLEditHost)
'    {
'        // Create new CSnap object using ATL
'        CComObject<CSnap>* pSnap;
'        hr = CComObject<CSnap>::CreateInstance(&pSnap);
'
'        // Query the new CSnap object for IHTMLEditHost interface
'        hr = pSnap->QueryInterface(IID_IHTMLEditHost, ppv);
'
'        // Cache a pointer to ISnap so you can tell the Snap behavior
'        // when to snap and change the snap increment
'        m_spSnap = (ISnap*)NULL; // Clear any previous pointers
'        hr = pSnap->QueryInterface(IID_ISnap, (void**)&m_spSnap);
'
'        // Set the snap increment
'        hr = pSnap->put_SnapIncrement(m_lSnapIncrement);
'    }
'
'    return hr;
'}
'


'Public Function IOleClientSite_GetContainer() As olelib.IOleContainer
'   Err.Raise E_NOTIMPL
'End Function
'
'Public Function IOleClientSite_GetMoniker(ByVal dwAssign As olelib.OLEGETMONIKER, ByVal dwWhichMoniker As olelib.OLEWHICHMK) As olelib.IMoniker
'   Err.Raise E_NOTIMPL
'End Function
'
'Public Sub IOleClientSite_OnShowWindow(ByVal fShow As olelib.BOOL)
'   Err.Raise E_NOTIMPL
'End Sub
'
'Public Sub IOleClientSite_RequestNewObjectLayout()
'   Err.Raise E_NOTIMPL
'End Sub
'
'Public Sub IOleClientSite_SaveObject()
'
'End Sub
'
'Public Sub IOleClientSite_ShowObject()
'   Err.Raise E_NOTIMPL
'End Sub
'
'
'Public Sub IOleInPlaceSite_CanInPlaceActivate()
'End Sub
'
'Public Sub IOleInPlaceSite_ContextSensitiveHelp(ByVal fEnterMode As olelib.BOOL)
'End Sub
'
'Public Sub IOleInPlaceSite_DeactivateAndUndo()
'End Sub
'
'Public Sub IOleInPlaceSite_DiscardUndoState()
'End Sub
'
'Public Function IOleInPlaceSite_GetWindow() As Long
'   IOleInPlaceSite_GetWindow = UserControl.hWnd
'End Function
'
'Public Sub IOleInPlaceSite_GetWindowContext(ppFrame As olelib.IOleInPlaceFrame, ppDoc As olelib.IOleInPlaceUIWindow, lprcPosRect As olelib.RECT, lprcClipRect As olelib.RECT, lpFrameInfo As olelib.OLEINPLACEFRAMEINFO)
'
'   Set ppFrame = Me
'   Set ppDoc = Me
'
'   lpFrameInfo.hwndFrame = UserControl.hWnd
'
'End Sub
'
'Public Sub IOleInPlaceSite_OnInPlaceActivate()
'End Sub
'
'Public Sub IOleInPlaceSite_OnInPlaceDeactivate()
'End Sub
'
'Public Sub IOleInPlaceSite_OnPosRectChange(lprcPosRect As olelib.RECT)
'End Sub
'
'Public Sub IOleInPlaceSite_OnUIActivate()
'End Sub
'
'Public Sub IOleInPlaceSite_OnUIDeactivate(ByVal fUndoable As olelib.BOOL)
'End Sub
'
'Public Sub IOleInPlaceSite_Scroll(ByVal scrollX As Long, ByVal scrollY As Long)
'End Sub

'====================================================================
'           Table functions
'====================================================================
Public Sub InsertTable(Rows As Long, Cols As Long)
    
    Dim C As Integer
    Dim R As Integer
    Dim Table As String
    Dim Col As String
    Dim Row As String
    
    Table = "<table border=""1"" cellpadding=""2"" cellspacing=""0"" style=""border-collapse: collapse"" bordercolor=""#C0C0C0"" width=""100%"">"
    
    Row = ""
    For R = 1 To Rows
        Col = ""
        For C = 1 To Cols
            Col = Col & "<TD>&nbsp;</TD>"
        Next C
        Row = Row & "<TR>" & Col & "</TR>"
    Next R
    
    Table = Table & Row & "</TABLE>"
    InsertHTMLCode Table
End Sub

'====================================================================
'           Split Table Cell
'====================================================================
Public Sub SplitSelectedCell()
    
    SplitCell GetElementUnderCaret
    
End Sub

Public Sub SplitCell(ByVal el As MSHTML.IHTMLElement)
    
    ' Entry: el is the cell to split
    Dim elTemp As MSHTML.IHTMLElement
    Dim currentRow As MSHTML.IHTMLTableRow
    Dim Row As MSHTML.IHTMLTableRow
    Dim currentCell As MSHTML.IHTMLTableCell
    Dim Cell As MSHTML.IHTMLTableCell
    Dim tbl As MSHTML.IHTMLTable

    Dim column As Integer
    Dim currentColumn As Integer


    If el Is Nothing Then
        ' Cannot split the element, so drop through
    Else
        If TypeOf el Is MSHTML.IHTMLTableCell Then
            ' Get a handle on the current row and cell
            Set currentCell = el
            Set currentRow = el.parentElement

            If currentCell.colSpan > 1 Then
                ' This cell already spans more than one column, so just decrement colSpan
                currentCell.colSpan = currentCell.colSpan - 1
            Else
                ' First, work out which column we are in
                currentColumn = 0

                For Each Cell In currentRow.cells
                    ' Only scan up to and including the current cell
                    If Cell.cellIndex >= currentCell.cellIndex Then
                        Exit For
                    End If
                    currentColumn = currentColumn + Cell.colSpan
                Next Cell

                ' Bubble up to the table and save a handle
                Set elTemp = el

                Do While elTemp.tagName <> "TABLE"
                    Set elTemp = elTemp.parentElement
                Loop

                ' elTemp now contains the table
                Set tbl = elTemp

                For Each Row In tbl.Rows
                    If currentRow.RowIndex = Row.RowIndex Then
                        ' On the current row we will insert a cell when done
                    Else
                        ' Locate the cell that corresponds to the current column
                        column = 0

                        For Each Cell In Row.cells
                            If column + Cell.colSpan > currentColumn Then
                                Exit For
                            End If
                            column = column + Cell.colSpan
                        Next Cell

                        ' Bump the colspan
                        Cell.colSpan = Cell.colSpan + 1
                    End If
                Next Row
            End If

            ' Finally, insert an extra cell, after this one, on the current row
            currentRow.insertCell (currentCell.cellIndex + 1)
        End If
    End If
End Sub
'====================================================================
'====================================================================
Public Sub SelectedCells()

    Dim StartCell As MSHTML.IHTMLTableCell
    Dim EndCell As MSHTML.IHTMLTableCell
    
    GetSelectedCells StartCell, EndCell
    If Not StartCell Is Nothing Then
        'Debug.Print "Start Cell: "; StartCell.cellIndex
    End If
    
    If Not EndCell Is Nothing Then
        'Debug.Print "End Cell: "; EndCell.cellIndex
    End If
    
End Sub

Public Sub GetSelectedCells(StartCell As MSHTML.IHTMLTableCell, EndCell As MSHTML.IHTMLTableCell)
'MSHTML.IHTMLElementCollection
    On Error Resume Next
    
    Dim rg As IHTMLTxtRange
    Dim rg1 As IHTMLTxtRange
    Dim ctlRg As IHTMLControlRange

    ' Branch on the type of selection and
    ' get the element under the caret or the site selected object
    ' and return it

    Select Case m_oDocument.selection.Type

    Case "None", "Text"

        Set rg = m_oDocument.selection.createRange

        ' Collapse the range so that the scope of the
        ' range of the selection is the the caret. That way the
        ' parentElement method will return the element directly
        ' under the caret. If you don't want to change the state of the
        ' selection, then duplicate the range and collapse it
        
        'http://groups.google.com/group/microsoft.public.inetsdk.programming.dhtml_editing/browse_frm/thread/ea3e39b2afe0dbf7/0ef42ce8ec99b235?lnk=st&q=split+cells+mshtml&rnum=2&hl=en#0ef42ce8ec99b235
        
        If Not rg Is Nothing Then
            Set rg1 = rg.duplicate
            rg.collapse True
            rg1.collapse False
            Set GetElementUnderCaret = rg.parentElement
            Set StartCell = rg.parentElement
            Set EndCell = rg1.parentElement
        End If

    Case "Control"
        ' An element is site selected
        Set ctlRg = m_oDocument.selection.createRange
 
        ' There can only be one site selected element at a time so the
        ' commonParentElement will return the site selected element
        'Set GetElementUnderCaret = ctlRg.commonParentElement
        
            Set StartCell = Nothing
            Set EndCell = Nothing
    End Select

End Sub

'Script procedures within an HTML document can be called through the document's window object.
'Set htmDoc = Me.WebBrowser1.Document
'Call htmDoc.parentWindow.subroutine1(param1, param2)

Public Sub InsertRows()

    Dim el As IHTMLElement
    Dim Row As IHTMLTableRow
    Dim td As IHTMLTableCell
    Dim Cell As IHTMLTableCell
    Dim Table As IHTMLTable
    Dim Cols As Long, Col As Long, RowIndex As Long
    
    Set el = GetElementUnderCaret
    If Not TypeOf el Is IHTMLTableCell Then Exit Sub
    
    Set td = el
    Set Row = el.parentElement
    RowIndex = Row.RowIndex
    Cols = Row.cells.length
    Set Table = td.offsetParent
    Set Row = Table.insertRow(RowIndex + 1) '-1: end of table
    For Col = 1 To Cols
        Row.insertCell (-1)
    Next Col
    
End Sub

Public Sub InsertColumns()

    Dim el As IHTMLElement
    Dim Row As IHTMLTableRow
    Dim td As IHTMLTableCell
    Dim Cell As IHTMLTableCell
    Dim Table As IHTMLTable
    Dim Cols As Long, Col As Long, RowIndex As Long
    Dim cellIndex As Long
    
    Set el = GetElementUnderCaret
    If Not TypeOf el Is IHTMLTableCell Then Exit Sub
    
    Set td = el
    Set Table = td.offsetParent 'Retrieves a reference to the container object that defines the offsetTop and offsetLeft properties of the object.
    Table.Cols = Table.Cols + 1
    'Table.Rows.length
    cellIndex = td.cellIndex + 1
    For Each Row In Table.Rows
        Set Cell = Row.insertCell(cellIndex)
        Cell.innerHTML = "&nbsp;"
    Next Row

End Sub

Public Sub InsertCells()

    Dim el As IHTMLElement
    Dim Row As IHTMLTableRow
    Dim td As IHTMLTableCell
    Dim Cell As IHTMLTableCell
    Dim Table As IHTMLTable
    Dim Cols As Long, Col As Long, RowIndex As Long
    Dim cellIndex As Long
    
    Set el = GetElementUnderCaret
    If Not TypeOf el Is IHTMLTableCell Then Exit Sub
    Set td = el
    Set Row = el.parentElement
    Set Table = td.offsetParent 'Retrieves a reference to the container object that defines the offsetTop and offsetLeft properties of the object.
    Table.Cols = Table.Cols + 1
    'Table.Rows.length
    cellIndex = td.cellIndex + 1
    Set Cell = Row.insertCell(cellIndex)
    Cell.innerHTML = "&nbsp;"

End Sub

Public Sub DeleteRows()

    Dim el As IHTMLElement
    Dim Row As IHTMLTableRow
    Dim td As IHTMLTableCell
    Dim Cell As IHTMLTableCell
    Dim Table As IHTMLTable
    Dim Cols As Long, Col As Long, RowIndex As Long
    
    Set el = GetElementUnderCaret
    If Not TypeOf el Is IHTMLTableCell Then Exit Sub
    
    Set td = el
    Set Row = el.parentElement
    RowIndex = Row.RowIndex
    
    Set Table = td.offsetParent
    Table.deleteRow (RowIndex)

End Sub

Public Sub DeleteColumns()

    Dim el As IHTMLElement
    Dim Row As IHTMLTableRow
    Dim td As IHTMLTableCell
    Dim Cell As IHTMLTableCell
    Dim Table As IHTMLTable
    Dim Cols As Long, Col As Long, RowIndex As Long
    Dim cellIndex As Long
    
    Set el = GetElementUnderCaret
    If Not TypeOf el Is IHTMLTableCell Then Exit Sub
    Set td = el
    Set Table = td.offsetParent 'Retrieves a reference to the container object that defines the offsetTop and offsetLeft properties of the object.
    Table.Cols = Table.Cols + 1
    'Table.Rows.length
    cellIndex = td.cellIndex
    For Each Row In Table.Rows
        Row.deleteCell (cellIndex)
    Next Row

End Sub

Public Sub DeleteCells()
    
    Dim el As IHTMLElement
    Dim Row As IHTMLTableRow
    Dim td As IHTMLTableCell
    Dim Cell As IHTMLTableCell
    Dim Table As IHTMLTable
    Dim Cols As Long, Col As Long, RowIndex As Long
    Dim cellIndex As Long
    
    Set el = GetElementUnderCaret
    If Not TypeOf el Is IHTMLTableCell Then Exit Sub
    Set td = el
    Set Row = el.parentElement
    Set Table = td.offsetParent 'Retrieves a reference to the container object that defines the offsetTop and offsetLeft properties of the object.
    Table.Cols = Table.Cols + 1
    'Table.Rows.length
    cellIndex = td.cellIndex
    Row.deleteCell (cellIndex)

End Sub
Public Sub MergeCells()

End Sub
Public Sub SplitCells()

End Sub

Sub InsertHTMLCode(strHTML As String)
    
    On Error Resume Next
    
    Dim rg As IHTMLTxtRange
    Dim ctlRg As IHTMLControlRange
    Dim sel As IHTMLSelectionObject
    
    Set rg = Nothing
    Set ctlRg = Nothing
    
    Set sel = m_oDocument.selection
    Select Case m_oDocument.selection.Type
        Case "Control":
            Set ctlRg = sel.createRange
            If Not ctlRg Is Nothing Then
                m_oDocument.selection.Clear
                Set rg = sel.createRange
                If Not rg Is Nothing Then
                    rg.collapse True
                    rg.pasteHTML (strHTML)
                End If
            End If
            
        Case "None", "Text":
            Set rg = sel.createRange
            If Not rg Is Nothing Then
                m_oDocument.selection.Clear
                rg.collapse True
                rg.pasteHTML (strHTML)
            End If
    End Select
    
    Set rg = Nothing
    Set ctlRg = Nothing
    Set sel = Nothing
    
End Sub

'====================================================================
Public Function GetFontNames() As String()

    Dim X As Long
    Dim FontNames() As String
    
    ReDim FontNames(0 To Screen.FontCount - 1)
    For X = 0 To Screen.FontCount - 1
        FontNames(X) = Screen.Fonts(X)
    Next X
    
    SortArray FontNames()
    GetFontNames = FontNames
    
End Function

'====================================================================
Public Sub SetZoom(iValue As Long)
    WebBrowser1.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, iValue
End Sub
'====================================================================
'====================================================================
'http://www.qbdsoftware.co.uk/moth/qweb/resprotocol.htm
'To load an HTML document from a resource file in Visual Basic use:
'WebBrowser.Navigate2 "res://myres.dll/%2323/%23101"
'res://path and filename[/resource type]/resource ID
'HTML: src="res://myresource.dll/%2323/%23101"
'Bitmap: src="res://myresource.dll/%232/%23102"
'Custom: src="res://myresource.dll/GIF/%23103"

'Dim StyleS As MSHTML.IHTMLStyle
'Set StyleS = DHtmlImage.Style
'StyleS.backgroundColor = "#000000"
'StyleS.fontFamily = "times"
'====================================================================
'====================================================================

Public Property Get Visible() As Boolean
    Visible = Me.Visible
End Property

Public Property Let Visible(ByVal bVisible As Boolean)
    Me.Visible = bVisible
    PropertyChanged "Visible"
End Property

Public Property Get SelectTables() As Boolean
    SelectTables = m_bSelectTables
End Property

Public Property Let SelectTables(ByVal bVisible As Boolean)
    m_bSelectTables = bVisible
    PropertyChanged "SelectTables"
End Property

