VERSION 5.00
Begin VB.UserControl sMenu 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5850
   ControlContainer=   -1  'True
   PropertyPages   =   "SkinMENU.ctx":0000
   ScaleHeight     =   28
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   390
   ToolboxBitmap   =   "SkinMENU.ctx":003C
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   0
   End
End
Attribute VB_Name = "sMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public MenuTree   As New Collection
Public MenuOpen   As Boolean
'Public Vertical   As Boolean

Private LastSel      As Long
Private m_Spacing    As Long
Private m_Style      As Long
Private m_HighLight  As Long
Private m_ItemsFont  As StdFont
Private m_BackColor  As OLE_COLOR
Private m_Vertical   As Boolean

Private wf As New wMenu
Private Flag As Long

Public Event ItemClick(Key As String)
Public Event ItemDescription(Description As String)

Public Enum cmnStyleConstants
   mnClassic = 0
   mnWindowsXP = 1
   mnExperiment = 2
   mnOpenStep = 3
End Enum



Private Function CollapseAll() As Boolean
 Dim i As Long
 On Error Resume Next
 For i = 1 To MenuTree.Count
        If MenuTree(i).Expanded Then
           DummyForm.Top = -1000
           DummyForm.Show ' serve per non far passare il focus all'applicazione principale (anti-flickering)
           Unload wf
           CollapseAll = True
         '  Refresh
           Exit Function
        End If
 Next
 
End Function

Public Function KeyToIndex(ByVal Index As Variant) As Integer
 '##BD Returns the integer index value of a button identified by either it's Key property or index.
 Dim i
 If Val(Index) = 0 Then
  For i = 1 To MenuTree.Count
   If UCase$(MenuTree(i).Name) = UCase$(Index) Then
    KeyToIndex = i
    Exit Function
   End If
  Next
 Else
  KeyToIndex = Val(Index)
 End If
 If KeyToIndex = 0 And Index <> 0 Then
  RaiseErrorEx "KeyToIndex", 35601, "Element not found. Key is missing or illegal."
 End If
End Function

Private Sub RaiseErrorEx(ByVal ProcName$, ByVal ErrNum As Long, Optional ByVal ErrMsg$ = "")
 If Ambient.UserMode Then
  '"Runtime" - raise error
  If Len(ErrMsg$) Then
   Err.Raise ErrNum, App.EXEName & "." & TypeName(Me) & ":" & ProcName$, ErrMsg$
  Else
   Err.Raise ErrNum, App.EXEName & "." & TypeName(Me) & ":" & ProcName$
  End If
 Else
  '"Design time" - display error
  If Len(ErrMsg$) = 0 Then
   On Error Resume Next
    Error ErrNum
    ErrMsg$ = Err.Description
   On Error GoTo 0
  End If
  VBA.MsgBox INTERR$ & vbCr & vbCr & ErrMsg$ & " (" & ErrNum & ")" & vbCr & vbCr & ERRTEXT$, vbCritical, App.EXEName & "." & TypeName(Me)
 End If
End Sub

Sub SendDescriptionEvent(Description As String)
  RaiseEvent ItemDescription(Description)
End Sub

Private Sub DrawOpenStepItem(x1 As Long, y1 As Long, X2 As Long, Y2 As Long, FromColor As Long, ToColor As Long, Mode As Long)
                       
   Dim c() As Long, Y As Long, n As Long, X As Long
 '  Line (x1, y1)-(x2, y1), QBColor(0), BF
                       
 On Error Resume Next


   
   n = (X2 - x1) / 20
   ReDim c(20)
   
   CreateGradateColors FromColor, ToColor, c
   
   For X = x1 To X2 Step n
       Line (X, y1)-(X + n, Y2), c(Y), BF
       Y = Y + 1
   Next

                  
                    Line (x1 + 1, y1 + 1)-(X2 - 1, y1 + 1), QBColor(15)
                    Line -(X2 - 1, Y2 - 1), QBColor(8)
                    Line -(x1 + 1, Y2 - 1), QBColor(8)
                    Line -(x1 + 1, y1 + 1), QBColor(15)
   
   UserControl.Refresh


End Sub
Property Get BackColor() As OLE_COLOR
   BackColor = m_BackColor
End Property
Property Let BackColor(Color As OLE_COLOR)
   m_BackColor = Color
End Property
Property Let Changed(Value As Boolean)
   PropertyChanged "Spacing"
End Property
Public Property Get ItemsFont() As StdFont
  Set ItemsFont = m_ItemsFont
End Property
Public Property Set ItemsFont(ByVal ItemsFont As StdFont)
 If ItemsFont Is Nothing Then
 ' RaiseErrorEx "ToolTipFont", 424
 Else
  Set m_ItemsFont = ItemsFont
  PropertyChanged "ItemsFont"
 End If
End Property


Private Sub CleanUP()
  Dim i As Long
  For i = 1 To MenuTree.Count
      MenuTree.Remove 1
  Next
End Sub
Property Get Vertical() As Boolean
  Vertical = m_Vertical
End Property

Property Let Vertical(Value As Boolean)
  m_Vertical = Value
  PropertyChanged "Vertical"
End Property
Property Get HighLightStyle() As Long
  HighLightStyle = m_HighLight
End Property

Property Let HighLightStyle(Value As Long)
  m_HighLight = Value
  PropertyChanged "HighLightStyle"
End Property
Property Get hwnd() As Long
  hwnd = UserControl.hwnd
End Property


Property Let ItemChecked(Key As String, Value As Boolean)
   Dim i As Long
   For i = 1 To MenuTree.Count
      If UCase(MenuTree(i).Name) = UCase(Key) Then
         MenuTree(i).Checked = Value
         Exit Property
      End If
   Next
End Property

Property Get Name() As String
  Name = "xMenu"
End Property

Sub Refresh()
   Dim i As Long, idn As Long, od As Long
   Dim X As Long, x1 As Long
   Dim Y As Long, y1 As Long
   
   Dim K As Long, k1 As Long
   Dim L As Long, l1 As Long
   
   
 '  Spacing = 14
   FontName = m_ItemsFont '"Tahoma"
   FontSize = m_ItemsFont.Size '8
   FontBold = m_ItemsFont.Bold '8
   X = Spacing * 0.5
   Y = Spacing * 0.5
   
   Select Case Style
     Case mnClassic
        BackColor = vbButtonFace
     '   BackColor = Ambient.BackColor
     Case mnWindowsXP
        BackColor = vbButtonFace
     Case mnExperiment
        BackColor = vbWhite
   End Select
   BackColor = vbButtonFace
   
   Cls
   
   If Style = mnOpenStep Then
      DrawOpenStepItem 0, 0, Width / Screen.TwipsPerPixelX, Height / Screen.TwipsPerPixelY, QBColor(15), QBColor(7), 0
   End If
   
  ' cre le informazioni ad albero
   For i = 1 To MenuTree.Count
       SetFilter i
       MenuTree(i).IsRoot = HasSubItems(i)
   Next
       
   If Vertical Then
      For i = 1 To MenuTree.Count
          idn = MenuTree(i).Ident
          If idn = 0 And MenuTree(i).Visible Then
             If MenuTree(i).IsRoot Then
                CurrentY = L
                CurrentX = 6
               ' Print MenuTree(i).Caption
                l1 = L + Spacing * 0.5 + TextHeight(MenuTree(i).Caption) + Spacing * 0.5
                k1 = K + Spacing * 0.5 + TextWidth(MenuTree(i).Caption) + Spacing * 0.5
                MenuTree(i).SetRect K, CurrentY - 2, k1, CurrentY + TextHeight("A") + 2
                MenuTree(i).X = 2 'X  + Spacing * 0.5
                MenuTree(i).Y = CurrentY
                L = l1
                RefreshItem i, 0
             
             Else
                CurrentY = 4
                CurrentX = X
               ' Print MenuTree(i).Caption
                x1 = X + Spacing * 0.5 + TextWidth(MenuTree(i).Caption) + Spacing * 0.5
                MenuTree(i).SetRect X, CurrentY - 2, x1, CurrentY + TextHeight("A") + 2
                MenuTree(i).X = X + Spacing * 0.5
                MenuTree(i).Y = CurrentY
                X = x1
                RefreshItem i, 0
             End If
          End If
      Next
   Else
      For i = 1 To MenuTree.Count
          idn = MenuTree(i).Ident
          If idn = 0 And MenuTree(i).Visible Then
             CurrentY = 4
             CurrentX = X
            ' Print MenuTree(i).Caption
             x1 = X + Spacing * 0.5 + TextWidth(MenuTree(i).Caption) + Spacing * 0.5
             MenuTree(i).SetRect X, CurrentY - 2, x1, CurrentY + TextHeight("A") + 2
             MenuTree(i).X = X + Spacing * 0.5
             MenuTree(i).Y = CurrentY
             X = x1
             RefreshItem i, 0
          End If
      Next
   End If
End Sub

Private Function PointToItem(X As Single, Y As Single) As Long
  Dim i As Long
  For i = 1 To MenuTree.Count
     If MenuTree(i).PointInside(X, Y) Then
        If MenuTree(i).Ident = 0 Then ' deve essere uno degli items principali
           PointToItem = i
           Exit Function
        End If
     End If
  Next
End Function

Private Sub RefreshItem(Index As Long, Mode As Long)
   Dim X As Long, Y As Long, x1 As Long, y1 As Long, X2 As Long, Y2 As Long
   Dim d As Long
   Dim b As Long
   b = BackColor
      
If MenuOpen And Mode = 1 Then Mode = -1
Select Case Style
 
  Case mnClassic
   
    With MenuTree(Index)
        .GetRect x1, y1, X2, Y2
        
        Line (x1, y1)-(X2, Y2), b, BF
        Select Case Mode
           Case 0 ' normal, no pressed
              CurrentX = .X
              CurrentY = .Y
              Print .Caption
           Case 1 ' up
              CurrentX = .X - 1
              CurrentY = .Y - 1
              Print .Caption
              Line (x1, y1)-(X2, y1), QBColor(15)
              Line -(X2, Y2), QBColor(8)
              Line -(x1, Y2), QBColor(8)
              Line -(x1, y1), QBColor(15)
           Case -1 ' Pressed
              CurrentX = .X + 1
              CurrentY = .Y + 1
              Print .Caption
              Line (x1, y1)-(X2, y1), QBColor(8)
              Line -(X2, Y2), QBColor(15)
              Line -(x1, Y2), QBColor(15)
              Line -(x1, y1), QBColor(8)
              
       End Select
       
   End With
   
 Case mnWindowsXP
 
    With MenuTree(Index)
        .GetRect x1, y1, X2, Y2
        Line (x1, y1)-(X2, Y2), b, BF
        Select Case Mode
           Case 0 ' normal, no pressed
              CurrentX = .X
              CurrentY = .Y
              Print .Caption
           Case 1 ' up
              d = 2
           '   Line (x1 + D, y1)-(x2 - D, y1), QBColor(15)
           '   Line -(x2, y1 + D), QBColor(8)
           '   Line -(x2, y2 - D), QBColor(8)
           '   Line -(x2 - D, y2), QBColor(8)
           '   Line -(x1 + D, y2), QBColor(15)
           '   Line -(x1, y2), QBColor(15)
              
              Line (x1, y1)-(X2, Y2), QBColor(15), B
              Line (x1 + d, y1 + d)-(X2 - d, Y2 - d), RGB(220, 220, 220), BF
              
              CurrentX = .X
              CurrentY = .Y
              Print .Caption
           Case -1 ' Pressed
              CurrentX = .X + 1
              CurrentY = .Y + 1
              Print .Caption
              Line (x1, y1)-(X2 - 3, y1), QBColor(15)
              Line -(X2, Y2), QBColor(8)
              Line -(x1, Y2), QBColor(8)
              Line -(x1, y1), QBColor(15)
              
       End Select
       
   End With
   
 Case mnExperiment
 
    With MenuTree(Index)
        .GetRect x1, y1, X2, Y2
        Line (x1, y1)-(X2, Y2), b, BF
        Select Case Mode
           Case 0 ' normal, no pressed
              CurrentX = .X
              CurrentY = .Y
              Print .Caption
           Case 1 ' up
'              Fade x1 + 1, y1 + 1, x2 - 1, y2 - 1, RGB(148, 195, 198), QBColor(15), 1
            '  Fade x1 + 1, y1 + 1, x2 - 1, y2 - 1, QBColor(15), RGB(148, 195, 198), 1
              Fade x1 + 1, y1 + 1, X2 - 1, Y2 - 1, QBColor(7), QBColor(15), 1
              DrawWidth = 10
              PSet (x1, y1), BackColor
              PSet (X2, y1), BackColor
              PSet (x1, Y2), BackColor
              PSet (X2, Y2), BackColor
              DrawWidth = 1
              
              
              b = ForeColor
              ForeColor = QBColor(8)
              CurrentX = .X
              CurrentY = .Y
              Print .Caption
              ForeColor = b
           Case -1 ' Pressed
              Fade x1 + 1, y1 + 1, X2 - 1, Y2 - 1, QBColor(15), QBColor(7), 1
              DrawWidth = 10
              PSet (x1, y1), BackColor
              PSet (X2, y1), BackColor
              PSet (x1, Y2), BackColor
              PSet (X2, Y2), BackColor
              DrawWidth = 1
              
              
              b = ForeColor
              ForeColor = QBColor(0)
              CurrentX = .X
              CurrentY = .Y
              Print .Caption
              ForeColor = b
       End Select
       
   End With
   
 Case mnOpenStep
 
    With MenuTree(Index)
        .GetRect x1, y1, X2, Y2
        Line (x1, y1)-(X2, Y2), b, BF
        Select Case Mode
           Case 0 ' normal, no pressed
              DrawOpenStepItem x1, y1, X2, Y2, QBColor(15), QBColor(7), Mode
              CurrentX = .X
              CurrentY = .Y
              Print .Caption
           Case 1 ' up
              DrawOpenStepItem x1, y1, X2, Y2, QBColor(15), QBColor(7), Mode
              b = ForeColor
              ForeColor = QBColor(9)
              CurrentX = .X
              CurrentY = .Y
              Print .Caption
              ForeColor = b
              
           Case -1 ' Pressed
              DrawOpenStepItem x1, y1, X2, Y2, QBColor(7), QBColor(15), Mode
              
              b = ForeColor
              ForeColor = QBColor(0)
              CurrentX = .X + 1
              CurrentY = .Y + 1
              Print .Caption
              ForeColor = b
       End Select
       
   End With
   

End Select

End Sub

Private Sub Fade(x1 As Long, y1 As Long, X2 As Long, Y2 As Long, FromColor As Long, ToColor As Long, Optional Mode As Long = 0)
   Dim c() As Long, Y As Long, n As Long, X As Long
   
Select Case Mode
Case 0 ' gradient orizzontale
   n = X2 - x1
   ReDim c(n)
   
   CreateGradateColors FromColor, ToColor, c
   
   For X = 0 To n
       Line (x1 + X, y1)-(x1 + X, Y2), c(X)
   Next
Case 1 ' verticale a tubo
   
   n = (Y2 - y1) * 0.5
   ReDim c(n)

   CreateGradateColors FromColor, ToColor, c
   
   For Y = 0 To n
       Line (x1, y1 + Y)-(X2, y1 + Y), c(Y)
   Next
   For Y = 0 To n
       Line (x1, Y2 - Y)-(X2, Y2 - Y), c(Y)
   Next

End Select
   
End Sub


Private Sub DrawTriangle(x1 As Long, y1 As Long, X2 As Long, Y2 As Long, Color As Long)
    Dim hBrush As Long, hPrev As Long, h2 As Long, d As Long
    Dim Y As Long, X As Long, b As Long
    Dim pt(2) As POINTAPI
    
    d = 5
    
    Y = y1 + (Y2 - y1) * 0.5
    X = X2 - d
    
'                    Line (x2 - Dx, y1 + Dy)-(x2 - Dx, y2 - Dy), QBColor(15)
'                    Line -(x2 - 2, Y), QBColor(15)
'                    Line -(x2 - Dx, y1 + Dy), QBColor(15)
   
   pt(0).X = X: pt(0).Y = Y
   pt(1).X = X - d: pt(1).Y = Y - d
   pt(2).X = X - d: pt(2).Y = Y + d
   
   b = ForeColor
   ForeColor = Color
   hBrush = CreateSolidBrush(Color)
   hPrev = SelectObject(hdc, hBrush)
   
   h2 = Polygon(hdc, pt(0), 3)
   
   h2 = SelectObject(hdc, hPrev)
   DeleteObject hBrush
   ForeColor = b
End Sub

Sub SendClickEvent(Key As String)
  '  Dim t As Double
  '  t = Timer
  '  Do While Timer - t < 0.5: Loop
    RaiseEvent ItemClick(Key)
End Sub

Private Sub SetFilter(Index As Long)
    Dim idn As Long
    idn = MenuTree(Index).Ident + 1
    For i = Index + 1 To MenuTree.Count
        If MenuTree(i).Ident < idn Then Exit For
        If MenuTree(i).Ident = idn Then
           MenuTree(i).RootIndex = Index
        End If
    Next

End Sub

Sub ShowMenu(Index As Long)
    Dim idn As Long, i As Long, P As POINTAPI
    Dim X As Long, Y As Long, w As Long, H As Long
    Dim j As Long
    
    On Error Resume Next
    Unload wf
    On Error GoTo 0
    
    MenuTree(Index).GetRect X, Y, w, H
    
    '*Alinhado ao Item
    'P.X = X + IIf(Vertical, w + X, 0)
    'P.Y = H - IIf(Vertical, TextHeight("A") + 2, 0)
    'Alinhado ao Menu
    P.X = X + IIf(Vertical, UserControl.Width / Screen.TwipsPerPixelX, 0)
    P.Y = H - IIf(Vertical, H, 0)
    ClientToScreen hwnd, P
   
   ' ClientToScreen ContainerhWnd, P
    P.X = P.X * Screen.TwipsPerPixelX
    P.Y = P.Y * Screen.TwipsPerPixelY
    Flag = True
    Set wf = New wMenu
    Load wf
    Set wf.MenuTree = MenuTree
    Set wf.ParentForm = Me
    wf.Font.Name = m_ItemsFont.Name
    wf.Font.Size = m_ItemsFont.Size
    wf.Font.Bold = m_ItemsFont.Bold
    
    wf.FilterIndex = Index
    wf.LeftCorner = P.X
    wf.TopCorner = P.Y
    wf.Style = Style
    wf.OwnerWidth = w - X
    wf.HighLight = HighLightStyle
    wf.FillItems
    On Error Resume Next
    wf.Show
    If Err.Number = 401 Then
       RaiseErrorEx "ShowMenu", 35607, "Can't show menu on modal forms."
     '  MsgBox Err.Number & " " & Err.Description, Err.Title
       Exit Sub
    End If
    On Error GoTo 0
  
    MenuTree(Index).Expanded = True
  '  LastSel = 0

    Flag = False
End Sub


Private Function HasSubItems(Index As Long) As Boolean
    Dim i As Long, idn As Long
    
' True se l'tem I ha subitems.

    idn = MenuTree(Index).Ident + 1
    For i = Index + 1 To MenuTree.Count
        If MenuTree(i).Ident < idn Then Exit For
        If MenuTree(i).Ident = idn Then
           HasSubItems = True
           Exit Function
        End If
    Next

End Function


Property Get Spacing() As Long
   Spacing = m_Spacing
End Property

Property Let Spacing(Value As Long)
   m_Spacing = Value
   PropertyChanged "Spacing"
End Property



Property Get Style() As Long
   Style = m_Style
End Property

Property Let Style(Value As Long)
 m_Style = Value
 PropertyChanged "Style"
End Property

Private Sub Timer1_Timer()
  Dim H As Long, P As POINTAPI
  If LastSel <> 0 Then
     GetCursorPos P
     ScreenToClient hwnd, P
     If MenuOpen = False Then
       If P.X < 0 Or P.X > Width / Screen.TwipsPerPixelX Or P.Y < 0 Or P.Y > Height / Screen.TwipsPerPixelY Then
           RefreshItem LastSel, 0
           LastSel = 0
           Timer1.Enabled = False
'          Debug.Print "Fuori"
       Else
   '       Debug.Print "Dentro"
       End If   ' il mouse non è sul controllo
    End If
 End If
  ' Exit Sub
     
'     h = GetActiveWindow
   '  If h <> hwnd Then
'           RefreshItem LastSel, 0
'           LastSel = 0
'           Timer1.Enabled = False
  '      End If
  '   End If
'  End If

End Sub

Private Sub UserControl_Initialize()
   AutoRedraw = -1
   m_HighLight = 1 ' solid bar
' Set m_ItemsFont = Ambient.Font
End Sub
Private Sub UserControl_InitProperties()
  Spacing = 12
  Set m_ItemsFont = Ambient.Font
  m_ItemsFont.Bold = False
  m_Vertical = False
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim i As Long
  
  i = PointToItem(X, Y)
  
  If i = 0 Then Exit Sub
  If LastSel <> i Then
     If LastSel <> 0 Then RefreshItem LastSel, 0
     LastSel = i
  End If
  RefreshItem i, -1
  UserControl.Refresh
  
  If MenuOpen = 0 Then
     ShowMenu i
     DoEvents
     MenuOpen = True
  Else
     Unload wf
     MenuOpen = False
  End If
  
End Sub




Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim i As Long, Sm As Integer
  Static SameItem As Long
  
  If Timer1.Enabled = False Then Timer1.Enabled = True
  If Button = 1 Then Exit Sub
  
  i = PointToItem(X, Y)
  If i = 0 And MenuOpen Then Exit Sub
  
  If LastSel <> i Then
     RaiseEvent ItemDescription("")
     If LastSel <> 0 Then
        RefreshItem LastSel, 0
       ' Sm = CollapseAll
        If MenuTree(LastSel).Expanded Then
           DummyForm.Top = -1000
           DummyForm.Show ' serve per non far passare il focus all'applicazione principale (anti-flickering)
           Unload wf
           Sm = True
       '    DoEvents
        
        End If
     End If
     LastSel = i
  End If
  
  If i = 0 Then Exit Sub
  
  RefreshItem i, 1
  If Sm Then
     ShowMenu i
     MenuOpen = True
  End If

End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  Dim i As Long, cnt As Long
  Dim o As New MenuItem
  
  m_Spacing = PropBag.ReadProperty("Spacing", 12)
  m_Style = PropBag.ReadProperty("Style", 0)
  cnt = PropBag.ReadProperty("ItemCount", 0)
  m_HighLight = PropBag.ReadProperty("HighLightStyle", 1)
  Set m_ItemsFont = PropBag.ReadProperty("ItemsFont", Ambient.Font)
  m_Vertical = PropBag.ReadProperty("Vertical", False)
  
  CleanUP
  For i = 1 To cnt
    Set o = New MenuItem
    o.Enabled = PropBag.ReadProperty("MenuEnabled" & i, -1)
    o.Checked = PropBag.ReadProperty("MenuChecked" & i, 0)
    o.Ident = PropBag.ReadProperty("MenuIdent" & i, 0)
    o.Caption = PropBag.ReadProperty("MenuCaption" & i, "")
    o.Name = PropBag.ReadProperty("MenuName" & i, "")
    o.UseMaskColor = PropBag.ReadProperty("MenuUseMaskColor" & i, -1)
    o.MaskColor = PropBag.ReadProperty("MenuMaskColor" & i, QBColor(13))
    Set o.APicture(ifwpPicNormal) = PropBag.ReadProperty("MenuPicture" & i, Nothing)
    o.Visible = PropBag.ReadProperty("MenuVisible" & i, -1)
    o.Description = PropBag.ReadProperty("MenuDescription" & i, "")
    MenuTree.Add o
  Next
  Set o = Nothing
  Refresh

End Sub

Private Sub UserControl_Resize()
  Dim w As Long, H As Long
  If Spacing = 0 Then Exit Sub
  Refresh
  Exit Sub
  w = Width / Screen.TwipsPerPixelX
  H = Height / Screen.TwipsPerPixelY
  DrawWidth = 1
  H = H - 2
  Line (0, H)-(w, H), QBColor(15)
  Line (0, H - 2)-(w, H - 2), QBColor(8)
  DrawWidth = 1
  
  
End Sub



Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 Dim i As Long
  
  PropBag.WriteProperty "Spacing", Spacing, 12
  PropBag.WriteProperty "Style", Style, 0
  PropBag.WriteProperty "HighLightStyle", m_HighLight, 1
  PropBag.WriteProperty "ItemCount", MenuTree.Count, 0
  PropBag.WriteProperty "ItemsFont", m_ItemsFont, Ambient.Font
  
  For i = 1 To MenuTree.Count
   With MenuTree(i)
    PropBag.WriteProperty "MenuEnabled" & i, .Enabled, -1
    PropBag.WriteProperty "MenuChecked" & i, .Checked, 0
    PropBag.WriteProperty "MenuIdent" & i, .Ident, 0
    PropBag.WriteProperty "MenuCaption" & i, .Caption, ""
    PropBag.WriteProperty "MenuName" & i, .Name, ""
    PropBag.WriteProperty "MenuUseMaskColor" & i, .UseMaskColor, -1
    PropBag.WriteProperty "MenuMaskColor" & i, .MaskColor, QBColor(13)
    PropBag.WriteProperty "MenuPicture" & i, .APicture(0), Nothing
    PropBag.WriteProperty "MenuVisible" & i, .Visible, -1
  '  PropBag.WriteProperty "MenuSeparator" & i, .Separator, 0
    PropBag.WriteProperty "MenuDescription" & i, .Description, ""
    
   End With
  Next

End Sub


