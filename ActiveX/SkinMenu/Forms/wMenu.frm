VERSION 5.00
Begin VB.Form wMenu 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2205
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   147
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   2460
   End
End
Attribute VB_Name = "wMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MenuTree      As Collection
Dim IndexRange       As New Collection
Public ParentForm    As Object
Public FilterIndex   As Long
Public LeftCorner    As Long
Public TopCorner     As Long
Public Style         As Long
Public MenuOpen      As Boolean
Public OwnerWidth    As Long
Public HighLight     As Long
Public Vertical      As Boolean

Dim wf               As wMenu
Dim IsLoading        As Boolean
Dim LastSel          As Long
Sub DrawOpenStepItem(x1 As Long, y1 As Long, x2 As Long, y2 As Long, FromColor As Long, ToColor As Long, Mode As Long)
                       
   Dim c() As Long, Y As Long, n As Long, X As Long
 '  Line (x1, y1)-(x2, y1), QBColor(0), BF
                       
 On Error Resume Next

   
   n = (x2 - x1) / 20
   ReDim c(20)
   
   CreateGradateColors FromColor, ToColor, c
   
   For X = x1 To x2 Step n
       Line (X, y1)-(X + n, y2), c(Y), BF
       Y = Y + 1
   Next

                  
                    Line (x1 + 1, y1 + 1)-(x2 - 1, y1 + 1), QBColor(15)
                    Line -(x2 - 1, y2 - 1), QBColor(8)
                    Line -(x1 + 1, y2 - 1), QBColor(8)
                    Line -(x1 + 1, y1 + 1), QBColor(15)
   
   Refresh


End Sub

Private Sub DrawPicture(Item As MenuItem, x1 As Long, y1 As Long, Width As Long, Height As Long, pic As StdPicture)
   Dim PE As New ascPaintEffects
   With Item
      If .Enabled = 0 Then
       PE.PaintDisabledPicture hdc, pic, x1, y1, Width, Height, 0, 0, .MaskColor
      Else
       If pic.Type = vbPicTypeIcon Then
        'DrawTransparentBitmap doesn't support icons
        PE.PaintStandardPicture hdc, pic, x1, y1, Width, Height, 0, 0
       Else
        If .UseMaskColor Then
          PE.PaintTransparentPicture hdc, pic, x1, y1, Width, Height, 0, 0, .MaskColor
        Else
          PE.PaintStandardPicture hdc, pic, x1, y1, Width, Height, 0, 0
        End If
       End If
      End If
   End With
   Set PE = Nothing
      
Exit Sub
                
                
                
                
                
                Dim cDc As Long, OldBM As Long
                cDc = CreateCompatibleDC(hdc)
                OldBM = SelectObject(cDc, pic.handle)
                   StretchBlt hdc, x1, y1, Width, Height, cDc, 0, 0, pic.Width / Screen.TwipsPerPixelX, pic.Height / Screen.TwipsPerPixelY, vbSrcCopy
                SelectObject cDc, OldBM
                DeleteDC cDc

End Sub

Private Sub DrawCheck(x1 As Long, y1 As Long, Width As Long, Height As Long, Color As Long)
        DrawWidth = 2
          
          Line (x1 + 4, y1 + Height)-(x1 + 6, y1 + Height + 4), Color
          Line -(x1 + Width, y1 + 4), Color
         
        DrawWidth = 1
         
End Sub

Sub DrawTriangle(x1 As Long, y1 As Long, x2 As Long, y2 As Long, Color As Long, Optional Effect3D As Long = 0)
    Dim hBrush As Long, hPrev As Long, h2 As Long, d As Long
    Dim Y As Long, X As Long, b As Long
    Dim pt(2) As POINTAPI
    
    d = 5
    
    Y = y1 + (y2 - y1) * 0.5
    X = x2 - d
    
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
   
   If Effect3D > 0 Then
        Line (pt(0).X + 1, pt(0).Y + 1)-(pt(1).X - 1, pt(1).Y - 1), QBColor(8)
        Line -(pt(2).X - 1, pt(2).Y + 1), QBColor(8)
        Line -(pt(0).X + 1, pt(0).Y + 1), QBColor(15)
   End If
   
End Sub

Public Sub FillItems()
    Dim X As Long, Y As Long, Idx As Long, w As Long, x1 As Long, xMax As Long, y1 As Long
    Dim j As Long, i As Long, u As Long
    Dim b As Long, bLeft As Long, By As Long, bRight As Long
    Dim H As Long, idn As Long
    
    If Style <> mnOpenStep Then
      b = 8      ' Bordo sup e inferiore
    End If
    If Style <> mnExperiment Then
      By = 6     ' Spazio verticale tra items
    Else
      By = 6
    End If
    bLeft = 30 ' Bordo sinistro
    bRight = 5 ' Bordo destro
    H = TextHeight("A")
    
    IsLoading = True
 '   FontName = "Tahoma"
 '   FontSize = 8
    Set IndexRange = New Collection
    
    idn = MenuTree(FilterIndex).Ident + 1
    Y = b * 0.5 'by
    For i = FilterIndex + 1 To MenuTree.Count
     With MenuTree(i)
        If .Ident < idn Then Exit For
        If .Ident = idn And .Visible Then
           IndexRange.Add i
           CurrentX = bLeft
           CurrentY = Y + 2
           .X = CurrentX
           .Y = CurrentY
            u = InStr(.Caption, "&")
            If u > 0 And Len(.Caption) < u Then
               .Accelerator = UCase(Mid(.Caption, u + 1, 1))
            End If
          ' Print .Caption
           y1 = Y + H + By
           x1 = bLeft + TextWidth(.Caption) + bRight
           .SetRect CurrentX, Y, x1, y1
           Y = y1
           If x1 > xMax Then xMax = x1
'           Set Img(j - 1).Picture = .APicture(0)
        End If
     End With
    Next
    
    x1 = bLeft
    Left = LeftCorner
    Top = TopCorner
    Height = (Y + b) * Screen.TwipsPerPixelY
    Width = (xMax + bLeft) * Screen.TwipsPerPixelX
    Cls
    
' allinea la larghezza degli items, tutti uguali

 '  If Style = mnExperiment Then
 '     SetRegions idn
 '  Else
    
    For i = FilterIndex + 1 To MenuTree.Count
      With MenuTree(i)
          If .Ident < idn Then Exit For
          If .Ident = idn And .Visible Then
           .GetRect X, Y, x1, y1
           .SetRect 3, Y, xMax + bLeft - 3, y1
          End If
      End With
    Next
 '  End If
    
   If Style = mnExperiment Then SetRegions idn
    
    IsLoading = False
    

End Sub


Private Function FindAccelerator(Key As String) As Long
  Dim i As Long, j As Long
  
  For i = 1 To IndexRange.Count
      j = IndexRange(i)
      If MenuTree(j).Accelerator = UCase(Key) Then
         FindAccelerator = j
         Exit Function
      End If
  Next

End Function


Sub ForceUnload()
  On Error Resume Next
  ParentForm.ForceUnload
  Unload Me
End Sub


Private Function IsInRange(Index As Long) As Boolean

 Dim i As Long
 
 For i = 1 To IndexRange.Count
     If IndexRange(i) = Index Then
        IsInRange = True
        Exit Function
     End If
 Next
 
End Function



Sub SendClickEvent(Key As String)
   ParentForm.SendClickEvent Key
End Sub

Sub SendDescriptionEvent(Dex As String)
  ParentForm.SendDescriptionEvent Dex
End Sub


Sub SetRegions(idn As Long)
   Dim Points(490) As PointXY
   Dim n As Long, i As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long
   n = -1
  
   hRgn& = CreateRectRgn(0, 0, 0, 0)
  
   For i = FilterIndex + 1 To MenuTree.Count
      With MenuTree(i)
          If .Ident < idn Then Exit For
          If .Ident = idn And .Visible Then
             .GetRect x1, y1, x2, y2
             x1 = x1 - 8 '+ 3
             x2 = x2 + 8 ''- 3
             y1 = y1 + 1
             y2 = y2 - 1
             
             n = n + 1
             Points(n).X = x1
             Points(n).Y = y1
             n = n + 1
             Points(n).X = x2
             Points(n).Y = y1
             n = n + 1
             Points(n).X = x2
             Points(n).Y = y2
             n = n + 1
             Points(n).X = x1
             Points(n).Y = y2
             n = n + 1
             Points(n).X = x1
             Points(n).Y = y1
             hPGRgn& = CreatePolygonRgn(Points(0), 5, 1)
             CombineRgn hRgn&, hRgn&, hPGRgn&, RGN_OR
             n = -1
          End If
      End With
    Next
  
   SetWindowRgn Me.hwnd, hRgn&, True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  'MsgBox shit & "  " & KeyCode
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim i As Long
  
  If Button = 2 Then
     Unload Me
     Exit Sub
  End If
  
  If Not wf Is Nothing Then Unload wf
  i = PointToItem(X, Y)
  If i = 0 Then Exit Sub
  If LastSel <> i Then
     If LastSel <> 0 Then RefreshItem LastSel, 0
     LastSel = i
  End If
  
  
  RefreshItem i, -1
  Refresh
  
If MenuTree(i).IsRoot Then
  
  If MenuOpen = 0 Then
     ShowMenu i
     DoEvents
     MenuOpen = True
  Else
     Unload wf
     MenuOpen = False
  End If

Else
  ' do click event
   If MenuTree(i).Enabled = False Or MenuTree(i).Caption = "-" Then Exit Sub
   SendClickEvent MenuTree(i).Name
   On Error Resume Next
   Unload wf
   ParentForm.ForceUnload
   Unload Me
End If
  
End Sub


Sub ShowMenu(Index As Long)
    Dim idn As Long, i As Long, P As POINTAPI
    Dim X As Long, Y As Long, w As Long, H As Long
    Dim j As Long
    
    On Error Resume Next
    Unload wf
    On Error GoTo 0
    
    MenuTree(Index).GetRect X, Y, w, H
    
    P.X = X + w
    P.Y = Y
    ClientToScreen hwnd, P
   
   ' ClientToScreen ContainerhWnd, P
    P.X = P.X * Screen.TwipsPerPixelX
    P.Y = P.Y * Screen.TwipsPerPixelY
    
    Set wf = New wMenu
    Load wf
    Set wf.MenuTree = MenuTree  ' passa la collection al form
    Set wf.ParentForm = Me      ' pasa il riferimento a se stesso
    wf.FilterIndex = Index
    wf.LeftCorner = P.X
    wf.TopCorner = P.Y
    wf.Style = Style
    wf.HighLight = HighLight
    wf.FillItems
    wf.Show
    
    MenuTree(Index).Expanded = True
    
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim i As Long
  Static s As Long
  i = PointToItem(X, Y)
  
  If LastSel <> i Then
     If i <> 0 Then
        SendDescriptionEvent MenuTree(i).Description
       ' ToolTipText = MenuTree(i).Description
     End If
     If LastSel <> 0 Then RefreshItem LastSel, 0
     LastSel = i
  End If
  If i = 0 Then Exit Sub
  RefreshItem i, -1

End Sub


Private Function PointToItem(X As Single, Y As Single) As Long
  Dim i As Long
  For i = 1 To MenuTree.Count
     If MenuTree(i).PointInside(X, Y) Then
        If IsInRange(i) Then
           PointToItem = i
           Exit Function
        End If
     End If
  Next
End Function

Private Sub RefreshItem(Index As Long, Mode As Long)
   Dim X As Long, Y As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long, x3 As Long
   Dim Dx As Long, Dy As Long, u As Long, Bx As Long, By As Long
   Dim b As Long, hColor As Long
   Dim P As StdPicture
   
   b = BackColor
   With MenuTree(Index)
        If Mode <> 0 And (.Enabled = False Or .Caption = "-") Then Exit Sub
        Set P = .APicture(0)
        
        .GetRect x1, y1, x2, y2
        
        Line (x1, y1)-(x2, y2), b, BF
        Select Case Mode
           Case 0 ' normal, no pressed
              If Style = mnWindowsXP Then
                x3 = 24
                Line (x1, y1)-(x3, y2), vbButtonFace, BF
              End If
              
              If Style = mnExperiment Then
                 Line (x1, y1)-(x2, y2), vbWhite, BF
                 Line (x1, y1 + 2)-(x2, y2 - 2), 0, B
                ' GradateFill Me, x1 - 3, y1, x2 + 3, y2, QBColor(7), QBColor(15), 3, 1
                  '  GradateFill Me, x1 - 3, y1, x2 + 3, y2, QBColor(1), QBColor(11), 2
              End If
              
              If Style = mnOpenStep Then
                   DrawOpenStepItem x1 - 3, y1, x2 + 3, y2, QBColor(15), QBColor(7), Mode
              End If
              
              If .Caption = "-" Then
                y1 = y1 + (y2 - y1) * 0.5
                If Style = mnWindowsXP Then
                     Line (x1 + 5, y1)-(x2 - 5, y1), QBColor(7)
                Else
                     Line (x1 + 5, y1 - 1)-(x2 - 5, y1 - 1), QBColor(8)
                     Line (x1 + 5, y1)-(x2 - 5, y1), QBColor(15)
                End If
                Exit Sub
              End If
              
              
              If .Enabled Then
                    CurrentX = .X
                    CurrentY = .Y
                    GoSub PrintCaption
                    If P Is Nothing Then
                    Else
                        DrawPicture MenuTree(Index), x1 + 2, y1 + 2, 16, 16, P
                    End If
                    If .IsRoot Then
                       If Style = mnOpenStep Then
                          DrawTriangle x1, y1, x2, y2, QBColor(7), 1 ' disegna il triangolo
                       Else
                          DrawTriangle x1, y1, x2, y2, QBColor(0) ' disegna il triangolo
                       End If
                    End If

              Else
                    If Style = mnWindowsXP Then
                            b = ForeColor
                            ForeColor = QBColor(7)
                                CurrentX = .X
                                CurrentY = .Y
                                GoSub PrintCaption
'                                Print .Caption
                            ForeColor = b
                    Else
                            b = ForeColor
                            ForeColor = QBColor(15)
                                CurrentX = .X + 2
                                CurrentY = .Y + 2
           '                     Print .Caption
                                GoSub PrintCaption
                           ForeColor = QBColor(7)
                                CurrentX = .X
                                CurrentY = .Y
                                GoSub PrintCaption
'                                Print .Caption
                            ForeColor = b
                    End If
             End If
             If .Checked Then DrawCheck x1 + 4, y1, 10, 10, QBColor(8)
           Case 1 ' up
              CurrentX = .X - 1
              CurrentY = .Y - 1
              GoSub PrintCaption
 '             Print .Caption
              Line (x1, y1)-(x2, y1), QBColor(15)
              Line -(x2, y2), QBColor(8)
              Line -(x1, y2), QBColor(8)
              Line -(x1, y1), QBColor(15)
           Case -1 ' Pressed
           
            hColor = 15
            Select Case HighLight
              Case 0
               ' GradateFill x1 + 1, y1 + 1, x2 - 1, y2 - 1, QBColor(15), RGB(100, 200, 200), 1
                 Line (x1 + 1, y1 + 1)-(x2 - 1, y2 - 1), vbInactiveTitleBar, B
                 hColor = 8
              Case 1:  Line (x1 + 1, y1 + 1)-(x2 - 1, y2 - 1), vbInactiveTitleBar, BF
              Case 2:  GradateFill Me, x1 + 1, y1 + 1, x2 - 1, y2 - 1, QBColor(8), QBColor(15)
              Case 3:  GradateFill Me, x1 + 1, y1 + 1, x2 - 1, y2 - 1, QBColor(1), QBColor(7)
              Case 4:  GradateFill Me, x1 + 1, y1 + 1, x2 - 1, y2 - 1, QBColor(8), RGB(210, 210, 210), 1
              Case 5
                GradateFill Me, x1 + 1, y1 + 1, x2 - 1, y2 - 1, QBColor(7), QBColor(15), 1
                hColor = 8
            End Select
            
              CurrentX = .X + 1
              CurrentY = .Y + 1
              ForeColor = QBColor(hColor)
              GoSub PrintCaption
'             Print .Caption
              ForeColor = QBColor(0)
             If Style = mnClassic Then
              Line (x1, y1)-(x2, y1), QBColor(8)
              Line -(x2, y2), QBColor(15)
              Line -(x1, y2), QBColor(15)
              Line -(x1, y1), QBColor(8)
             End If
              If P Is Nothing Then
              Else
                 If Not .Checked Then DrawPicture MenuTree(Index), x1 + 1, y1 + 1, 16, 16, P
            '    If Not .Checked Then DrawPicture x1 + 1, y1 + 1, 32, 32, .TemporaryPicture
              End If
       
              If .IsRoot Then DrawTriangle x1, y1, x2, y2, QBColor(8)        ' disegna il triangolo
              If .Checked Then
                DrawCheck x1 + 4, y1, 10, 10, QBColor(0)
                DrawCheck x1 + 2, y1 - 2, 10, 10, QBColor(hColor)
              End If
     End Select
   
   End With
   Refresh

Exit Sub

PrintCaption:
   Dim cp As String
   cp = MenuTree(Index).Caption
   u = InStr(cp, "&")
   If u = Len(cp) Then
      u = 0
      cp = Mid(cp, 1, Len(cp) - 1)
   End If
   If Len(cp) = 0 Then Return
   If u > 0 Then
      Dx = TextWidth(Mid$(cp, 1, u))
      Dy = TextHeight("X")
      x3 = TextWidth(Mid$(cp, u + 1, 1))
      Bx = CurrentX
      Dx = Bx + Dx - x3
      By = CurrentY + Dy - 1
      cp = Mid(cp, 1, u - 1) & Mid(cp, u + 1, 1000)
   End If
   
   Print cp

   If u > 0 Then
      Line (Dx, By)-(Dx + x3, By)
   End If

Return

End Sub

Sub Form_Resize()
  Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long
  Dim i As Long, idn As Long
  Static RunOnce As Boolean
  
  If IsLoading Then Exit Sub
  
   x2 = ScaleWidth - 1
   y2 = ScaleHeight - 1
  
   Select Case Style
     Case mnClassic
        BackColor = vbButtonFace
        x2 = ScaleWidth - 1
        y2 = ScaleHeight - 1
  
           Line (x1, y1)-(x2, y1), QBColor(15)
           Line -(x2, y2), QBColor(8)
           Line -(x1, y2), QBColor(8)
           Line -(x1, y1), QBColor(15)
        
     Case mnWindowsXP
        BackColor = vbWindowBackground
        Line (x1, y1)-(x2, y2), QBColor(8), B
        If OwnerWidth > 0 Then
'           Line (x1, y1)-(OwnerWidth, y1), QBColor(15)
        End If
        x2 = 24
        Line (x1 + 1, y1 + 1)-(x2, y2 - 1), vbButtonFace, BF
   End Select
  
  
 If Not RunOnce Then
   
   idn = MenuTree(FilterIndex).Ident + 1
   For i = FilterIndex + 1 To MenuTree.Count
      With MenuTree(i)
          If .Ident < idn Then Exit For
          If .Ident = idn And .Visible Then
           RefreshItem i, 0
          End If
      End With
    Next
    
    RunOnce = True
 End If
  
  
End Sub



Private Sub Form_Unload(Cancel As Integer)
  If Not wf Is Nothing Then Unload wf
 ' If ParentForm.Name <> "wMenu" Then
     ParentForm.MenuOpen = False
 ' End If
  MenuTree(FilterIndex).Expanded = False
  Set wMenu = Nothing
End Sub

Private Sub Timer1_Timer()
  Dim H As Long
  H = GetActiveWindow
  If H <> hwnd Then
     If wf Is Nothing Then
        If ParentForm.Name = "wMenu" Then
         ' Debug.Print " ParentForm Scaricato"
          ParentForm.ForceUnload
        End If
        Unload Me
     End If
  End If
End Sub


