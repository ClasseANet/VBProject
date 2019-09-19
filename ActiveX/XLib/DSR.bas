Attribute VB_Name = "DSR"
Option Explicit
Private Const GWL_STYLE As Long = (-16&)
Private Const WS_BORDER As Long = &H800000
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const WS_CAPTION As Long = &HC00000

Private Const SWP_NOSENDCHANGING = &H400
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOCOPYBITS = &H100
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Sub SetRunTimeFormProperty(pForm As Form)
   Dim CurStyle As Long
   Dim NewStyle As Long

   CurStyle = GetWindowLong(pForm.hwnd, GWL_STYLE)
   NewStyle = SetWindowLong(pForm.hwnd, GWL_STYLE, CurStyle) ' Xor (WS_BORDER)) ' Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX))
   'Call SetWindowLong(pForm.hwnd, GWL_STYLE, GetWindowLong(pForm.hwnd, GWL_STYLE) Xor (WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX))
   'Call SetWindowPos(pForm.hwnd, 0&, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED)
End Sub
Public Sub AcoplarForm(pForm As Form, nPane As Integer, pSys As Object, Optional bDefineFoco As Boolean = True, Optional pMDI As Object)
   With pForm
      .BorderStyle = vbBSNone
      .ClipControls = False
      .WindowState = vbMaximized
      '.MaxButton = False
      '.MinButton = False
      '.ShowInTaskbar = False
   End With
   If IsMissing(pMDI) Or pMDI Is Nothing Then
      With pSys.MDI.DockingPaneManager
         If .Panes(nPane).Handle <> pForm.hwnd Then
            Call SetMDI(pForm.hwnd, pSys.MDI.hwnd)
            .Panes(nPane).Handle = pForm.hwnd
         End If
      End With
   Else
      With pMDI.DockingPaneManager
         If .Panes(nPane).Handle <> pForm.hwnd Then
            'Call SetMDI(pForm.hWnd, pSys.MDI.hWnd)
            .Panes(nPane).Handle = pForm.hwnd
         End If
      End With
   End If
   
   '* Definir foco
   With pForm
      If bDefineFoco Then
         On Error Resume Next
         Dim i As Integer
         Dim iTab As Integer
         Dim bAchou As Boolean
         
         For iTab = 0 To .Controls.Count - 1
            bAchou = False
            For i = 0 To .Controls.Count - 1
               If .Controls(i).TabIndex = iTab Then
                  bAchou = .Controls(i).Visible
                  bAchou = bAchou And .Controls(i).Enabled
                  bAchou = bAchou And (Err = 0)
                  If bAchou Then .Controls(i).SetFocus
                  bAchou = bAchou And (Err = 0)
                  If bAchou Then
                     .Controls(i).SetFocus
                     iTab = .Controls.Count
                  End If
                  Exit For
               End If
               Err = 0
            Next
         Next
      End If
   End With
End Sub
Public Function CriarToolbar(pSys As Object, pNmToolBar As String) As Object
   Dim oToolBar As Object 'CommandBars
   Dim oBar     As Object 'CommandBar
   Dim n        As Object 'CommandBar
   
   Set oToolBar = pSys.MDI.CommandBars
   With pSys
      '* Verificar se Toolbar Existe
      For Each n In oToolBar
         If n.Title = pNmToolBar Then
            Set oBar = n
            Exit For
         End If
      Next
      
      '* Se Toolbar não Existe então cria
      If oBar Is Nothing Then
         Set oBar = oToolBar.Add(pNmToolBar, 4)  ' 0=xtpBarTop, 4=xtpBarFloating
         oBar.Visible = False
      End If
   End With
   Set CriarToolbar = oBar
End Function
Public Function FormatarData(pStrDate As String) As String
   Dim sText As String
   Dim nPos  As Integer
   Dim nPos2  As Integer

   sText = pStrDate
   
   If Trim(sText) = "" Then sText = Format(Now(), "dd/mm/yyyy")
   
   nPos = InStr(sText, "/")
   nPos2 = InStr(nPos + 1, sText, "/")
   If nPos > 0 Then
      If nPos <= 2 Then
         sText = StrZero(Mid(sText, 1, nPos - 1), 2) & Mid(sText, nPos)
         nPos = InStr(sText, "/")
         nPos2 = InStr(nPos + 1, sText, "/")
      End If
      If nPos2 = 0 Then
         sText = sText & Format(Now(), "/yyyy")
         nPos = InStr(sText, "/")
         nPos2 = InStr(nPos + 1, sText, "/")
      End If
      If nPos2 - nPos <= 2 And nPos2 - nPos > 0 Then
         sText = Mid(sText, 1, nPos) & StrZero(Mid(sText, nPos + 1, nPos2 - nPos - 1), 2) & Mid(sText, nPos2)
      End If
   End If
   
   sText = Replace(sText, "/", "")
   If Len(sText) <= 2 Then
      sText = StrZero(sText, 2) + Format(Now(), "/mm/yyyy")
   ElseIf Len(sText) = 3 Then
      If Val(Mid(sText, 2, 2)) > 12 Then
         sText = Mid(sText, 1, 2) + "/" + StrZero(Mid(sText, 3, 2), 2) + Format(Now(), "/yyyy")
      Else
         sText = StrZero(Mid(sText, 1, 1), 2) + "/" + StrZero(Mid(sText, 2, 2), 2) + Format(Now(), "/yyyy")
      End If
   ElseIf Len(sText) = 4 Then
      sText = Mid(sText, 1, 2) + "/" + StrZero(Mid(sText, 3, 2), 2) + Format(Now(), "/yyyy")
   ElseIf Len(sText) = 5 Then
      sText = Mid(sText, 1, 2) + "/" + StrZero(Mid(sText, 3, 2), 2) + Mid(Format(Now(), "/yyyy"), 1, 3) & StrZero(Mid(sText, 5), 2)
   ElseIf Len(sText) > 5 Then
      If Mid(sText, 5, 4) >= Left(Year(Now), Len(Mid(sText, 5, 4))) Then
         sText = Mid(sText, 1, 2) + "/" + Mid(sText, 3, 2) + "/" + Left(Left(Year(Now) - 100, 4 - Len(Mid(sText, 5, 4))) + Mid(sText, 5, 4), 4)
      Else
         sText = Mid(sText, 1, 2) + "/" + Mid(sText, 3, 2) + "/" + Left(Left(Year(Now), 4 - Len(Mid(sText, 5, 4))) + Mid(sText, 5, 4), 4)
      End If
   End If
   If IsDate(sText) Then
      FormatarData = sText
   End If
End Function
 Public Function FormatarHora(pStrHour As String, Optional pSegundo As Boolean = False) As String
   Dim sHora As String
   
   sHora = Replace(pStrHour, ":", "")
   If Len(sHora) <= 2 Then
      If Val(sHora) < 23 Then
         sHora = StrZero(sHora, 2) & ":00"
      ElseIf Val(sHora) >= 24 And Val(sHora) < 60 Then
         sHora = "00:" & StrZero(sHora, 2)
      ElseIf Val(sHora) >= 60 Then
         sHora = StrZero(Mid(sHora, 1, 1), 2) & ":" & StrZero(Mid(sHora, 1, 1), 2)
      End If
   ElseIf Len(sHora) = 3 Then
      sHora = StrZero(Mid(sHora, 1, 1), 2) & ":" & Mid(sHora, 2, 2)
   ElseIf Len(sHora) = 4 Then
      sHora = Mid(sHora, 1, 2) & ":" & Mid(sHora, 3, 2)
   End If
   If pSegundo Then
      If Len(Replace(pStrHour, ":", "")) = 5 Then
         sHora = sHora & ":" & StrZero(Mid(Replace(pStrHour, ":", ""), 5, 1), 2)
      ElseIf Len(Replace(pStrHour, ":", "")) = 6 Then
         sHora = sHora & ":" & Mid(Replace(pStrHour, ":", ""), 5, 2)
      Else
         sHora = sHora & ":00"
      End If
   End If
   
   If IsDate(sHora) Then
      FormatarHora = sHora
   End If
End Function
Function SendSMS(ByVal pUrl As String) As String
   Dim Status As String
   Dim oWHttp As Object
        
   'Set OWHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
   If oWHttp Is Nothing Then Set oWHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
   If oWHttp Is Nothing Then Set oWHttp = CreateObject("WinHttp.WinHttpRequest")
   If oWHttp Is Nothing Then Set oWHttp = CreateObject("MSXML2.ServerXMLHTTP")
   If oWHttp Is Nothing Then Set oWHttp = CreateObject("Microsoft.XMLHTTP")
       
   With oWHttp
      .Open "GET", pUrl
      .Send
    
      Status = .ResponseText
   End With
   SendSMS = Status
End Function

' Função auxiliar para codificar a mensagem para formato URL
Function UrlEncode(ByVal urlText As String) As String
    Dim i As Long
    Dim ansi() As Byte
    Dim ascii As Integer
    Dim encText As String

    ansi = StrConv(urlText, vbFromUnicode)

    encText = ""
    For i = 0 To UBound(ansi)
        ascii = ansi(i)

        Select Case ascii
        Case 48 To 57, 65 To 90, 97 To 122
            encText = encText & Chr(ascii)
        Case 32
            encText = encText & "+"
        Case Else
            If ascii < 16 Then
                encText = encText & "%0" & Hex(ascii)
            Else
                encText = encText & "%" & Hex(ascii)
            End If
        End Select
    Next i
    
    UrlEncode = encText
End Function
Public Function ReadTextFile(strPath As String) As String
    On Error GoTo ErrTrap
    Dim intFileNumber As Integer
    
    If Dir(strPath) = "" Then Exit Function
    intFileNumber = FreeFile
    Open strPath For Input As #intFileNumber
    
    ReadTextFile = Input(LOF(intFileNumber), #intFileNumber)
ErrTrap:
    Close #intFileNumber
End Function
Public Sub MontarToolbarDinamico(ByRef pMDI As Object)
'   Dim Control As CommandBarControl
'   Dim ToolBar As CommandBar
'   Dim TBCmdBar As Object
'
'   Dim Sql        As String
'   Dim MyRs       As Object
'   Dim nOrdAntes  As Integer
   
'   Set ToolBar = pMDI.CommandBars.Add("Standard", xtpBarTop)
'
'   Sql = "Select * "
'   Sql = Sql & " From GBARCMD"
'   Sql = Sql & " Where CODSIS = " & SqlStr(Sys.CODSIS)
'   Sql = Sql & " Order By GRUPO, ORDEM, ID"
'
'   nOrdAntes = 0
'   If Sys.xDb.AbreTabela(Sql, MyRs) Then
'      While Not MyRs.EOF
'         Set Control = ToolBar.Controls.Add(XTPControlType.xtpControlButton, MyRs("ID"), MyRs("DSCMODU"))
'         If MyRs.AbsolutePosition > 1 Then
'            Control.BeginGroup = (nOrdAntes = xVal(MyRs("ORDEM")))
'         End If
'         Control.Style = IIf(IsNull(MyRs("IMAGEM")), xtpButtonCaption, xtpButtonIcon)
'         Control.Style = xtpButtonIcon
'
'
'         nOrdAntes = xVal(MyRs("ORDEM"))
'         MyRs.MoveNext
'      Wend
'   End If

End Sub
Public Sub DockBarRightOf(pBarToDock As Variant, pBarOnLeft As Variant, Optional pSys As Object)
   Dim ToolBar As Object 'CommandBars
   Dim nBar    As Object
   Dim Left    As Long
   Dim Top     As Long
   Dim Right   As Long
   Dim Bottom  As Long
   
   Dim BarToDock As Object 'CommandBar
   Dim BarOnLeft As Object 'CommandBar
    
   If Not IsEmpty(pSys) And Not pSys Is Nothing Then
      Set ToolBar = pSys.MDI.CommandBars
   End If
      
   If InArray(TypeName(pBarToDock), Array("IMenuBar", "ICommandBar")) Then
      Set BarToDock = pBarToDock
   ElseIf TypeName(pBarOnLeft) = "String" Then
      For Each nBar In ToolBar
         If nBar.Title = pBarToDock Then
            Set BarToDock = nBar
            Exit For
         End If
      Next
   ElseIf TypeName(pBarOnLeft) = "Integer" Then
      Set BarToDock = ToolBar(pBarToDock)
   End If
   
   Set ToolBar = BarToDock.CommandBars
   
   If InArray(TypeName(pBarOnLeft), Array("IMenuBar", "ICommandBar")) Then
      Set BarOnLeft = pBarOnLeft
   ElseIf TypeName(pBarOnLeft) = "String" Then
      For Each nBar In ToolBar
         If nBar.Title = pBarOnLeft Then
            Set BarOnLeft = nBar
            Exit For
         End If
      Next
   ElseIf TypeName(pBarOnLeft) = "Integer" Then
      Set BarOnLeft = ToolBar(pBarOnLeft)
   End If
    
   If BarToDock Is Nothing Then Exit Sub
   If BarOnLeft Is Nothing Then Exit Sub
       
    ToolBar.RecalcLayout
        
    BarOnLeft.GetWindowRect Left, Top, Right, Bottom
    ToolBar.DockToolBar BarToDock, Right, (Bottom + Top) / 2, BarOnLeft.Position
End Sub
'Public Function AddReportRecord(Control As ReportControl, Parent As ReportRecord, Columns As Variant, Optional Icon, Optional HasCheckbox, Optional TreeColumn As Integer = 0, Optional GroupCaption) As ReportRecord
Public Function AddReportRecord(Control As Object, Parent As Object, Columns As Variant, Optional Icon, Optional HasCheckbox, Optional TreeColumn As Integer = 0, Optional GroupCaption) As Object
   Dim xRecord As Object 'ReportRecord
   Dim xItem   As Object 'ReportRecordItem
   Dim i       As Integer
   
   
   If Parent Is Nothing Then
      Set xRecord = Control.Records.Add
   Else
      Set xRecord = Parent.Childs.Add()
      Control.Columns(TreeColumn).TreeColumn = True
   End If
   
   Set xItem = xRecord.AddItem(Columns(0))
   If Not IsMissing(Icon) Then xItem.Icon = Icon
   If Not IsMissing(HasCheckbox) Then xItem.HasCheckbox = HasCheckbox
   If Not IsMissing(GroupCaption) Then xItem.GroupCaption = GroupCaption
   
   
   For i = 1 To Control.Columns.Count - 1 'UBound(Columns)
      If i <= UBound(Columns) Then
         xRecord.AddItem Columns(i)
      Else
         xRecord.AddItem ""
      End If
   Next
   'Set Item = Record.AddItem(Price)
   'Item.Format = "$ %s"
   
   Set AddReportRecord = xRecord
End Function

