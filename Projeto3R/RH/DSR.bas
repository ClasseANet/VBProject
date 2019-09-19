Attribute VB_Name = "DSR"
Option Explicit
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Const GWL_EXSTYLE = (-20)
Private Const GWL_STYLE = (-16&)
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const ULW_OPAQUE = &H4
Private Const WS_EX_LAYERED = &H80000
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
'Public Sub SaveParam(pCODPARAM As String, pVLPARAM As String, Optional pDSCPARAM, Optional pCODSIS)
Public Function AcessopEspecial(oSys As Object, sModulo As String) As Boolean
   Dim bAcesso As Boolean
   Dim sAux As String
   
   With oSys.User
      If .Acessos(sModulo) <> "" Then
         sAux = UCase(Decrypt2(.SenhaUsu))
         
         If UCase(InputBoxPassword("Entre com a Senha", "Acesso Especial")) = sAux Then
            bAcesso = True
         End If
      Else
         Dim MyUser As Object
         
         'Call ExibeSenha
         Dim Splash As Object
         Set Splash = CriarObjeto("CONEXAO.Splash")
         With Splash
            Set .Sys = oSys
            .DebugSys = False
            .CODSIS = oSys.CODSIS
            .Alias = oSys.XDb.Alias
            .Server = oSys.XDb.Server
            .dbName = oSys.XDb.dbName
            .UID = oSys.XDb.UID
            .PWD = oSys.XDb.PWD
             
            DoEvents
            Screen.MousePointer = vbDefault
                
            'gIDUSU = .IDUSU
            .Show vbModal
             
            If Trim(.IDUSU) <> "" Then
               Set MyUser = CriarObjeto("BANCO.TB_USUARIO")
               Set MyUser.XDb = oSys.XDb
               If MyUser.Pesquisar(Ch_IDUSU:=.IDUSU) Then
                  If MyUser.Acessos(sModulo) <> "" Then
                     bAcesso = True
                  End If
               End If
            End If
         End With
         '******************
         
      End If
   End With
   AcessopEspecial = bAcesso
   Set MyUser = Nothing
   Set Splash = Nothing
End Function
Public Function AddButtonBar(Controls As Object, _
                              Id As Long, Caption As String, _
                              Optional BeginGroup As Boolean = False, _
                              Optional ControlType As Integer = 1, _
                              Optional Category As String = "") As Object
'Public Function AddButtonBar(Controls As CommandBarControls, _
                              Id As Long, Caption As String, _
                              Optional BeginGroup As Boolean = False, _
                              Optional ControlType As XTPControlType = xtpControlButton, _
                              Optional Category As String = "") As CommandBarControl
   
   'Dim oMenuItem As CommandBarControl
   Dim oMenuItem As Object
    
   Set oMenuItem = Controls.Add(ControlType, Id, Caption)
   With oMenuItem
      .BeginGroup = BeginGroup
    
      .Category = Category
      .Parameter = SetTag(.Parameter, "CARREGADO", 0)
      .Parameter = SetTag(.Parameter, "MENUCHILD", "S")
   End With
   
   Set AddButtonBar = oMenuItem
End Function

Public Sub ExecuteScript(xConn As Object, pPathFile As String, Optional pTerminator As String = ";")
   Dim Sql As String
   Dim SqlAux As String
   Dim sStatus As String
   Dim sTerminator As String
   
   'Dim x As DS_BANCO
   
   If ExisteArquivo(pPathFile) Then
      Sql = ReadTextFile(pPathFile)
      Sql = Replace(Sql, Chr(239), "")
      Sql = Replace(Sql, Chr(187), "")
      Sql = Replace(Sql, Chr(191), "")
      
      While InStr(Sql, "/*") <> 0
          Sql = Mid(Sql, 1, InStr(Sql, "/*") - 1) & Mid(Sql, InStr(InStr(Sql, "/*"), Sql, "*/") + 2)
      Wend
      While InStr(Sql, "--") <> 0
         If InStr(InStr(Sql, "--"), Sql, Chr(13)) <> 0 Then
            Sql = Mid(Sql, 1, InStr(Sql, "--") - 1) & Mid(Sql, InStr(InStr(Sql, "--"), Sql, Chr(13)) + 2)
         Else
            Sql = Mid(Sql, 1, InStr(Sql, "--") - 1)
         End If
      Wend
      If InStr(Sql, "--") <> 0 Then
         Sql = Mid(Sql, 1, InStr(Sql, "--") - 1)
      End If
           
      While InStr(Sql, pTerminator)
         SqlAux = Mid(Sql, 1, InStr(Sql, pTerminator))
         If TypeName(xConn) = "DS_BANCO" Then
            If Not xConn.Executa(SqlAux) Then
               sStatus = sStatus & "Erro : " & SqlAux & vbNewLine
            End If
         Else
            xConn.Execute SqlAux
         End If
         Sql = Mid(Sql, InStr(Sql, pTerminator) + 1)
      Wend
      
      sTerminator = "GO" & Chr(13)
      While InStr(UCase(Sql), sTerminator)
         SqlAux = Mid(Sql, 1, InStr(UCase(Sql), sTerminator) - 3)
         If TypeName(xConn) = "DS_BANCO" Then
            If Not xConn.Executa(SqlAux) Then
               sStatus = sStatus & "Erro : " & SqlAux & vbNewLine
            End If
         Else
            xConn.Execute SqlAux
         End If
         Sql = Mid(Sql, InStr(UCase(Sql), sTerminator) + 4)
      Wend
   
   End If
   If Trim(sStatus) <> "" Then
      sStatus = Now() & vbNewLine & sStatus
      Call WriteIniFile(App.Path & "\" & "ExeScr.log", Right(pPathFile, InStr(StrReverse(pPathFile), "\") - 1), "STATUS", sStatus)
'      MsgBox sStatus
   End If
End Sub
'================================================
'================================================
Public Sub SetVisualTheme(pSys As Object, Optional pForm As Object)
    On Error Resume Next
    Dim Form As Form
    Dim Ctrl As Object
                        
   If pForm Is Nothing Then
      For Each Form In Forms
         For Each Ctrl In Form.Controls
            Ctrl.VisualTheme = pSys.MDI.CommandBars.VisualTheme
         Next
      Next
   Else
      For Each Ctrl In pForm.Controls
         Ctrl.VisualTheme = pSys.MDI.CommandBars.VisualTheme
      Next
   End If
End Sub
Public Sub SetRunTimeFormProperty(pForm As Form)
   Dim CurStyle As Long
   Dim NewStyle As Long


   CurStyle = GetWindowLong(pForm.hwnd, GWL_STYLE)
   NewStyle = SetWindowLong(pForm.hwnd, GWL_STYLE, CurStyle) ' Xor (WS_BORDER)) ' Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX))
   'Call SetWindowLong(pForm.hwnd, GWL_STYLE, GetWindowLong(pForm.hwnd, GWL_STYLE) Xor (WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX))
   'Call SetWindowPos(pForm.hwnd, 0&, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED)
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
'Public Function CriarButtonToolbar(pToolbar As Object, pType As XTPControlType, pId As Long, _
         Optional pCaption As String, Optional pCategory As String, Optional pStyle As Integer = 2, _
         Optional pBeginGroup As Boolean, Optional pIconId As Long, Optional pChecked As Boolean, _
         Optional pParameter) As Object
Public Function CriarButtonToolbar(pToolbar As Object, pType As Integer, pId As Long, _
         Optional pCaption As String, Optional pCategory As String, Optional pStyle As Integer = 2, _
         Optional pBeginGroup As Boolean, Optional pIconId As Long, Optional pChecked As Boolean, _
         Optional pParameter) As Object
   
   Dim oControl As Object 'CommandBarControl
      
   With pToolbar
      Set oControl = .Controls.Find(pType, pId)
      If oControl Is Nothing Then
         Set oControl = .Controls.Add(pType, pId, pCaption)
         With oControl
            .Category = pCategory
            .IconId = pIconId
            .Checked = pChecked
            .BeginGroup = pBeginGroup
            .Style = pStyle
            If Not IsMissing(pParameter) Then
               .Parameter = pParameter
            End If
         End With
      End If
   End With
   
   Set CriarButtonToolbar = oControl
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
Public Function eFeriado(xConn As Object, ByVal pData As String, Optional pBanco As String = "BANCO_3R", Optional pTabela As String = "TB_GFERIADO") As Boolean
   Dim Tb As Object
   Dim Sql As String
   
   pData = Format(pData, "dd/mm/yyyy")
   
   Set Tb = CriarObjeto(pBanco & "." & pTabela)
   With Tb
      Set .XDb = xConn
      If .Pesquisar(Ch_DATA:=pData) Then
         eFeriado = True
      Else
         Sql = " ESCOPO=1"
         Sql = Sql & " And Day(DATA)=" & Day(CDate(pData))
         Sql = Sql & " And Month(DATA)=" & Month(CDate(pData))
         If .Pesquisar(Ch_Where:=Sql) Then
            eFeriado = True
         End If
      End If
   End With
End Function


Public Function GetValueXmlNode(ByRef ObjNode As Object, strNode As String) As String
'Public Function GetValueXmlNode(ByRef ObjNode As IXMLDOMElement, strNode As String) As String
   If ObjNode.selectSingleNode(strNode) Is Nothing Then
      GetValueXmlNode = ""
   Else
      GetValueXmlNode = ObjNode.selectSingleNode(strNode).Text
   End If
End Function
Public Function SelecionarArquivo(CmD As Object, Optional cDialogTitle = "Find File", Optional cfilename = "", Optional cFilter = "*.*", Optional cFilterIndex = 1, Optional pFlags As Long)
   Dim LenP%, LenF%
       
   On Error GoTo OpenError
   
   With CmD
      .DialogTitle = cDialogTitle
      .FileName = cfilename
      .Filter = cFilter ' "Access Files (*.mdb)|*.mdb"
      .FilterIndex = cFilterIndex
      '.Tag = ""
      .CancelError = True
      If IsMissing(pFlags) Then
         .Flags = 4096  '(&H1000) 'FileOpenConstants.cdlOFNFileMustExist
      Else
         .Flags = pFlags
      End If
      .ShowOpen
      LenP% = Len(.FileName)
      LenF% = Len(.FileTitle)
      SelecionarArquivo = UCase(.FileName)
      .Tag = UCase(Mid(.FileName, 1, LenP% - LenF%))
   End With
   Exit Function
OpenError:
   Screen.MousePointer = vbDefault
   CmD.FileName = ""
   If Err = 3049 Then
     If MsgBox(Error & vbLf & vbLf & "Attempt to Repair it?", 4 + 48) = vbYes Then
   '      Resume AttemptRepair
     End If
   End If
   If Err = 438 Then 'Object doesn't support this property or method
      Resume Next
   End If
   If Err <> 32755 And Err <> 3049 Then   'check for common dialog cancelled
   '    ShowError
   End If
End Function
Public Function isTransparent(ByVal hwnd As Long) As Boolean
   On Error Resume Next
   Dim Msg As Long
   Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
   If (Msg And WS_EX_LAYERED) = WS_EX_LAYERED Then
     isTransparent = True
   Else
     isTransparent = False
   End If
   If Err Then
     isTransparent = False
   End If
End Function
Public Function MakeTransparent(ByVal hwnd As Long, Perc As Integer) As Long
   Dim Msg As Long
   On Error Resume Next
   If Perc < 0 Or Perc > 255 Then
     MakeTransparent = 1
   Else
     Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
     Msg = Msg Or WS_EX_LAYERED
     SetWindowLong hwnd, GWL_EXSTYLE, Msg
     SetLayeredWindowAttributes hwnd, 0, Perc, LWA_ALPHA
     MakeTransparent = 0
   End If
   If Err Then
     MakeTransparent = 2
   End If
End Function
Public Function MakeOpaque(ByVal hwnd As Long) As Long
   Dim Msg As Long
   On Error Resume Next
   Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
   Msg = Msg And Not WS_EX_LAYERED
   SetWindowLong hwnd, GWL_EXSTYLE, Msg
   SetLayeredWindowAttributes hwnd, 0, 0, LWA_ALPHA
   MakeOpaque = 0
   If Err Then
     MakeOpaque = 2
   End If
End Function

 Public Function Transparency(ByVal hwnd As Long, Optional ByVal Col As Long = vbBlack, Optional ByVal PcTransp As Byte = 255, Optional ByVal TrMode As Boolean = True) As Boolean
' Return : True if there is no error.
' hWnd   : hWnd of the window to make transparent
' Col : Color to make transparent if TrMode=False
' PcTransp  : 0 à 255 >> 0 = transparent  -:- 255 = Opaque
   Dim DisplayStyle As Long
   On Error GoTo Saida
   DisplayStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
   If DisplayStyle <> (DisplayStyle Or WS_EX_LAYERED) Then
      DisplayStyle = (DisplayStyle Or WS_EX_LAYERED)
      Call SetWindowLong(hwnd, GWL_EXSTYLE, DisplayStyle)
   End If
   Transparency = (SetLayeredWindowAttributes(hwnd, Col, PcTransp, IIf(TrMode, LWA_COLORKEY Or LWA_ALPHA, LWA_COLORKEY)) <> 0)
Saida:
    If Not Err.Number = 0 Then Err.Clear
End Function
Public Sub ActiveTransparency(pForm As Form, d As Boolean, F As Boolean, Perc As Integer, Optional Color As Long)
   Dim bResult As Boolean
   
   If d And F Then
   'Makes color (here the background color of the shape) transparent
   'upon value of T_Transparency
       bResult = Transparency(pForm.hwnd, Color, Perc, False)
   ElseIf d Then
       'Makes form, including all components, transparent
       'upon value of T_Transparency
       bResult = Transparency(pForm.hwnd, 0, Perc, True)
   Else
       'Restores the form opaque.
       bResult = Transparency(pForm.hwnd, , 255, True)
   End If
End Sub
Public Sub ExibirMensagemI(pSys As Object, pMsg As String, Optional pPopUp As Object)
    'On Error Resume Next
    Dim i         As Integer
    Dim y         As Integer
    Dim LastPane  As Integer
    Dim nWidth    As Integer '170
    Dim nHeight   As Integer '130
    Dim Popup     As XtremeSuiteControls.PopupControl
    Dim Popup0    As XtremeSuiteControls.PopupControl
    Dim nIndex    As Integer
    
    
    nWidth = 220
    nHeight = 90
    Set Popup = Nothing
    Set Popup0 = Nothing
    If pPopUp Is Nothing Then
      Set Popup = pSys.MDI.PopupControl(0)
      For i = pSys.MDI.PopupControl.lbound To pSys.MDI.PopupControl.Ubound
         If ExisteIndex(pSys.MDI.PopupControl, i) Then
            If pSys.MDI.PopupControl(i).ItemCount > 0 Then
               For y = 0 To pSys.MDI.PopupControl(i).ItemCount - 1
                  If pSys.MDI.PopupControl(i).State = 2 Then
                     If pSys.MDI.PopupControl(i).Item(y).Caption = pMsg Then
                        Exit Sub
                     End If
                  End If
               Next
            End If
         Else
            Load pSys.MDI.PopupControl(i)
         End If
         If pSys.MDI.PopupControl(i).State = 0 Then Exit For
      Next
      nIndex = i
      If nIndex > pSys.MDI.PopupControl.Ubound Then Load pSys.MDI.PopupControl(nIndex)
            
      Set Popup = pSys.MDI.PopupControl(nIndex)
      If nIndex > 0 Then Set Popup0 = pSys.MDI.PopupControl(nIndex - 1)
      
      If Popup.State = 2 Then
         nIndex = pSys.MDI.PopupControl.Ubound + 1
         Load pSys.MDI.PopupControl(nIndex)
         Set Popup = pSys.MDI.PopupControl(nIndex)
         Set Popup0 = pSys.MDI.PopupControl(nIndex - 1)
      End If
    Else
      Set Popup = pPopUp
    End If
    With Popup
      'lastPane = IIf(chkMultiplePopup, ID_POPUP2, ID_POPUP0)
      LastPane = 0
    
      For i = 0 To LastPane
         .Animation = 2
         .AnimateDelay = 256
         .ShowDelay = 0
         .Transparency = 240
         .VisualTheme = xtpPopupThemeCustom
         .SetSize nWidth, nHeight
         .Right = Screen.Width / Screen.TwipsPerPixelX
         .Bottom = (Screen.Height - 510) / Screen.TwipsPerPixelY
         .AllowMove = True
               
         '**** SetGreenTheme Popup ****
         '*****************************
         Dim Item As PopupControlItem
      
         .RemoveAllItems
         .Icons.removeAll
         Set Item = .AddItem(0, 0, nWidth, nHeight, "", RGB(30, 120, 30), RGB(255, 255, 255))
      
         Set Item = .AddItem(5, 25, nWidth - 5, nHeight - 5, "", RGB(70, 130, 70), RGB(255, 255, 255))
      
         If ExisteArquivo(pSys.ExePath & pSys.CODSIS & ".ico") Then
            Set Item = .AddItem(5, 30, 12, 47, "")
            Item.SetIcon LoadPicture(pSys.ExePath & pSys.CODSIS & ".ico").Handle, xtpPopupItemIconNormal
         End If
         Set Item = .AddItem(50, 30, nWidth - 10, nHeight - 20, pMsg)
         Item.TextAlignment = DT_WORDBREAK Or DT_LEFT Or DT_VCENTER
         Item.TextColor = RGB(255, 255, 0)
         Item.CalculateHeight
         Item.Hyperlink = False
         Item.Id = 0
      
         'Set Item = Popup.AddItem(104, 27, nWidth, 45, "more...")
         
         Set Item = .AddItem(5, 0, nWidth, 25, " Mensagem do Sistema")
         Item.TextAlignment = DT_SINGLELINE Or DT_LEFT Or DT_VCENTER
         Item.TextColor = RGB(255, 255, 255)
         Item.Bold = True
         Item.Hyperlink = False
      
         If ExisteArquivo(pSys.ExePath & "Close.bmp") Then
            Set Item = .AddItem(nWidth - 20, 6, nWidth - 6, 19, "")
            Item.SetIcons LoadPicture(pSys.ExePath & "Close.bmp").Handle, 0, xtpPopupItemIconNormal Or xtpPopupItemIconSelected Or xtpPopupItemIconPressed
         End If
         Item.Id = 1
      
         If nIndex > 0 And Not Popup0 Is Nothing Then
            .Right = Popup0.Right
            .Bottom = Popup0.Bottom - Popup0.Height
            .AnimateDelay = 256 'Popup0.AnimateDelay + 256
            .ShowDelay = 0 'Popup0.ShowDelay + 1000
            .Show
         End If
         .Show
      Next
   End With
End Sub
Public Function ExisteIndex(pObj As Object, pIndex As Integer) As Boolean
   Dim Value
   On Error Resume Next
   
   Value = pObj(pIndex)
   If Err = 449 Then
      Err = 0
      Set Value = pObj(pIndex)
   End If
   
   ExisteIndex = (Err = 0)
End Function
Public Sub LimparTela(frm As Object)
   Dim i As Integer
   On Error Resume Next
   
   For i% = 0 To frm.Controls.Count - 1
      
      Select Case UCase(TypeName(frm.Controls(i)))
         
         Case "TEXTBOX"
            frm.Controls(i) = ""
         
         Case "MASKEDBOX"
            Dim sMask As String
            sMask = frm.Controls(i).Mask
            frm.Controls(i).Mask = ""
            frm.Controls(i).Text = ""
            frm.Controls(i).Mask = sMask
         
         Case "LABEL"
            If frm.Controls(i).BorderStyle = 1 Then
               frm.Controls(i) = ""
            End If
         
         Case "OPTIONBUTTON":
            If frm.Controls(i).Index = 0 Then
               frm.Controls(i).Value = True
            End If
         
         Case "COMBOBOX"
           If frm.Controls(i).ListCount > 0 Then
              frm.Controls(i).ListIndex = 0
           Else
              frm.Controls(i).ListIndex = -1
           End If
           
         Case "CHECKBOX"
            frm.Controls(i).Value = 0
      End Select
   Next
End Sub

