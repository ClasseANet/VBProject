Attribute VB_Name = "DSR"
Option Explicit

'Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
'Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
'Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
'Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

'Constants that are used by the API
'Const WM_CLOSE = &H10
'Const INFINITE = &HFFFFFFFF
'Const SYNCHRONIZE = &H100000

'Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

'Private Const GWL_EXSTYLE = (-20)
'Private Const GWL_STYLE = (-16&)
'Private Const LWA_COLORKEY = &H1
'Private Const LWA_ALPHA = &H2
'Private Const ULW_COLORKEY = &H1
'Private Const ULW_ALPHA = &H2
'Private Const ULW_OPAQUE = &H4
'Private Const WS_EX_LAYERED = &H80000
'Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
'Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Private Const SBBOTTOM As Long = 7
'Private Const WMVSCROLL As Long = &H115
Public Function AcessopEspecial(oSys As Object, sModulo As String) As Boolean
   Dim bAcesso As Boolean
   Dim sAux As String
   
   With oSys.USER
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
            .Alias = oSys.xdb.Alias
            .Server = oSys.xdb.Server
            .dbName = oSys.xdb.dbName
            .UID = oSys.xdb.UID
            .PWD = oSys.xdb.PWD
             
            DoEvents
            Screen.MousePointer = vbDefault
                
            'gIDUSU = .IDUSU
            .Show vbModal
             
            If Trim(.IDUSU) <> "" Then
               Set MyUser = CriarObjeto("BANCO.TB_USUARIO")
               Set MyUser.xdb = oSys.xdb
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
'================================================
'================================================
'Public Function GetValueXmlNode(ByRef ObjNode As Object, strNode As String) As String
'Public Function SelecionarArquivo(CmD As Object, Optional cDialogTitle = "Find File", Optional cfilename = "", Optional cFilter = "*.*", Optional cFilterIndex = 1, Optional pFlags As Long)
'Public Sub ExibirMensagemI(pSys As Object, pMsg As String, Optional pPopUp As Object)
'================================================
'================================================
Public Function GetValueXmlNode(ByRef ObjNode As Object, strNode As String) As String
'Public Function GetValueXmlNode(ByRef ObjNode As IXMLDOMElement, strNode As String) As String
   If ObjNode.selectSingleNode(strNode) Is Nothing Then
      GetValueXmlNode = ""
   Else
      GetValueXmlNode = ObjNode.selectSingleNode(strNode).Text
   End If
End Function
Public Function SelecionarArquivo(Optional CmD As Object, Optional cDialogTitle = "Find File", Optional cfilename = "", Optional cFilter = "*.*", Optional cFilterIndex = 1, Optional pFlags As Long)
   Dim LenP%, LenF%
       
   On Error GoTo OpenError
   If CmD Is Nothing Then
      Set CmD = CriarObjeto("MSComDlg.CommonDialog", False)
   End If
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
Public Sub ExibirMensagemI(pSys As Object, pMsg As String, Optional pPopUp As Object)
    'On Error Resume Next
    Dim i         As Integer
    Dim y         As Integer
    Dim LastPane  As Integer
    Dim nWidth    As Integer '170
    Dim nHeight   As Integer '130
    Dim Popup     As Object 'XtremeSuiteControls.PopupControl
    Dim Popup0    As Object 'XtremeSuiteControls.PopupControl
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
                  If pSys.MDI.PopupControl(i).state = 2 Then
                     If pSys.MDI.PopupControl(i).Item(y).Caption = pMsg Then
                        Exit Sub
                     End If
                  End If
               Next
            End If
         Else
            Load pSys.MDI.PopupControl(i)
         End If
         If pSys.MDI.PopupControl(i).state = 0 Then Exit For
      Next
      nIndex = i
      If nIndex > pSys.MDI.PopupControl.Ubound Then Load pSys.MDI.PopupControl(nIndex)
            
      Set Popup = pSys.MDI.PopupControl(nIndex)
      If nIndex > 0 Then Set Popup0 = pSys.MDI.PopupControl(nIndex - 1)
      
      If Popup.state = 2 Then
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
         .VisualTheme = 5 'xtpPopupThemeCustom
         .SetSize nWidth, nHeight
         .Right = Screen.Width / Screen.TwipsPerPixelX
         .Bottom = (Screen.Height - 510) / Screen.TwipsPerPixelY
         .AllowMove = True
               
         '**** SetGreenTheme Popup ****
         '*****************************
         Dim Item As Object 'PopupControlItem
      
         .RemoveAllItems
         .Icons.RemoveAll
         Set Item = .AddItem(0, 0, nWidth, nHeight, "", RGB(30, 120, 30), RGB(255, 255, 255))
      
         Set Item = .AddItem(5, 25, nWidth - 5, nHeight - 5, "", RGB(70, 130, 70), RGB(255, 255, 255))
      
         If ExisteArquivo(pSys.ExePath & pSys.CODSIS & ".ico") Then
            Set Item = .AddItem(5, 30, 12, 47, "")
            Item.SetIcon LoadPicture(pSys.ExePath & pSys.CODSIS & ".ico").Handle, 1 'xtpPopupItemIconNormal
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

