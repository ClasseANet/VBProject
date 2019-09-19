VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_VerifVersao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarVERIF  As FrmVerifVersao
Attribute mvarVERIF.VB_VarHelpID = -1
Private WithEvents mvarFTP    As FrmCAFTP
Attribute mvarFTP.VB_VarHelpID = -1
Private WithEvents xInet      As Inet
Attribute xInet.VB_VarHelpID = -1

Dim mvarMe                 As Object
Dim PrgBar                 As Object
Dim sFormLog               As String
'Dim mvarArqCtrl               As String
Dim bLiberado              As Boolean
Dim bDsvm                  As Boolean
Dim gCollArq               As Collection

Private mvarSys            As Object
Private mvarArqCtrl        As String '= "Versao.txt"
Private mvarRemotePath     As String '= "/private/Cliente/Comum"
Private mvarLocalPath      As String '= MySys.PathSetup
Private mvarLocalPathRPT   As String '= MySys.PathRpt
Private mvarisSystem       As Boolean
Private mvarCODSIS         As String
Private mvarIDUSU          As String
Private mvarFtpIP          As String
Private mvarFtpUsu         As String
Private mvarFTPPwd         As String
Private mvarConectado      As Boolean
Private mvarResult         As Boolean
Private mvarStatus         As String
Private mvarState          As InetCtlsObjects.StateConstants
Private mvarAutoInstall    As Boolean
Private mvarAutoStart      As Boolean
Private mvarFileList       As String
Private mvarFileIni        As String
Public Property Let FileList(ByVal vData As String)
   mvarFileList = vData
   
End Property
Public Property Get FileList() As String
   FileList = mvarFileList
End Property
Public Property Let AutoStart(ByVal vData As Boolean)
   mvarAutoStart = vData
End Property
Public Property Get AutoStart() As Boolean
   AutoStart = mvarAutoStart
End Property
Public Property Let AutoInstall(ByVal vData As Boolean)
   mvarAutoInstall = vData
End Property
Public Property Get AutoInstall() As Boolean
   AutoInstall = mvarAutoInstall
End Property
Public Property Let ArqCtrl(ByVal vData As String)
   mvarArqCtrl = vData
End Property
Public Property Get ArqCtrl() As String
   ArqCtrl = mvarArqCtrl
End Property
Public Property Get State() As InetCtlsObjects.StateConstants
   State = mvarState
End Property
Public Property Get Status() As String
   Status = mvarStatus
End Property
Public Property Get Result() As Boolean
   Result = mvarResult
End Property
Public Property Get Conectado() As Boolean
   Conectado = mvarConectado
End Property
Public Property Let RemotePath(ByVal vData As String)
   mvarRemotePath = vData
End Property
Public Property Get RemotePath() As String
   RemotePath = mvarRemotePath
End Property
Public Property Let LocalPath(ByVal vData As String)
   mvarLocalPath = vData
   mvarFileIni = mvarLocalPath & "CAVERSAO.ini"
   If Not ExisteArquivo(mvarFileIni) Then
      mvarLocalPath = App.Path & "\"
      mvarFileIni = mvarLocalPath & "CAVERSAO.ini"
      If Not ExisteArquivo(mvarFileIni) Then
         mvarLocalPath = Environ("programfiles") & "\ClasseA\Admin\Dll\"
         mvarFileIni = mvarLocalPath & "CAVERSAO.ini"
      End If
   End If
End Property
Public Property Get LocalPath() As String
   LocalPath = ResolvePathName(mvarLocalPath)
End Property
Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
   CODSIS = mvarSys.CODSIS
End Property
Public Property Get Sys() As Object
   Set Sys = mvarSys
End Property
Public Property Let LocalOrig(ByVal vData As String)
   mvarRemotePath = vData
End Property
Public Property Get LocalOrig() As String
   LocalOrig = mvarRemotePath
End Property
Public Property Let LocalDest(ByVal vData As String)
   LocalPath = vData
End Property
Public Property Get LocalDest() As String
   LocalDest = mvarLocalPath
End Property
Public Property Let isSystem(ByVal vData As Boolean)
   mvarisSystem = vData
End Property
Public Property Get isSystem() As Boolean
   isSystem = mvarisSystem
End Property
Public Property Let CODSIS(ByVal vData As String)
   mvarCODSIS = vData
   Call DefineVarSis
End Property
Public Property Get CODSIS() As String
   CODSIS = mvarCODSIS
End Property
Public Property Let IDUSU(ByVal vData As String)
   mvarIDUSU = vData
End Property
Public Property Get IDUSU() As String
   IDUSU = mvarIDUSU
End Property
Public Property Let FtpIP(ByVal vData As String)
   mvarFtpIP = vData
End Property
Public Property Get FtpIP() As String
   FtpIP = mvarFtpIP
End Property
Public Property Let FtpUsu(ByVal vData As String)
   mvarFtpUsu = vData
End Property
Public Property Get FtpUsu() As String
   FtpUsu = mvarFtpUsu
End Property
Public Property Let FTPPwd(ByVal vData As String)
   mvarFTPPwd = vData
End Property
Public Property Get FTPPwd() As String
   FTPPwd = mvarFTPPwd
End Property
Private Sub Class_Initialize()
   bDsvm = (UCase(Mid(App.Path, 1, Len("C\:SISTEMAS"))) = "C:\SISTEMAS")
   mvarRemotePath = "/private/Cliente/Comum"
   If Trim(mvarLocalPath) = "" Then LocalPath = "C:\TMP\"
   mvarLocalPathRPT = mvarLocalPath
   mvarFtpIP = "ftp.classeanet.com.br"
   'mvarFtpIP = "ftp.classeaconsultoria.com.br"
   mvarFtpUsu = "classeanet"
   mvarFTPPwd = "d0lphin72" 'ramos10"
   If ExisteArquivo(mvarFileIni) Then
      Dim sFtpIP As String
      Dim sFtpUsu As String
      Dim sFTPPwd As String
      
      sFtpIP = ReadIniFile(mvarFileIni, "Config", "Host", "")
      sFtpUsu = ReadIniFile(mvarFileIni, "Config", "UID", "")
      sFTPPwd = Decrypt2(ReadIniFile(mvarFileIni, "Config", "PWD", ""))
   
      mvarFtpIP = IIf(sFtpIP = "", mvarFtpIP, sFtpIP)
      mvarFtpUsu = IIf(sFtpUsu = "", mvarFtpUsu, sFtpUsu)
      mvarFTPPwd = IIf(sFTPPwd = "", mvarFTPPwd, sFTPPwd)
   End If
   sFormLog = ""
   mvarArqCtrl = "VERSAO.TXT"
   mvarAutoInstall = True
   mvarFileList = App.Path & "\" & "Files.txt"
End Sub
Public Sub Show()
   ShowCAVs
End Sub
Public Sub ShowCAVs()
   Set mvarVERIF = New FrmVerifVersao
   mvarVERIF.Show vbModal
End Sub
Public Sub ShowFTP()
   Set mvarFTP = New FrmCAFTP
   mvarFTP.gCODSIS = mvarCODSIS
   mvarFTP.txtPassword.Text = mvarFTPPwd
   mvarFTP.Show vbModal
End Sub
Private Sub Class_Terminate()
   On Error Resume Next
   Unload mvarVERIF
   Unload mvarFTP
   
   xInet.Cancel
   Set xInet = Nothing
   
   Set mvarVERIF = Nothing
   Set mvarFTP = Nothing
   Set PrgBar = Nothing
   Set mvarSys = Nothing
End Sub
Private Sub mvarFTP_Activate()
   On Error Resume Next
   Set mvarMe = mvarFTP
   Set xInet = mvarFTP.InetFTP
   
   'mvarFTP.dirList.Path = GetSpecialFolder(38) & "ClasseA\UpLoad\"
   mvarFTP.dirList.Path = "c:\Tmp\UpLoad\"
End Sub

Private Sub mvarFTP_CmdAddListClick()
   Dim i As Integer
   Dim k As Integer
   Dim sFileIni As String
   Dim sFile As String
   
   For k = 0 To mvarFTP.filList.ListCount - 1
      If mvarFTP.filList.Selected(k) = True Then
         sFile = mvarFTP.dirList.Path & "\" & mvarFTP.filList.List(k)
   
         i = 1
         sFileIni = ReadIniFile(mvarFileList, "DPIL", "FILE" & i, "")
         While sFileIni <> ""
            If sFileIni = sFile Then
               sFileIni = ""
            Else
               i = i + 1
               sFileIni = ReadIniFile(mvarFileList, "DPIL", "FILE" & i, "")
            End If
         Wend
         Call WriteIniFile(mvarFileList, "DPIL", "FILE" & i, sFile)
      End If
   Next
   MsgBox "Inclusão realizada!"

End Sub

Private Sub mvarFTP_CmdConectaClick()
   On Error GoTo TrataErro
   
   If mvarFTP.txtAddress.Text = "" Then
      MsgBox "Informe o nome do servidor FTP.", vbCritical
      mvarFTP.txtAddress.SetFocus
      Exit Sub
   End If
   
   If mvarConectado Then
      Call DesconectarFTP
      mvarConectado = False
      mvarResult = False
      mvarFTP.CmdConecta.Caption = "Conectar"
      Exit Sub
   End If
   
   Call ConectarFTP(mvarFTP.txtAddress, mvarFTP.txtUsername, mvarFTP.txtPassword)
   With mvarFTP
      .CmdEnvia.Enabled = mvarConectado And bLiberado
      .CmdRecebe.Enabled = mvarConectado
      .CmdMkDir.Enabled = mvarConectado And bLiberado
      .CmdDelete.Enabled = mvarConectado And bLiberado
      .lstRemote.Enabled = mvarConectado And bLiberado
      .CmdConecta.Caption = IIf(mvarConectado, "&Desconectar", "&Conectar")
   End With
   
   If mvarConectado Then
      mvarResult = True
      Call AbrirPasta(mvarRemotePath)
      Call PopulaLstRemote
   End If
   Exit Sub
   
TrataErro:
    MsgBox Err.Description, vbExclamation, "FTP"
End Sub

Private Sub mvarFTP_CmdDeleteClick()
    Dim sOper As String
    
    Dim sArq   As String
    Dim bResp  As Integer
    Dim i      As Integer
    
    bResp = ExibirPergunta("Confirma a operação ?", "Delete")
    If bResp Then
        For i = 0 To mvarFTP.lstRemote.ListCount - 1
            If mvarFTP.lstRemote.Selected(i) = True Then
                sArq = mvarFTP.lstRemote.List(i)
                '* Verifica se é um diretorio ou arquivo
                If Right(sArq, 1) = "/" Then
                    sOper = "rmdir " & Mid(sArq, 1, Len(sArq) - 1)
                Else
                    sOper = "delete " & sArq
                End If
                Call ExecutaComando(sOper)
            End If
        Next i
        Call PopulaLstRemote
    End If
End Sub

Private Sub mvarFTP_CmdEnviaClick()
   Dim i             As Integer
   Dim sArqLocal     As String
   Dim sArqRemote    As String
   Dim sPathLocal    As String
   Dim sPathRemote   As String
   
   Dim sVsRemote     As String
   Dim sVsLocal      As String
   Dim sVsArq        As String
      
   Dim bExisteVs     As Boolean
   Dim sVerMaq       As String
   Dim sDtMaq        As String
   Dim sComando      As String
   Dim bCompacta     As Boolean
   
   'Call BaixarArquivo(mvarRemotePath, mvarArqCtrl , mvarLocalPath, mvarArqCtrl )
   sVsArq = mvarArqCtrl
   sVsRemote = mvarFTP.lblRemoteDirectory.Caption & "/"
   sVsLocal = mvarFTP.dirList.List(mvarFTP.dirList.ListIndex) & "\"
   'sVsLocal = mvarLocalPath
   
   'Call ExecutaComando("recv " & sVsRemote & " " & sVsLocal)
   Call BaixarArquivo(sVsRemote, sVsArq, sVsLocal, sVsArq)
   bExisteVs = ExisteArquivo(sVsLocal & sVsArq)
   
   For i = 0 To mvarFTP.filList.ListCount - 1
      If mvarFTP.filList.Selected(i) = True Then
         sPathLocal = mvarFTP.dirList.List(mvarFTP.dirList.ListIndex) & "\"
         sPathRemote = mvarFTP.lblRemoteDirectory.Caption & "/"
         sArqLocal = mvarFTP.filList.List(i)
         sArqRemote = mvarFTP.filList.List(i)
         
         sVerMaq = ""
         sDtMaq = ""
         
         sVerMaq = GetFileVersion(sPathLocal & sArqLocal)
         sDtMaq = FileDateTime(sPathLocal & sArqLocal)
         If xVal(Replace(sVerMaq, ".", "")) <> 0 Then
            Call WriteIniFile(sVsLocal & sVsArq, "VERSAO", sArqLocal, sVerMaq)
         End If
         Call WriteIniFile(sVsLocal & sVsArq, "DATA", sArqLocal, sDtMaq)
         '********
         '* Compactar Arquivo
         sComando = Environ("PROGRAMFILES") & "\winzip\winzip32.exe"
         bCompacta = (mvarFTP.ChkZIP.Value = xtpChecked And ExisteArquivo(sComando))
         If bCompacta Then
            sComando = sComando & " -a -ef -ex "
            sComando = sComando & sPathLocal & sArqLocal & ".zip "
            sComando = sComando & sPathLocal & sArqLocal
            
            Call ExcluirArquivoRemoto(sPathRemote & sArqRemote, False)
            Call SincShell(sComando, vbHide, True)
            Call EnviarArquivo(sPathLocal, sArqLocal & ".zip", sPathRemote, sArqRemote & ".zip")
            If ExisteArquivo(sPathLocal & sArqLocal & ".zip") Then
               On Error Resume Next
               Kill sPathLocal & sArqLocal & ".zip"
            End If
         Else
            Call ExcluirArquivoRemoto(sPathRemote & sArqRemote & ".zip", False)
            Call EnviarArquivo(sPathLocal, sArqLocal, sPathRemote, sArqRemote)
         End If
      End If
   Next i
   
   Call EnviarArquivo(sVsLocal, sVsArq, sVsRemote, sVsArq)
   
   Call PopulaLstRemote
   mvarFTP.filList.Refresh
End Sub

Private Sub mvarFTP_CmdEnviarListaClick()
   Dim i             As Integer
   Dim sArqLocal     As String
   Dim sPathLocal    As String
   Dim sPathRemote   As String
   
   Dim sVsRemote     As String
   Dim sVsLocal      As String
   Dim sVsArq        As String
   
   Dim sVerMaq       As String
   Dim sDtMaq        As String
   Dim sVerRem       As String
   Dim sDtRem        As String

   Dim sArqZip       As String
   Dim bEnvia        As Boolean
   Dim bBaixou       As Boolean
   Dim bContinua     As Boolean
   Dim sLstError     As String
   Dim sLstOk        As String
   Dim CollArq       As Collection
   
   
   'Call BaixarArquivo(mvarRemotePath, mvarArqCtrl , mvarLocalPath, mvarArqCtrl )
   sLstError = ""
   sVsArq = mvarArqCtrl
   sVsRemote = mvarFTP.lblRemoteDirectory.Caption & "/"
   sVsLocal = ResolvePathName(App.Path)  'mvarFTP.dirList.List(mvarFTP.dirList.ListIndex) & "\"
   
   If Not mvarConectado Then
      Call mvarFTP_CmdConectaClick
   End If
   
   If ExisteArquivo(sVsLocal & sVsArq) Then
      Call ExcluirArquivo(sVsLocal & sVsArq)
   End If
   bBaixou = BaixarArquivo(sVsRemote, sVsArq, sVsLocal, sVsArq)
   If Not bBaixou Then
      Call DesconectarFTP(False)
      Wait 2
      Call ConectarFTP(mvarFtpIP, mvarFtpUsu, mvarFTPPwd)
      If mvarConectado Then
         Call AbrirPasta(sVsRemote)
         bBaixou = BaixarArquivo(sVsRemote, sVsArq, sVsLocal, sVsArq)
      End If
   End If
   bEnvia = bBaixou
   If ExisteArquivo(sVsLocal & sVsArq) Then
      bEnvia = True
   Else
      bEnvia = False
      bBaixou = False
      While Not bBaixou
         If vbYes = ExibirPergunta("Arquivo '" & sVsLocal & sVsArq & "' não encontrado." & vbNewLine & vbNewLine & "Tentar baixar novamente?") Then
            Call mvarFTP_CmdConectaClick
            sVsRemote = mvarFTP.lblRemoteDirectory.Caption & "/"
            bBaixou = BaixarArquivo(sVsRemote, sVsArq, sVsLocal, sVsArq)
            bEnvia = ExisteArquivo(sVsLocal & sVsArq)
         Else
            bEnvia = False
            bBaixou = True
         End If
      Wend
   End If
   If Not bEnvia Then
      If vbYes = ExibirPergunta("Deseja criar arquivo '" & sVsLocal & sVsArq & "'?", pDefaultYes:=False) Then
         Call WriteIniFile(sVsLocal & sVsArq, "VERSAO", sArqLocal, "")
         bEnvia = True
      End If
   End If
   If ExisteArquivo(sVsLocal & sVsArq) Then
      Set CollArq = ListaDirRemoto(sPathRemote)
      bContinua = True
      If CollArq.Count = 0 Then
         bContinua = (vbYes = ExibirPergunta("Não foi detectado nenhum arquivo remoto. Exporta todos os arquivos?", pDefaultYes:=False))
      End If
      
      If bContinua Then
         For i = 0 To mvarFTP.LstArquivos.ListCount - 1
            mvarFTP.LstArquivos.Selected(i) = True
            sArqLocal = GetNameFromPath(mvarFTP.LstArquivos.List(i), sPathLocal)
            sArqZip = sArqLocal & ".zip"
            If UCase(Mid(sArqLocal, 1, 3)) = "G3R" Then
 '              MsgBox "G3R"
               sArqLocal = sArqLocal
            End If
            If ExisteArquivo(sPathLocal & sArqLocal) Then
               sPathRemote = mvarFTP.lblRemoteDirectory.Caption & "/"
               
               sVerRem = ""
               sVerMaq = ""
               
               sVerRem = ReadIniFile(sVsLocal & sVsArq, "VERSAO", sArqLocal, "")
               sVerMaq = GetFileVersion(sPathLocal & sArqLocal)
               If sVerMaq = "" Then
                  sVerRem = Format(ReadIniFile(sVsLocal & sVsArq, "DATA", sArqLocal, ""), "yyyymmddhhmmss")
                  sVerMaq = Format(FileDateTime(sPathLocal & sArqLocal), "yyyymmddhhmmss")
               End If
               
               If (sVerRem < sVerMaq) Or (Trim(sVerRem & sVerRem) = "") Then
                  bEnvia = True
               Else
                  If ExisteItem(CollArq, sArqZip) Then
                     bEnvia = False
                  Else
                     bEnvia = True
                  End If
                  
               End If
               
               If bEnvia Then
                  If bDsvm Then MsgBox "stop"
                  sArqZip = sArqLocal & ".zip"
                  Call Zip(sPathLocal & sArqLocal, sPathLocal & sArqZip)
                  
                  Call ExcluirArquivoRemoto(sPathRemote & sArqZip, False)
                  If EnviarArquivo(sPathLocal, sArqZip, sPathRemote, sArqZip) Then
                     Call ExcluirArquivo(sPathLocal & sArqZip)
               
                     sVerRem = ReadIniFile(sVsLocal & sVsArq, "VERSAO", sArqLocal, "")
                     sVerMaq = GetFileVersion(sPathLocal & sArqLocal)
                     sDtMaq = Format(FileDateTime(sPathLocal & sArqLocal), "yyyy/mm/dd hh:mm:ss")
                     If xVal(Replace(sVerMaq, ".", "")) <> 0 Then
                        Call WriteIniFile(sVsLocal & sVsArq, "VERSAO", sArqLocal, sVerMaq)
                     End If
                     Call WriteIniFile(sVsLocal & sVsArq, "DATA", sArqLocal, sDtMaq)
                     Call EnviarArquivo(sVsLocal, sVsArq, sVsRemote, sVsArq)
                     
                     sLstOk = IIf(sLstOk = "", "Lista de Arquivo Atualizados: ", sLstOk)
                     sLstOk = sLstOk & vbNewLine & sPathLocal & sArqLocal
                     
                  Else
                     sLstError = IIf(sLstError = "", "Lista de Erros : ", sLstError)
                     sLstError = sLstError & vbNewLine & sPathLocal & sArqLocal
                  End If
               End If
            End If
         Next i
      End If
      If Trim(sLstError) = "" Then
         sLstOk = IIf(sLstOk = "", "Não existe arquivo para atualizar.", sLstOk)
         MsgBox sLstOk
      Else
         MsgBox sLstError
      End If
      If vbYes = ExibirPergunta("Atualiza Lista Remota?", "Host Remoto", False) Then
         Call PopulaLstRemote
      End If
   Else
      Call PopulaLstRemote
      MsgBox "Arquivo '" & sVsLocal & sVsArq & "' não encontrado."
   End If
   'Call EnviarArquivo(sVsLocal, sVsArq, sVsRemote, sVsArq)
   
End Sub

Private Sub mvarFTP_Load()
   On Error Resume Next
   
   Set mvarMe = mvarFTP
   Set xInet = mvarFTP.InetFTP
   Call LiberarTela(False)
   Call PopulaTela
   
   Screen.MousePointer = vbDefault
End Sub
Private Sub PopulaTela()
   On Error Resume Next
   With mvarFTP
      .txtAddress.Text = mvarFtpIP
      .txtUsername.Text = mvarFtpUsu
      .txtPassword.Text = mvarFTPPwd
   End With
End Sub

Private Sub mvarFTP_MnuFerr(Index As Integer)
   Select Case Index
      Case 0: Call VerificarVersao
      Case 1: '* Separador
      Case 2: Call LiberarTela(True)
   End Select
End Sub
Private Sub VerificarVersao()
   Dim MyVerif As TL_VerifVersao
   
   Set MyVerif = New TL_VerifVersao
   MyVerif.ShowCAVs
   Set MyVerif = Nothing
End Sub
Private Sub LiberarTela(bLibera As Boolean)
   If bLibera Then
      bLibera = ("DIO" = UCase(InputBox("Digite a chave de entrada.", "Chave...")))
   End If
   With mvarFTP
      .txtAddress.Enabled = bLibera
      .txtUsername.Enabled = bLibera
      .txtPassword.Enabled = bLibera
      .CmdEnvia.Enabled = bLibera
      .CmdDelete.Enabled = bLibera
      .CmdMkDir.Enabled = bLibera
      .lstRemote.Enabled = bLibera
   End With
   If bLibera Then bLiberado = bLibera
End Sub

Private Sub mvarFTP_TabControl1BeforeItemClick(ByVal Item As XtremeSuiteControls.ITabControlItem, Cancel As Variant)
   Dim i As Integer
   Dim sFileIni As String
   
   If Item.Index = 1 Then
      i = 1
      sFileIni = ReadIniFile(mvarFileList, "DPIL", "FILE" & i, "")
      mvarFTP.LstArquivos.Clear
      While sFileIni <> ""
         If sFileIni <> "" Then
            mvarFTP.LstArquivos.AddItem sFileIni
         End If
         i = i + 1
         sFileIni = ReadIniFile(mvarFileList, "DPIL", "FILE" & i, "")
      Wend
   End If
End Sub

Private Sub mvarFTP_Unload(Cancel As Integer)
   On Error GoTo Trata_erro
    
   If mvarConectado Then
      Call DesconectarFTP
   End If
   Exit Sub
Trata_erro:
   MsgBox "Erro na aplicação FTP", vbCritical
End Sub

Private Sub mvarVERIF_LblTituloDblClick()
   Dim sPWD As String
   
   sPWD = InputBoxPassword("Informe Nova Senha")
   If Trim(sPWD) <> "" Then
      mvarFTPPwd = sPWD
      If vbYes = ExibirPergunta("Deseja salvar a nova senha?", "FTP", pDefaultYes:=False) Then
         'Call WriteIniFile(mvarFileIni, "Config", "Host", "ftp.classeanet.com.br")
         'Call WriteIniFile(mvarFileIni, "Config", "UID", mvarFtpUsu)
         Call WriteIniFile(mvarFileIni, "Config", "PWD", Encrypt2(mvarFTPPwd))
      End If
   End If
End Sub

Private Sub mvarVERIF_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   'this procedure receives the callbacks from the System Tray icon.
   Dim Result As Long
   Dim msg As Long
   'the value of X will vary depending upon the scalemode setting
   With mvarVERIF
      If .ScaleMode = vbPixels Then
         msg = X
      Else
         msg = X / Screen.TwipsPerPixelX
      End If
      Select Case msg
         Case WM_LBUTTONUP        '514 restore form window
            .WindowState = vbNormal
            Result = SetForegroundWindow(.hwnd)
            .Show
         Case WM_LBUTTONDBLCLK    '515 restore form window
            .WindowState = vbNormal
            Result = SetForegroundWindow(.hwnd)
            .Show
         Case WM_RBUTTONUP        '517 display popup menu
            Result = SetForegroundWindow(.hwnd)
            .PopupMenu .mPopupSys
      
         Case WM_MOUSEMOVE
            gNotifyIcon.szTip = .lblStatus.Caption & vbNullChar
            If InStr(.lblStatus.Caption, "...") > 0 Then
               If Trim(Mid(.lblStatus.Caption, InStr(.lblStatus.Caption, "...") + 3)) <> "" Then
                  If Not PrgBar Is Nothing Then
                     gNotifyIcon.szTip = "Baixando... " & ValBr((PrgBar.Value / PrgBar.Max) * 100, 0) & "%" & vbNullChar
                  End If
               End If
            End If
            Shell_NotifyIcon NIM_MODIFY, gNotifyIcon
      End Select
   End With
End Sub

Private Sub xInet_StateChanged(ByVal State As Integer)
   Dim sSituacao As String
   Static sSituacaoAntes As String
   
   DoEvents
   mvarState = State
   sSituacao = ""
   Select Case mvarState
       Case icResolvingHost:     sSituacao = "Resolvendo Host"
       Case icHostResolved:      sSituacao = "Host Resolvido"
       Case icConnecting:        sSituacao = "Conectando ..."
       Case icConnected:
          sSituacao = "Conectado"
          mvarConectado = True
'       Case icRequesting:        sSituacao = "Requesitando ..."
'       Case icRequestSent:       sSituacao = "Requesição enviada"
       Case icReceivingResponse: sSituacao = "Recebendo ..."
'       Case icResponseReceived:  sSituacao = "Resposta recebida"
       Case icDisconnecting:     sSituacao = "Desconectando ..."
       Case icDisconnected:      sSituacao = "Desconectado": mvarConectado = False
       Case icError:             sSituacao = xInet.ResponseInfo
       Case icResponseCompleted: sSituacao = "Operação Completa"
   End Select
   
   mvarStatus = sSituacao
   mvarResult = InArray(mvarState, Array(icHostResolved, icConnected, icDisconnecting, icResponseReceived, icResponseCompleted))
   
   If Trim(sSituacao) <> "" And sSituacao <> sSituacaoAntes Then
      sSituacaoAntes = sSituacao
      
      If Right(sSituacao, 2) = vbNewLine Then
         sSituacao = Mid(sSituacao, 1, Len(sSituacao) - 2)
      End If
      sFormLog = sFormLog & sSituacao & vbNewLine
      
      If InStr(sSituacao, vbNewLine) <> 0 And Len(sSituacao) > 3 Then
         sSituacao = Mid(sSituacao, 3 + Len(sSituacao) - InStr(3, StrReverse(Trim(sSituacao)), Chr(13)))
      End If
      sSituacao = Trim(Replace(sSituacao, vbNewLine, " "))
   
      If mvarMe.Enabled And mvarMe.Visible Then
         mvarMe.lblStatus.Caption = sSituacao
      End If
   
      If Not mvarFTP Is Nothing Then
         If mvarFTP.Enabled And mvarFTP.Visible Then
            mvarFTP.txtLog.Text = sFormLog
            mvarFTP.txtLog.SelStart = Len(mvarFTP.txtLog.Text)
         End If
      End If
   End If
End Sub

Private Sub mvarVERIF_Activate()
   On Error Resume Next
   Set mvarMe = mvarVERIF
   Set xInet = mvarVERIF.InetFTP
   
   If GetTag(mvarVERIF, "1VEZ", "True") Then
      Call SetTag(mvarVERIF, "1VEZ", "False")
      Call CriarSysTray
   End If
   
   
   If mvarAutoStart Then
      mvarVERIF.LblTexto = "Esta função irá verificar arquivos de instalação do Sistema."
      Call VerificarVersoes
      mvarVERIF_CmdSairClick
   Else
      mvarVERIF.LblTexto = "Esta função irá verificar arquivos de atualização do Sistema."
   End If
End Sub

Private Sub mvarVERIF_Load()
   On Error Resume Next
   Set mvarMe = mvarVERIF
   Set xInet = mvarVERIF.InetFTP
   'Call ConectarBanco
   Call MontaTela
      
   mvarVERIF.Refresh
      
   Screen.MousePointer = vbDefault
End Sub
Private Sub CriarSysTray()
   With gNotifyIcon
      .cbSize = Len(gNotifyIcon)
      .hwnd = mvarVERIF.hwnd
      .uId = vbNull
      .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
      .uCallBackMessage = WM_MOUSEMOVE
      .hIcon = mvarVERIF.Icon
      .szTip = "Baixando Arquivo..." & vbNullChar
   End With
   Shell_NotifyIcon NIM_ADD, gNotifyIcon
End Sub
Private Sub mvarVERIF_Unload(Cancel As Integer)
   Dim i As Integer
   
   'this removes the icon from the system tray
   Shell_NotifyIcon NIM_DELETE, gNotifyIcon
   
   On Error Resume Next
   xInet.Execute , "quit"
   
   If mvarConectado Then
      Call DesconectarFTP
   End If
   Do While xInet.StillExecuting
      DoEvents
      i = i + 1
      If i > 32000 Then
         xInet.Cancel
         Exit Do
      End If
   Loop
   
   mvarConectado = False
   Set PrgBar = Nothing
   Set xInet = Nothing
End Sub
Private Sub mvarVERIF_LblStatusDblClick()
'   Dim MyEditor As New EDITOR
'
'   If mvarisSystem Then
'      With MyEditor
'         .Texto = sFormLog
'         .Caption = "Log de Atualização."
'         .Width = 10000
'         .Height = 9000
'         .Show
'      End With
'      Set MyEditor = Nothing
'   End If
End Sub
Private Sub mvarVERIF_MnuMainClick(Index As Integer)
   Select Case Index
      Case 0
         Dim Sql As String
         
         Call ConectarBanco
         
         Sql = "Delete "
         Sql = Sql & " From GLOGIN"
         Sql = Sql & " Where CODSIS=" & SqlStr(mvarCODSIS)
         Sql = Sql & " And IDUSU<>" & SqlStr(mvarIDUSU)
         Sql = Sql & " And COMPUTERNAME<>" & SqlStr(ComputerName())

         If mvarSys.XDb.Executa(Sql) Then
            MsgBox "Conexões derrubadas.", vbInformation + vbOKOnly, "Controle de Versão."
         End If
      Case 1
      
   End Select
End Sub
Private Sub mvarVERIF_CmdSairClick()
   On Error Resume Next
   If mvarConectado Then Call DesconectarFTP
   Call xInet.Execute(Operation:="quit")
   Set xInet = Nothing
   
   Unload mvarVERIF
End Sub
Private Sub mvarVERIF_CmdVerificarClick()
   Call VerificarVersoes
   If mvarAutoInstall Then
      mvarVERIF_CmdSairClick
   End If
End Sub
'Private Sub VerificarVersoes_Old()
'   Dim CollArq    As Collection
'   Dim vArq       As Variant
'   Dim nCount     As Integer
'   Dim nVez       As Integer
'   Dim bReload    As Boolean
'   Dim sSituacao  As String
'   Dim sVerFTP    As String
'   Dim sVerMaq    As String
'   Dim sPathVsTxt As String
'   Dim sPathAux   As String
'   Dim bZiped     As Boolean
'   Dim sComando   As String
'   Dim bBaixaZip  As Boolean
'   Dim i          As Integer
'   Dim sWinZip32  As String
'   Dim lPath      As String
'   Dim bExisteDescompactador As Boolean
'   Dim sComandoBase As String
'
'   Screen.MousePointer = vbHourglass
'
'   mvarVERIF.CmdVerificar.Visible = False
'   mvarVERIF.LblTitulo.Caption = "Baixando..."
'   If Not PrgBar Is Nothing Then PrgBar.Scrolling = 2
'
'   sSituacao = ""
'   mvarVERIF.lblStatus.Caption = ""
'
'   nVez = 3
'   If Not mvarConectado Then
'      For i = 1 To nVez
'         Call ConectarFTP(mvarFtpIP, mvarFtpUsu, mvarFTPPwd)
'         If mvarConectado Then
'            i = nVez
'         Else
'            Wait 3
'         End If
'      Next
'   End If
'
'   If mvarConectado Then
'      nCount = 0
'      lPath = "C:\"
'      sPathVsTxt = mvarLocalPath
'
'
'      '*******
'      '* Baixar arquivo de controle de  versão 'Versao.txt'
'      Call CriarDiretorio(sPathVsTxt)
'      Call ExcluirArquivo(sPathVsTxt & mvarArqCtrl, False)
'      Call BaixarArquivo(mvarRemotePath, mvarArqCtrl, sPathVsTxt, mvarArqCtrl)
'
'      'sWinZip32 = Environ("PROGRAMFILES") & "\winzip\winzip32.exe"
'      'If ExisteArquivo(sWinZip32) Then
'      '   bExisteDescompactador = True
'      '   sComandoBase = sWinZip32
'      '   sComandoBase = sComandoBase & " -e -o "
'      'End If
'
'
'      Set CollArq = ListaDirRemoto(mvarRemotePath)
'      For Each vArq In CollArq
'         If Trim(vArq) <> "" Then
'
'            bZiped = (Right(vArq, 4) = ".zip")
'            If mvarAutoInstall Then
'               If bZiped Then
'                  bBaixaZip = (bZiped And ExisteArquivo(sWinZip32))
'                  'vArq = Mid(vArq, 1, Len(vArq) - 4)
'               End If
'            Else
'               bBaixaZip = True
'            End If
'
'            If Not bZiped Or (bZiped And bBaixaZip) Then
'               If AplicarFiltro(CStr(vArq)) Then
'                  If UCase(Right(CStr(vArq), 4)) = ".RPT" Then
'                     'sPathAux = mvarLocalPathRPT
'                     sPathAux = mvarLocalPath & "RPT\"
'                  ElseIf UCase(Right(CStr(vArq), 4)) = ".MSI" And mvarAutoInstall Then
'                     'sPathAux = Mid(mvarLocalPath, 1, Len(mvarLocalPath) - 4) & "Instalacao\"
'                     sPathAux = mvarLocalPath & "Instalacao\"
'                  Else
'                     sPathAux = mvarLocalPath
'                  End If
'                  Call CriarDiretorio(sPathAux)
'
'                  '***********
'                  '* Verfica versões ou datas dos arquivos local e remoto
'                  sVerFTP = ""
'                  sVerMaq = ""
'                  If ExisteArquivo(sPathAux & CStr(vArq)) Then
'                     sVerFTP = ReadIniFile(sPathVsTxt & mvarArqCtrl, "VERSAO", CStr(vArq), "999.999.999.999")
'                     sVerMaq = GetFileVersion(sPathAux & CStr(vArq))
'                     If sVerMaq = "" Then
'                        sVerFTP = Format(ReadIniFile(sPathVsTxt & mvarArqCtrl, "DATA", CStr(vArq), Now()), "yyyy/mm/dd hh:mm:ss")
'                        sVerMaq = Format(FileDateTime(sPathAux & CStr(vArq)), "yyyy/mm/dd hh:mm:ss")
'                     End If
'                  Else
'                     sVerFTP = "999"
'                  End If
'
'                  If sVerMaq < sVerFTP And (UCase(vArq) <> UCase(mvarArqCtrl)) Then
'                     sSituacao = "Copiando " & CStr(vArq) & " ..."
'                     mvarVERIF.lblStatus.Caption = "Copiando " & CStr(vArq) & " ..."
'
'                     '***********
'                     '* Criar cópia de segurança
'                     If ExisteArquivo(sPathAux & CStr(vArq)) And mvarAutoInstall Then
'                        Call CopiarArquivo(sPathAux & CStr(vArq), sPathAux & CStr(vArq) & ".Bak")
'                     End If
'
'                     sSituacao = "Baixando Arquivo " & CStr(vArq) & " ..."
'                     mvarVERIF.lblStatus.Caption = sSituacao
'                     mvarVERIF.lblStatus.Tag = mvarVERIF.lblStatus.Caption
'                     mvarVERIF.Refresh
'
'                     '***********
'                     '* Baixar arquivo atualizado
'                     'sComando = sWinZip32
'                     'bZiped = (bZiped And ExisteArquivo(sComando))
'
'                     If bZiped And bBaixaZip Then
'                        If Not ExisteArquivo(sPathAux & CStr(vArq)) Or mvarAutoInstall Then
'                           Call BaixarArquivo(mvarRemotePath, CStr(vArq), lPath, CStr(vArq))
'                           Call CopiarArquivo(lPath & CStr(vArq), sPathAux & CStr(vArq))
'                           On Error Resume Next
'                           Kill lPath & CStr(vArq)
'                        End If
'                        '********
'                        '* Descompactar Arquivo
'                        mvarVERIF.lblStatus.Caption = "Descompactando arquivo..."
'
'                        If bExisteDescompactador Then
'                           sComando = sComandoBase & sPathAux & vArq
'                           sComando = sComando & " " & sPathAux
'
'                           If mvarAutoInstall Then
'                              If ExisteArquivo(sPathAux & Mid(vArq, 1, Len(vArq) - 3) & "dll") Or ExisteArquivo(sPathAux & Mid(vArq, 1, Len(vArq) - 3) & "ocx") Then
'                                 On Error Resume Next
'                                 Kill sPathAux & vArq
'                              End If
'
'                              Call SincShell(sComando, vbHide, True)
'
'                              If ExisteArquivo(sPathAux & Mid(vArq, 1, Len(vArq) - 3) & "dll") Or ExisteArquivo(sPathAux & Mid(vArq, 1, Len(vArq) - 3) & "ocx") Then
'                                 If ExisteArquivo(sPathAux & vArq) Then
'                                    On Error Resume Next
'                                    Kill sPathAux & vArq
'                                 End If
'                              End If
'
'                           Else
'                              Call SincShell(sComando, vbHide, True)
'                           End If
'                        End If
'                     Else
'                        If Not ExisteArquivo(sPathAux & CStr(vArq)) Or mvarAutoInstall Then
'                           Call BaixarArquivo(mvarRemotePath, CStr(vArq), lPath, CStr(vArq))
'                           Call CopiarArquivo(lPath & CStr(vArq), sPathAux & CStr(vArq))
'                           On Error Resume Next
'                           Kill lPath & CStr(vArq)
'                        End If
'                     End If
'                     If ExisteArquivo(sPathAux & CStr(vArq)) Then
'                        nCount = nCount + 1
'                     End If
'
'
'                     '***********
'                     '* Verificar se copiou senão voltar Cópia de Segurança
'                     If mvarAutoInstall Then
'                        If ExisteArquivo(sPathAux & CStr(vArq)) Then
'                           '***********
'                           '* Salvar Cópia de segurança
'                           If dir(sPathAux & "Bak\", vbDirectory) = "" Then
'                              Call MkDir(sPathAux & "Bak\")
'                           End If
'
'                           If ExisteArquivo(sPathAux & CStr(vArq) & ".Bak") Then
'                              If ExisteArquivo(sPathAux & "Bak\" & CStr(vArq)) Then
'                                 If GetFileVersionNumber(sPathAux & CStr(vArq)) > GetFileVersionNumber(sPathAux & CStr(vArq) & ".Bak") Or GetFileVersionNumber(sPathAux & CStr(vArq)) = 0 Then
'                                    bReload = True
'                                    Call CopiarArquivo(sPathAux & CStr(vArq) & ".Bak", sPathAux & "Bak\" & CStr(vArq))
'                                 End If
'                              Else
'                                 Call CopiarArquivo(sPathAux & CStr(vArq) & ".Bak", sPathAux & "Bak\" & CStr(vArq))
'                              End If
'                              If ExisteArquivo(sPathAux & "Bak\" & CStr(vArq)) Then
'                                 Call ExcluirArquivo(sPathAux & CStr(vArq) & ".Bak")
'                              End If
'                           End If
'
'                        Else
'                           '******
'                           '* Retornar Cópia de Segurança
'                           If ExisteArquivo(sPathAux & CStr(vArq) & ".Bak") Then
'                              Call CopiarArquivo(sPathAux & CStr(vArq) & ".Bak", sPathAux & CStr(vArq))
'                              Call ExcluirArquivo(sPathAux & CStr(vArq) & ".Bak")
'                           End If
'                        End If
'                     End If
'                     sSituacao = "Arquivo " & CStr(vArq) & IIf(ExisteArquivo(sPathAux & CStr(vArq)), " ", " não ") & " copiado."
'                     sFormLog = sFormLog & sSituacao & vbNewLine
'                     mvarVERIF.lblStatus.Caption = sSituacao
'                  End If
'               End If
'            End If
'         End If
'      Next
'
'      '*******
'      '* Apagar arquivo de controle de  versão 'Versao.txt' baixado.
'      'Call ExcluirArquivo(sPathVsTxt & mvarArqCtrl)
'
'
'      If nCount = 1 Then
'         sSituacao = "Atualização Concluída. (1 arquivo baixado)"
'      Else
'         sSituacao = "Atualização Concluída. (" & nCount & " arquivos baixados)"
'      End If
'
'      sFormLog = sFormLog & sSituacao & vbNewLine
'      mvarVERIF.lblStatus.Caption = sSituacao
'
'      If bReload Then
'         MsgBox "Para que atualização tenha efeito, o Sistema deverá ser re-inicializado.", vbInformation + vbOKOnly, "Atualização de Sistema"
'      End If
'   End If
'   If Not PrgBar Is Nothing Then PrgBar.Scrolling = 0
'   mvarVERIF.CmdVerificar.Visible = mvarAutoInstall
'   Screen.MousePointer = vbDefault
'   Exit Sub
'Saida:
'   mvarVERIF.LblTitulo.Caption = "Atualizar Sistema"
'   mvarVERIF.CmdVerificar.Visible = True
'End Sub
Public Function ConectarFTP(ByRef pFtpIP As String, ByRef pFtpUsu As String, ByRef pFTPPwd As String, Optional pExibeMsg As Boolean = True) As Boolean
   On Error GoTo Trata_erro
   
   Dim nProtocol As Integer
   Dim nRemotePort As Integer
   Dim nAccesType As Integer
   Dim sProxy As String
   Dim sRemoteHost As String
   Dim nRequestTimeout As Integer
   
   Dim sFtpIP As String
   Dim sFtpUsu As String
   Dim sFTPPwd As String
   
   
   mvarConectado = False
   nProtocol = icFTP '2
   nRemotePort = 21
   nAccesType = icUseDefault '0
   
   mvarFileIni = App.Path & "\" & "CAVERSAO.ini"
   If Not ExisteArquivo(mvarFileIni) Then
      mvarFileIni = mvarLocalPath & "CAVERSAO.ini"
   End If
   
   sProxy = ""
   sRemoteHost = ""
   nRequestTimeout = 60
   If ExisteArquivo(mvarFileIni) Then
      nProtocol = xVal(ReadIniFile(mvarFileIni, "Config", "Protocol", "2"))
      nRemotePort = xVal(ReadIniFile(mvarFileIni, "Config", "RemotePort", "21"))
      nAccesType = xVal(ReadIniFile(mvarFileIni, "Config", "AccesType", "0"))
      sProxy = ReadIniFile(mvarFileIni, "Config", "Proxy", "")
      sRemoteHost = ReadIniFile(mvarFileIni, "Config", "RemoteHost", "")
      nRequestTimeout = xVal(ReadIniFile(mvarFileIni, "Config", "RequestTimeout", "60"))
   End If
   
   If xInet Is Nothing Then
      Set mvarVERIF = New FrmVerifVersao
      Set mvarMe = mvarVERIF
      Set xInet = mvarVERIF.InetFTP
   End If
   DoEvents
   With xInet
      .Protocol = nProtocol
      .RemotePort = nRemotePort
      .AccessType = nAccesType
      
      .Proxy = sProxy
      .RemoteHost = sRemoteHost
      .RequestTimeout = nRequestTimeout
      '.ResponseCode = .ResponseCode
      '.ResponseInfo = .ResponseInfo
      '.StillExecuting = .StillExecuting
      
      
      .URL = pFtpIP
      .UserName = pFtpUsu
      .Password = pFTPPwd
      '.OpenURL
   End With
   
   mvarConectado = True

   Call ExecutaComando("dir", pExibeMsg:=pExibeMsg)

   
   ConectarFTP = mvarConectado
   Exit Function
Trata_erro:
   If Not UCase(Err.Description) = UCase("Automation error") Then
      If ExisteArquivo(mvarFileIni) Then
         sFtpIP = ReadIniFile(mvarFileIni, "Config", "Host", "ftp.classeanet.com.br")
         sFtpUsu = ReadIniFile(mvarFileIni, "Config", "UID", "classeanet")
         sFTPPwd = Decrypt2(ReadIniFile(mvarFileIni, "Config", "PWD", Encrypt2(mvarFTPPwd)))
         If pFtpUsu = "classeanet" Then
            pFtpIP = sFtpIP
            pFtpUsu = sFtpIP
            pFTPPwd = sFtpUsu
         End If
         Call PopulaTela
         If Not ConectarFTP(sFtpIP, sFtpUsu, sFTPPwd, False) Then
            mvarConectado = True
            ConectarFTP = mvarConectado
         End If
      End If
      If Not mvarConectado Then
         If pExibeMsg Then MsgBox Err.Description, vbExclamation, "FTP"
      End If
   End If
End Function
Public Sub DesconectarFTP(Optional pExibeMsg As Boolean = True)
   Call ExecutaComando("quit", pExibeMsg:=pExibeMsg)
End Sub
Public Sub ExecutaComando(ByVal sOperacao As String, Optional pTermina As Boolean = True, Optional pExibeMsg As Boolean = True)
   On Error GoTo Trata_erro
   DoEvents
   If mvarConectado Then
      If xInet.StillExecuting Then
         xInet.Cancel
      End If
      'sFormLog = ""
      sFormLog = sFormLog & "Comando: " & sOperacao & vbNewLine
   
      xInet.Execute , sOperacao
      
      If pTermina Then
         TerminaComando
      End If
   End If
   Exit Sub
   
Trata_erro:
   On Error Resume Next
   Dim sRemoteHost As String
   If Not xInet Is Nothing Then
      sRemoteHost = xInet.RemoteHost
   End If
   If sRemoteHost = "" Then
      sRemoteHost = ReadIniFile(mvarLocalPath & "CAVERSAO.ini", "Config", "RemoteHost", "")
   End If
   If Err.Number <> 0 Then
      If pExibeMsg Then MsgBox "Não foi possivel efetuar operação com : " & sRemoteHost & vbCrLf & " erro : " & Err.Number
   End If
   mvarConectado = False
   mvarResult = False
End Sub
Private Sub TerminaComando()
   Dim nAux As Double
   DoEvents
   'nAux = Timer()
   Do While xInet.StillExecuting
      'If Timer() - nAux > 10 Then
      '   Exit Do
      'End If
      DoEvents
      If Not mvarConectado Then
         xInet.Cancel
         Exit Do
      End If
   Loop
End Sub
Public Function BaixarArquivo(Optional pLocalOrig As String, Optional pArqOrig As String, Optional pLocalDest As String, Optional pArqDest As String, Optional pReWrite As Boolean = True, Optional bViewFlood As Boolean = False) As Boolean
   Dim sOper      As String
   Dim sArqOri    As String
   Dim sArqDest   As String
   Dim nTam       As Double
   Dim nTamTotal  As Double
   Dim sLabelOrig As String
   Dim bVeVerificar As Boolean
   
   DoEvents
   On Error GoTo TrataErro
   
   If Not mvarVERIF Is Nothing Then
      bVeVerificar = mvarVERIF.CmdVerificar.Visible
      mvarVERIF.CmdVerificar.Visible = False
   End If
    
   If Trim$(pArqOrig) = "" Then
      pArqOrig = InputBox("Informe o nome do arquivo.", "Baixar Arquivo")
   End If
   If Trim$(pLocalOrig) = "" Then
      pLocalOrig = InputBox("Informe o local origem do arquivo.", "Baixar Arquivo")
   End If
   pLocalOrig = pLocalOrig & IIf(Right(pLocalOrig, 1) = "/", "", "/")
   
   
   If Trim$(pArqDest) = "" Then
      pArqDest = pArqOrig
   End If
   If Trim$(pLocalDest) = "" Then
      pLocalDest = InputBox("Informe o local destino do arquivo.", "Baixar Arquivo")
   End If
   pLocalDest = ResolvePathName(pLocalDest)
   
   sArqOri = pLocalOrig & pArqOrig
   sArqDest = pLocalDest & pArqDest
   If Trim$(sArqOri) <> "" Then
      DoEvents
      sOper = "SIZE " & sArqOri
      Call ExecutaComando(sOper)
      nTamTotal = CDbl("0" & xInet.GetChunk(1024, icString))
      nTamTotal = Round(nTamTotal / 1024, 2)
      
      If ExisteArquivo(sArqDest) Then
         If pReWrite Then
            Call ExcluirArquivo(sArqDest, False)
            
            sOper = "recv " & sArqOri & " """ & sArqDest & """"
            Call ExecutaComando(sOper, False)
         End If
      Else
         sOper = "recv " & sArqOri & " """ & sArqDest & """"
         Call ExecutaComando(sOper, False)
      End If
      
      sLabelOrig = mvarMe.lblStatus.Caption
      Do While xInet.StillExecuting
         DoEvents
         If Not mvarConectado Then
            xInet.Cancel
            Exit Do
         End If
         If sArqDest <> "" Then
            DoEvents
            If ExisteArquivo(sArqDest) Then
               On Error Resume Next
               nTam = Round(FileLen(sArqDest) / 1024, 2)
            End If
            DoEvents
            mvarMe.lblStatus.Caption = mvarMe.lblStatus.Tag & " " & Format(nTam, "#0,0.00") & " / " & Format(nTamTotal, "#0,0.00") & " Kb"
            
            If PrgBar.Scrolling = 0 And Not mvarMe.Visible Then
               PrgBar.Scrolling = 1
            End If
            
'MsgBox "PrgBar.Scrolling = " & PrgBar.Scrolling & vbNewLine & "bViewFlood = " & bViewFlood & vbNewLine & "mvarMe.Visible= " & mvarMe.Visible & vbNewLine & "GetTag(mvarMe, 'Visible', '0') = " & GetTag(mvarMe, "Visible", "0") & vbNewLine
            If PrgBar.Scrolling = 1 Then
               If bViewFlood And Not mvarMe.Visible And GetTag(mvarMe, "Visible", "0") = 0 Then
                  Call SetTag(mvarMe, "Visible", "1")
                  mvarMe.Show
                  mvarMe.Visible = True
               End If
               PrgBar.Min = 0
               PrgBar.Max = nTamTotal
               PrgBar.Value = nTam
            End If
            'mvarMe.lblStatus.Refresh
         End If
         DoEvents
         mvarMe.lblStatus.Caption = mvarMe.lblStatus.Tag & " " & Format(nTam, "#0,0.00") & " / " & Format(nTamTotal, "#0,0.00") & " Kb"
      Loop
      mvarMe.lblStatus.Caption = sLabelOrig
      mvarMe.lblStatus.Refresh
   End If
   BaixarArquivo = ExisteArquivo(sArqDest)
   
   If Not mvarVERIF Is Nothing Then
      mvarVERIF.CmdVerificar.Visible = bVeVerificar
   End If
   Exit Function
TrataErro:
   If UCase(Err.Description) <> UCase("Automation error") Then
      Resume Next
   End If
End Function
Public Function EnviarArquivo(Optional pLocalOrig As String, Optional pArqOrig As String, Optional pLocalDest As String, Optional pArqDest As String, Optional pExibeMsg As Boolean = True) As Boolean
   Dim sOper      As String
   Dim sLabelOrig As String
   Dim nTamTotal  As Boolean
    
   DoEvents
   If Trim$(pArqOrig) = "" Then
      pArqOrig = InputBox("Informe o nome do arquivo.", "Enviar Arquivo")
   End If
   If Trim$(pLocalOrig) = "" Then
      pLocalOrig = InputBox("Informe o local origem do arquivo.", "Enviar Arquivo")
   End If
   pLocalOrig = ResolvePathName(pLocalOrig)
   
   If Trim$(pLocalDest) = "" Then
      pLocalDest = InputBox("Informe o local destino do arquivo.", "Enviar Arquivo")
   End If
   If Trim$(pArqDest) = "" Then pArqDest = pArqOrig
   pLocalDest = pLocalDest & IIf(Right(pLocalDest, 1) = "/", "", "/")
    
   If Trim$(pLocalOrig & pArqOrig) <> "" Then
      If ExisteArquivo(pLocalOrig & pArqOrig) Then
         nTamTotal = Round(FileLen(pLocalOrig & pArqOrig) / 1024, 2)
      End If
   
'      Call CopiarArquivo(pLocalOrig & pArqOrig, "C:\" & pArqOrig)
      sOper = "send """ & pLocalOrig & pArqOrig & """ " & pLocalDest & pArqDest
'      sOper = "send C:\" & pArqOrig & " " & pLocalDest & pArqDest
      Call ExecutaComando(sOper, pExibeMsg:=pExibeMsg)
                  
      sLabelOrig = mvarMe.lblStatus.Caption
      
'      Call ExecutaComando(sOper, False, pExibeMsg:=pExibeMsg)
'
'      sLabelOrig = mvarMe.lblStatus.Caption
'      Do While xInet.StillExecuting
'         If Not bConectado Then
'            xInet.Cancel
'            Exit Do
'         End If
'         If pArqDest <> "" Then
'            sOper = "SIZE " & pArqDest
'            Call ExecutaComando(sOper, pExibeMsg:=pExibeMsg)
'            nTam = CDbl("0" & xInet.GetChunk(1024, icString))
'            nTam = Round(nTam / 1024, 2)
'
'            DoEvents
'            mvarMe.lblStatus.Caption = mvarMe.lblStatus.Tag & " " & Format(nTam, "#0.00") & " / " & Format(nTamTotal, "#0.00") & " Kb"
'            'mvarMe.lblStatus.Refresh
'         End If
'         DoEvents
'         mvarMe.lblStatus.Caption = mvarMe.lblStatus.Tag & " " & Format(nTam, "#0.00") & " / " & Format(nTamTotal, "#0.00") & " Kb"
'      Loop
      mvarMe.lblStatus.Caption = sLabelOrig
      mvarMe.lblStatus.Refresh
   End If
   EnviarArquivo = (sLabelOrig = "Operação Completa")
On Error Resume Next
'Kill "C:\" & pArqOrig
End Function
Public Sub AbrirPasta(pCaminho As String)
   Dim sAux As String
   Dim sPasta As String
   Dim sAtual As String
   Dim sAntes As String
   Dim sOper  As String
   
   pCaminho = pCaminho & IIf(Right(pCaminho, 1) = "/", "", "/")
   
   If Trim(pCaminho) <> "/" Then
   
      Call ExecutaComando("pwd")
      sAtual = xInet.GetChunk(1024, icString)
      sAtual = sAtual & IIf(Right(sAtual, 1) = "/", "", "/")
      If sAtual = Mid(pCaminho, 1, Len(sAtual)) Then
         sAux = Mid(pCaminho, Len(sAtual))
      Else
         sAux = pCaminho
      End If
      
      While InStr(2, sAux, "/") <> 0 And mvarConectado
         sPasta = Mid(sAux, 1, InStr(2, sAux, "/"))
         If Right(sPasta, 1) = "/" Then
            sPasta = Mid(sPasta, 1, Len(sPasta) - 1)
         End If
         sOper = "cd " & Replace(sPasta, "/", "")
         Call ExecutaComando(sOper)

         sAntes = sAtual
         Call ExecutaComando("pwd")
         sAtual = xInet.GetChunk(1024, icString)
         
         If sAntes = sAtual Or sAtual = "" Then
            MsgBox "Sistema não consegue chegar em : " & pCaminho & vbNewLine & "Pasta Anterior : " & sAntes & vbNewLine & "Pasta Atual : " & sAtual, vbOKOnly + vbExclamation, "FTP Message"
            Call DesconectarFTP
            mvarConectado = False
            mvarResult = False
         End If

         sAux = Mid(sAux, Len(sPasta) + 1)
      Wend
   Else
      If Right(pCaminho, 1) = "/" Then
         pCaminho = Mid(pCaminho, 1, Len(pCaminho) - 1)
      End If
      
      
      Call ExecutaComando("cd " & pCaminho)
   End If
End Sub
Public Function ListaDirRemoto(Optional pPasta As String, Optional pTipo_Arq_Pasta As Integer = 1) As Collection
   Dim i       As Integer
   Dim nIni    As Integer
   Dim nTam    As Integer
   Dim vData   As Variant
   Dim sAux    As String
   Dim oColl    As Collection
   Dim oCollDir As Collection
   Dim oCollArq As Collection
   Dim sAtual  As String
   
   Set oColl = New Collection
   Set oCollArq = New Collection
   Set oCollDir = New Collection
   
   On Error GoTo TrataErro
   nIni = 1

   DoEvents
   Call ExecutaComando("pwd")
   sAtual = xInet.GetChunk(1024, icString)
   If Right(sAtual, 1) <> "/" Then sAtual = sAtual & "/"
   If Trim(pPasta) = "" Then pPasta = sAtual
   If pPasta = sAtual Then
      If pTipo_Arq_Pasta = 1 Then
         If Not gCollArq Is Nothing Then
            If gCollArq.Count > 0 Then
               Set ListaDirRemoto = gCollArq
               Exit Function
            End If
         End If
      End If
   Else
      Call AbrirPasta(pPasta)
      
      Call ExecutaComando("pwd")
      sAux = xInet.GetChunk(1024, icString)
      If sAtual = sAux Or sAux = "" Or Not mvarConectado Then
         GoTo Saida
      End If
      sAtual = sAux
   End If

   If bDsvm Then MsgBox "stop"
   Call ExecutaComando("dir")

   Do
      vData = xInet.GetChunk(1024, icString)
      DoEvents

      For i = 1 To Len(vData)
         If Mid(vData, i, 1) = Chr(13) Then
            sAux = Mid(vData, nIni, nTam)
            If nTam > 0 And sAux <> "./" Then
               oColl.Add sAux, sAux
               If InStr(sAux, "/") Then
                  oCollDir.Add Mid(vData, nIni, nTam), Mid(vData, nIni, nTam)
               Else
                  oCollArq.Add Mid(vData, nIni, nTam), Mid(vData, nIni, nTam)
               End If
            End If
            nIni = i + 2
            nTam = -1
         Else
            nTam = nTam + 1
         End If
      Next i
   Loop While LenB(vData) > 0
   GoTo Saida
TrataErro:
   Resume Next
Saida:
   If pTipo_Arq_Pasta = 1 Then
      Set ListaDirRemoto = oCollArq
   ElseIf pTipo_Arq_Pasta = 2 Then
      Set ListaDirRemoto = oCollDir
   Else
      Set ListaDirRemoto = oColl
   End If
   
End Function
Public Sub CriarPasta(Optional pPasta As String)
   Dim sOper As String
   Dim sAux  As String
   Dim sPasta As String
   
   If Trim$(pPasta) = "" Then
      pPasta = InputBox("Informe o nome da pasta.", "Cria Pasta")
   End If
   If Trim$(pPasta) <> "" Then
      If InStr(pPasta, "/") <> 0 Then
         sAux = pPasta
         While InStr(sAux, "/") <> 0
            sPasta = Mid(sAux, 1, InStr(sAux, "/"))
            If Right(sPasta, 1) = "/" Then sPasta = Mid(sPasta, 1, Len(sPasta) - 1)
            If Trim(sPasta) <> "" And InStr(sPasta, "/") = 0 Then
               Call ExecutaComando("mkdir " & sPasta)
               Call ExecutaComando("cd " & sPasta)
            End If
            sAux = Mid(sAux, Len(sPasta) + 2)
         Wend
      Else
         Call ExecutaComando("mkdir " & pPasta)
      End If
   End If
End Sub
Public Sub ExcluirArquivoRemoto(Optional pArquivo As String, Optional bConfirma As Boolean = True, Optional pExibeMsg As Boolean = True)
   Dim sOper As String
   Dim nResp As Integer
      
   If Trim$(pArquivo) = "" Then
      pArquivo = InputBox("Informe o nome do arquivo ou pasta(/).", "Excluir Arquivo/Pasta")
   End If
   If Trim$(pArquivo) <> "" Then
      If bConfirma Then
         nResp = MsgBox("Confirma a operação ?", vbQuestion + vbYesNo, "Delete")
      Else
         nResp = vbYes
      End If
      If nResp = vbYes Then
         ' Verifica se é um diretorio ou arquivo
         If Right(pArquivo, 1) = "/" Then
             sOper = "rmdir " & Mid(pArquivo, 1, Len(pArquivo) - 1)
         Else
             sOper = "delete " & pArquivo
         End If
         Call ExecutaComando(sOper, pExibeMsg:=pExibeMsg)
      End If
   End If
End Sub
Private Function AplicarFiltro(pArq As String) As Boolean
   Dim sFiles As String
   Dim bFiltro As Boolean
   
   pArq = UCase(pArq)
   bFiltro = True
   
   sFiles = UCase(ReadIniFile(mvarLocalPath & mvarArqCtrl, "FILTRO", "FILES", ""))
   If Trim(sFiles) <> "" Then
      bFiltro = (InStr(sFiles, "|" & pArq & "|") <> 0) And bFiltro
   End If
   AplicarFiltro = bFiltro
End Function
Private Sub MontaTela()
   'On Error Resume Next
   
   mvarMe.lblStatus.Caption = ""
   mvarMe.MnuMain(0).Enabled = mvarisSystem
   If mvarAutoStart And Not mvarFTP Is Nothing Then
      mvarVERIF.LblTexto = "Esta função irá verificar arquivos de instalação do Sistema."
   Else
      mvarVERIF.LblTexto = "Esta função irá verificar arquivos de atualização do Sistema."
   End If
   
   Set PrgBar = mvarMe.Controls.Add("Codejock.ProgressBar.11.2.2", "PrgBar")
   With PrgBar
      .Visible = True
      .Enabled = True
      .Top = 1180
      .Height = 255
      .Left = 240
      .Width = 6100
      .UseVisualStyle = True
      .Scrolling = 0
   End With
End Sub
Private Sub ConectarBanco()
   Dim sLocalReg  As String
   Dim sPathReg   As String
   Dim bConectado As Boolean
   
   With mvarSys
      If mvarSys Is Nothing Then
         Set mvarSys = CreateObject("SysA.SetA")
      End If
      If .XDb Is Nothing Then
         bConectado = False
      Else
         bConectado = .XDb.Conectado
      End If
      
      If Not mvarSys Is Nothing Then
         If Not bConectado Then
            sLocalReg = .LocalReg
            sPathReg = App.Path & "\"
            sLocalReg = IIf(UCase(Right(sLocalReg, 4)) = ".REG", sLocalReg, sPathReg & dir(sPathReg & "*.reg"))
            sPathReg = GetSpecialFolder(38) & "ClasseA\Suprimento\"
            sLocalReg = IIf(UCase(Right(sLocalReg, 4)) = ".REG", sLocalReg, sPathReg & dir(sPathReg & "*.reg"))
            sPathReg = GetSpecialFolder(38) & "ClasseA\Producao\"
            sLocalReg = IIf(UCase(Right(sLocalReg, 4)) = ".REG", sLocalReg, sPathReg & dir(sPathReg & "*.reg"))
            sPathReg = GetSpecialFolder(38) & "ClasseA\Manutencao\"
            sLocalReg = IIf(UCase(Right(sLocalReg, 4)) = ".REG", sLocalReg, sPathReg & dir(sPathReg & "*.reg"))
            .LocalReg = IIf(UCase(Right(sLocalReg, 4)) = ".REG", sLocalReg, .LocalReg)
            
            If .IDUSU = "" Then .IDUSU = "DIO"
      
            If .XDb.Server = "" Then .XDb.Server = .GetParam("SERVER", "S2", "Conection 0")
            If .XDb.dbName = "" Then .XDb.dbName = .GetParam("DBNAME", "S2", "Conection 0")
            If .XDb.uId = "" Then .XDb.uId = .GetParam("UID", "S2", "Conection 0")
            If .XDb.PWD = "" Then .XDb.PWD = Decrypt2(.GetParam("PWD", "S2", "Conection 0"))
            If Not .XDb.Conectado Then .XDb.SrvConecta
         End If
         If .XDb.Conectado Then
            mvarMe.Caption = mvarMe.Caption & "."
            mvarisSystem = .USER.isSystem
            mvarCODSIS = .CODSIS
            mvarIDUSU = .IDUSU
         End If
         LocalPath = ResolvePathName(IIf(.PathSetup = "", mvarLocalPath, .PathSetup))
         mvarLocalPathRPT = ResolvePathName(IIf(.PathRPT = "", mvarLocalPathRPT, .PathRPT))
      End If
   End With
End Sub
Private Sub PopulaLstRemote()
   Dim CollArq    As Collection
   Dim vArq       As Variant
   
   mvarFTP.lstRemote.Clear
   
   Set CollArq = ListaDirRemoto(mvarRemotePath)
   Set gCollArq = CollArq
   For Each vArq In CollArq
      mvarFTP.lstRemote.AddItem CStr(vArq)
   Next
   
   Call ExecutaComando("pwd")
   mvarFTP.lblRemoteDirectory.Caption = xInet.GetChunk(1024, icString)
End Sub
Private Sub DefineVarSis()
   mvarRemotePath = "/private/Cliente/Comum"
   LocalPath = "C:\Tmp\"
   If mvarCODSIS = "P3R" Then
      If mvarAutoInstall Then
         mvarRemotePath = "/private/Cliente/Dpil/Comum/DLL"
         LocalPath = Environ("programfiles") & "\ClasseA\Admin\Dll\"
      Else
         mvarRemotePath = "/private/Cliente/Dpil/Comum/Instalacao"
         LocalPath = Environ("programfiles") & "\ClasseA\Admin\Instalacao\Setup\"
      End If
   End If
End Sub
Private Sub VerificarVersoes()
   Dim CollArq    As Collection
   Dim vArq       As Variant
   Dim sArqOrig   As String
   Dim sArqZip    As String
   Dim nCount     As Integer
   Dim nVez       As Integer
   Dim sSituacao  As String
   Dim sVerFTP    As String
   Dim sVerMaq    As String
   Dim sVerAux    As String
   Dim sPathVsTxt As String
   Dim sPathAux   As String
   Dim sPathBak   As String
   Dim bZiped     As Boolean
   Dim i          As Integer
   Dim bRecebe    As Boolean
   Dim bBaixou    As Boolean
   Dim bAtualiza  As Boolean
   
   'Screen.MousePointer = vbHourglass
      
   If Not IsWebConnected Then
      MsgBox "Não foi detectado conexão de Internet." & vbNewLine & "Para continuar esta operação você deverá estabelecer sua conexão." & vbNewLine & vbNewLine & "Por favor tente mais tarde.", vbInformation
      GoTo Saida
   End If
      
   mvarVERIF.CmdVerificar.Visible = False
   mvarVERIF.CmdSair.Visible = False
   mvarVERIF.LblTitulo.Caption = "Baixando..."
   If Not PrgBar Is Nothing Then PrgBar.Scrolling = 2
   
   sSituacao = ""
   mvarVERIF.lblStatus.Caption = ""
      
   nVez = 3
   If Not mvarConectado Then
      For i = 1 To nVez
         Call ConectarFTP(mvarFtpIP, mvarFtpUsu, mvarFTPPwd)
         If mvarConectado Then
            i = nVez
         Else
            Wait 3
         End If
      Next
   End If
         
   If Not mvarSys Is Nothing Then
      mvarCODSIS = mvarSys.CODSIS
   Else
      mvarCODSIS = "P3R"
   End If
   Call DefineVarSis
   
   If mvarConectado Then
      nCount = 0
      sPathVsTxt = mvarLocalPath
      
      '*******
      '* Baixar arquivo de controle de  versão 'Versao.txt'
      If mvarAutoInstall Then
         Call CriarDiretorio(sPathVsTxt)
         Call ExcluirArquivo(sPathVsTxt & mvarArqCtrl, False)
      
         If Not BaixarArquivo(mvarRemotePath, mvarArqCtrl, sPathVsTxt, mvarArqCtrl) Then
            bRecebe = True
            While bRecebe
               If vbYes = ExibirPergunta("Arquivo '" & sPathVsTxt & mvarArqCtrl & "' não encontrado." & vbNewLine & vbNewLine & "Tentar baixar novamente?") Then
                  Call DesconectarFTP(False)
                  Wait 2
                  Call ConectarFTP(mvarFtpIP, mvarFtpUsu, mvarFTPPwd)
                  If mvarConectado Then
                     bRecebe = Not BaixarArquivo(mvarRemotePath, mvarArqCtrl, sPathVsTxt, mvarArqCtrl)
                  End If
               Else
                  bRecebe = False
               End If
            Wend
         End If
      End If
      
      Set CollArq = ListaDirRemoto(mvarRemotePath)
      If CollArq.Count = 1 Then
         PrgBar.Scrolling = 1
      End If
      For Each vArq In CollArq
         If Trim(vArq) <> "" Then
            bZiped = (Right(vArq, 4) = ".zip")
            If bZiped Then
               sArqOrig = Mid(vArq, 1, Len(vArq) - 4)
            Else
               sArqOrig = CStr(vArq)
               'sArqOrig = Mid(vArq, 1, Len(vArq) - 4)
            End If
            sArqZip = sArqOrig & ".Zip"
                       
            If AplicarFiltro(CStr(vArq)) Then
               If UCase(Right(CStr(vArq), 4)) = ".RPT" Then
                  sPathAux = mvarLocalPath & "RPT\"
               ElseIf UCase(Right(CStr(vArq), 4)) = ".MSI" And mvarAutoInstall Then
                  sPathAux = mvarLocalPath & "Instalacao\"
               Else
                  sPathAux = mvarLocalPath
               End If
               Call CriarDiretorio(sPathAux)
            
               '***********
               '* Verfica versões ou datas dos arquivos local e remoto
               sVerFTP = ""
               sVerMaq = ""
               bAtualiza = False
               sVerFTP = ReadIniFile(sPathVsTxt & mvarArqCtrl, "VERSAO", sArqOrig, "999.999.999.999")
               If ExisteArquivo(sPathAux & sArqOrig) Then
                  sVerMaq = GetFileVersion(sPathAux & sArqOrig)
                  If sVerMaq = "" Then
                     sVerFTP = Format(ReadIniFile(sPathVsTxt & mvarArqCtrl, "DATA", sArqOrig, Now()), "dd/mm/yyyy hh:mm:ss")
                     sVerMaq = Format(FileDateTime(sPathAux & sArqOrig), "dd/mm/yyyy hh:mm:ss")
                     If False Then
                        Dim DtC As FILETIME
                        Dim DtA As FILETIME
                        Dim DtM As FILETIME
                        Dim hFile As Long
                        Dim createtime As SYSTEMTIME ' receives a converted form of ctime
                        Dim retval As Long ' return value
                                             
                        hFile = CreateFile(sPathAux & sArqOrig, GENERIC_WRITE, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0&)
                        If hFile = -1 Then ' if the file could not be opened
                           Debug.Print "Could not open the file C:\\MyProgram\\datafile.txt."
                        End If
      
                        ' Get the various times and dates associated with the file:
                        retval = GetFileTime(hFile, DtC, DtA, DtM)
                        ' Convert the creation time from a FILETIME structure to a SYSTEMTIME structure (for usability):
                        retval = FileTimeToSystemTime(DtC, createtime)
                        ' Display the creation date of the file:
                        Debug.Print "Creation Date:"; createtime.wMonth; "-"; createtime.wDay; "-"; createtime.wYear
                        
                        ' Close the file
                        retval = CloseHandle(hFile)
                     End If
                     
                     If DateDiff("h", CDate(sVerMaq), CDate(sVerFTP)) > 3 Then
                        bAtualiza = True
                     End If
                     sVerFTP = Format(ReadIniFile(sPathVsTxt & mvarArqCtrl, "DATA", sArqOrig, Now()), "yyyymmddhhmmss")
                     sVerMaq = Format(FileDateTime(sPathAux & sArqOrig), "yyyymmddhhmmss")
                     
                  Else
                     If sVerMaq < sVerFTP And (UCase(vArq) <> UCase(mvarArqCtrl)) Then
                        bAtualiza = True
                     End If
                  End If
               Else
                  bAtualiza = True
               End If
               
               If bAtualiza Then
                  sSituacao = "Copiando " & sArqOrig & " ..."
                  mvarVERIF.lblStatus.Caption = "Copiando " & sArqOrig & " ..."
                  
                  '***********
                  '* Criar cópia de segurança
                  sPathBak = sPathAux & "Bak\"
                  If mvarAutoInstall Then
                     Call CriarDiretorio(sPathBak)
                     Call ExcluirArquivo(sPathBak & CStr(vArq), False)
                     Call Zip(sPathAux & sArqOrig, sPathBak & sArqZip)
                  End If
                  
                  sSituacao = "Baixando Arquivo " & sArqOrig & " ...(1) "
                  mvarVERIF.lblStatus.Caption = sSituacao
                  mvarVERIF.lblStatus.Tag = mvarVERIF.lblStatus.Caption
                  mvarVERIF.Refresh
                        
                  '***********
                  '* Baixar arquivo atualizado
                  nVez = 3
                  For i = 1 To nVez
                     bBaixou = BaixarArquivo(mvarRemotePath, CStr(vArq), sPathAux, CStr(vArq), mvarAutoInstall)
                     If bBaixou Then
                        i = nVez
                     Else
                        sSituacao = "Baixando Arquivo " & sArqOrig & " ...(" & CStr(i) & ") "
                        mvarVERIF.lblStatus.Caption = sSituacao
                        mvarVERIF.lblStatus.Tag = mvarVERIF.lblStatus.Caption
                        mvarVERIF.Refresh
                        Call ConectarFTP(mvarFtpIP, mvarFtpUsu, mvarFTPPwd)
                     End If
                  Next
                  
                  If bBaixou Then
                     If bZiped Then
                        '********
                        '* Descompactar Arquivo
                        mvarVERIF.lblStatus.Caption = "Descompactando arquivo..."
                        'If ExisteArquivo(sPathAux & sArqOrig) Then
                        '   Call ExcluirArquivo(sPathAux & sArqOrig)
                        'End If
                        Call Unzip(sPathAux, CStr(vArq), sPathAux, False)
                        If mvarAutoInstall Then
                           If ExisteArquivo(sPathAux & sArqOrig) Then
                              Call ExcluirArquivo(sPathAux & CStr(vArq))
                           End If
                        End If
                     End If
                  Else
                     bBaixou = bBaixou
                  End If
                  '***********
                  '* Verificar se copiou senão voltar Cópia de Segurança
                  If mvarAutoInstall Then
                     sVerAux = GetFileVersion(sPathAux & sArqOrig)
                     If sVerAux = "" And ExisteArquivo(sPathAux & sArqOrig) Then
                        sVerAux = Format(FileDateTime(sPathAux & sArqOrig), "yyyymmddhhmmss")
                     End If
                     If sVerAux = sVerFTP Then
                        nCount = nCount + 1
                     Else
                        If mvarAutoInstall Then
                           If ExisteArquivo(sPathBak & sArqZip) Then
                              Call Unzip(sPathBak, sArqZip, sPathAux, False)
                           End If
                        End If
                     End If
                  End If
                  sSituacao = "Arquivo " & CStr(vArq) & IIf(ExisteArquivo(sPathAux & CStr(vArq)), " ", " não ") & " baixado."
                  sFormLog = sFormLog & sSituacao & vbNewLine
                  mvarVERIF.lblStatus.Caption = sSituacao
                  If vArq <> "VERSAO.TXT" Then
                     nCount = nCount + 1
                  End If
               Else
                  sVerAux = GetFileVersion(sPathAux & sArqOrig)
                  If sVerAux = "" And ExisteArquivo(sPathAux & sArqOrig) Then
                     sVerAux = Format(FileDateTime(sPathAux & sArqOrig), "yyyymmddhhmmss")
                  End If
               End If
            End If
         End If
      Next
      
      '*******
      '* Apagar arquivo de controle de  versão 'Versao.txt' baixado.
      If mvarAutoInstall Then
         Call ExcluirArquivo(sPathVsTxt & mvarArqCtrl, False)
      End If
      
      If nCount = 0 Then
         sSituacao = "Não há atualização a ser baixada."
      ElseIf nCount = 1 Then
         sSituacao = "Atualização Concluída.GetFileTime (1 arquivo baixado)"
      Else
         sSituacao = "Atualização Concluída. (" & nCount & " arquivos baixados)"
      End If
      
      sFormLog = sFormLog & sSituacao & vbNewLine
      mvarVERIF.lblStatus.Caption = sSituacao
      
      If mvarAutoInstall Then
         If nCount = 0 Then
            MsgBox sSituacao, vbInformation + vbOKOnly, "Atualização de Sistema"
         Else
            MsgBox "Para que atualização tenha efeito, o Sistema deverá ser re-inicializado.", vbInformation + vbOKOnly, "Atualização de Sistema"
         End If
      End If
   End If
   If Not PrgBar Is Nothing Then PrgBar.Scrolling = 0
   mvarVERIF.CmdVerificar.Visible = mvarAutoInstall
   Screen.MousePointer = vbDefault
   mvarVERIF.CmdSair.Visible = True
   Exit Sub
Saida:
   mvarVERIF.CmdSair.Visible = True
   mvarVERIF.LblTitulo.Caption = "Atualizar Sistema"
   mvarVERIF.CmdVerificar.Visible = True
End Sub


