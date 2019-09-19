VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_MDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe  As MDI
Attribute mvarMe.VB_VarHelpID = -1
Public WithEvents StatusBar As XtremeCommandBars.StatusBar
Attribute StatusBar.VB_VarHelpID = -1
Public WithEvents Workspace As TabWorkspace
Attribute Workspace.VB_VarHelpID = -1

Private mvarSys      As Object
Public Property Set Sys(ByVal vData As Object)
    Set mvarSys = vData
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Public Sub Show()
   mvarMe.Show
End Sub
Private Sub Class_Initialize()
   Dim MySplash As Object
   
   If InStr(UCase(App.Path), "\SISTEMAS\") = 0 Then
      Set MySplash = CreateObject("SPLP3R.Tela")
      MySplash.Show
   End If
   
   
   Set mvarMe = New MDI
   
   
   
   Call LimpaInstaciaObj
   Call InstaciaObj
      Set MySplash = CreateObject("SPLP3R.Tela")
      MySplash.Show
   
   
   '*********
   '* Testa se já existe uma cópia da aplicação rodando e define formato Data e número.
   If gIDUSU = "" Then
      If AppAtiva(App) Then
         If gDebug Then MsgBox "End"
         End
      End If
   End If
   
   If Not Sys Is Nothing Then
      With Sys
         .ExePath = App.Path
      End With
   End If
   If gDebug Then MsgBox "Entrar ExibeSenha"
   Call ExibeSenha
   Exit Sub
End Sub
Private Sub Class_Terminate()
   On Error Resume Next
   Unload mvarMe
   Set mvarMe = Nothing
End Sub
Private Sub mvarMe_Activate()
   If GetTag(mvarMe, "PrimeiraVez", "S") = "S" Then
      Call SetTag(mvarMe, "PrimeiraVez", "N")

'      MontarMenu
'      MontarToolbar
'      MontarStatusBar
'      MontarPanes
      LoadMenuDefault
   End If
   If Sys.Propriedades("FCOMANDO") <> "" Then
      If Sys.Propriedades("FCOMANDO") = "End" Then
         End
      End If
   End If
   Sys.Propriedades("Debug") = gDebug
End Sub
Private Sub mvarMe_Load()
   Screen.MousePointer = vbDefault
   
   ConfigurarAmbiente
   
   MontarMenu
   MontarToolbar
   MontarStatusBar
   MontarPanes
End Sub
Private Sub ConfigurarAmbiente(Optional bReload As Boolean)
   Dim sPath   As String
   Dim sArq    As String
   Dim sEstilo As String
   
   If Sys.CODSIS = "P3R" Then
      Sys.Style = 0
      Sys.Skin = IIf(Sys.xdb.Alias = "PRODUCAO", 2, 3)
      Sys.ShowTabWorkspace = True
   End If
   
   sPath = GetSpecialFolder(38) & "ClasseA\Arquivos Comuns\Styles\"
   'sPath = Environ("programfiles") & "\ClasseA\Arquivos Comuns\Styles\"
   Select Case Sys.Style
      Case 0: sArq = sPath & "WinXP.Luna.cjstyles"
      Case 1: sArq = sPath & "WinXP.Royale.cjstyles"
      Case 2: sArq = sPath & "Vista.cjstyles"
      Case 3: sArq = sPath & "Office2007.cjstyles"
   End Select
      
   If Not ExisteArquivo(sArq) Then Call ExtractResData("WINXPLUNA", "STYLE", sArq)
      
   Select Case Sys.Skin
      Case 0: sEstilo = ""                      '* SKIN ROYALE
      Case 1: sEstilo = "NormalBlue.ini"        '* SKIN LUNA BLUE
      Case 2: sEstilo = "NormalHomestead.ini"   '* SKIN LUNA OLIVE
      Case 3: sEstilo = "NormalMetallic.ini"    '* SKIN LUNA METALLIC
   End Select
    
   
   'Set mvarme.Icon = MyLoadPicture
  
'MsgBox "sArq= " & sArq & vbNewLine & "sEstilo= " & sEstilo
   With mvarMe.SkinFramework
      .LoadSkin sArq, sEstilo
      If Not bReload Then
         .ApplyWindow mvarMe.hwnd
         .ApplyOptions = .ApplyOptions Or xtpSkinApplyMetrics
      End If
   End With
   
   sArq = App.Path & "\Close.bmp"
   If Not ExisteArquivo(sArq) Then Call ExtractResData("CLOSE", "BMP", sArq)

   mvarMe.picHolder.Height = 0
   
   'CommandBarsGlobalSettings.App = App
   With mvarMe.CommandBars
      If Not bReload Then
         .GlobalSettings.App = App
      End If
     Set Workspace = .ShowTabWorkspace(Sys.ShowTabWorkspace)
      .VisualTheme = xtpThemeNativeWinXP   'xtpThemeWhidbey 'xtpThemeOffice2003
      .ToolTipContext.Style = xtpToolTipLuna
   End With
   
   If Workspace Is Nothing Then
      Set Workspace = mvarMe.CommandBars.ShowTabWorkspace(Sys.ShowTabWorkspace)
   End If

   mvarMe.DockingPaneManager.SetCommandBars mvarMe.CommandBars
End Sub
Private Sub MontarMenu()
   Dim Erro429    As Boolean
   Dim SetErro    As String

   On Error GoTo TrataErro

   If gDebug Then MsgBox "Menu is Nothing = " & IIf(MdiMenu Is Nothing, "True", "False")
   
   If MdiMenu Is Nothing Then
      If gDebug Then MsgBox "Set Menu = CriarObjeto(Menu.ControlMenu)"
      Set MdiMenu = CriarObjeto("Menu.ControlMenu")
   End If
   With MdiMenu
      Set .Sys = Sys
      .SisDebug = gDebug
      If gDebug Then MsgBox "Menu.ControlMenu.MontarMenu"
      Call .MontarMenu
   End With
   
   GoTo Saida
TrataErro:
   If Err = 429 Then
      MsgBox "MontarMenu" & vbNewLine & Err.Number & " - " & Err.Description, vbOKOnly + vbCritical, "Atenção!"
   End If
Saida:
   Screen.MousePointer = vbDefault
End Sub
Private Sub LoadMenuDefault()
   If gDebug Then MsgBox "LoadMenuDefault"
   On Error Resume Next
         
   If True Then
      If Not MdiMenu Is Nothing Then
         Call MdiMenu.LoadMenuDefault
      End If
   Else
      Dim MyObj As Object
      Set MyObj = CreateObject("SHORTBAR3R.TL_SHORTBAR")
      Set MyObj.Sys = Sys
      MyObj.Show
   End If

   If gDebug Then MsgBox "LoadMenuDefault End"
End Sub
Private Sub MontarToolbar()
   'Me.CommandBars.FindControl(Type, Id, Visible, Recursive)
'   Call MontarToolbarDinamico(Me)
End Sub
Private Sub MontarStatusBar()
'   Dim StatusBar As XtremeCommandBars.IStatusBar
   Dim StPane As StatusBarPane
   Dim dData  As Date
'   If Not Me.Visible Then Exit Sub
   
   With mvarMe.ProgressBar
      .Visible = True
      .UseVisualStyle = False
      .FlatStyle = Not .UseVisualStyle
      .BackColor = &H8000000F
      .BarColor = &H8000000C
      .Font.Bold = True
      .Text = ""
       
   End With
   
   Set StatusBar = mvarMe.CommandBars.StatusBar
   With StatusBar
      .RemoveAll
      .Visible = True
      Set StPane = .AddPane(101)  '* CONEXÃO
      With StPane
         .Style = 0
         .Text = "[" & Sys.xdb.Alias & "]"
         .Alignment = xtpAlignmentCenter
         .TextColor = IIf(Sys.xdb.Conectado, vbBlack, &H8000000C)
         .Tooltip = "[" & Sys.xdb.Server & "].[" & Sys.xdb.dbName & "]"
         .Width = 100 ' Len(.Text) * 8
      End With
      
      Set StPane = .AddPane(102)  '* USUARIO
      With StPane
         .Style = 0
         .Text = Sys.USER.IDUSU
         .Alignment = xtpAlignmentCenter
         .TextColor = vbBlack
         .Tooltip = Sys.USER.NMUSU
         .Width = Len(.Text) * 8 '* (Screen.Width / Me.ScaleWidth)
      End With
      
      Set StPane = .AddPane(103)  '* PERFIL/EQUIPE...
      With StPane
         .Style = 0
         .Text = "equipe"
         .Alignment = xtpAlignmentCenter
         .TextColor = vbBlack
         .Tooltip = "Equipe"
         .Width = Len(.Text) * 8
      End With
      
      Set StPane = .AddPane(104)   '* STRETCH
      With StPane
         .Style = SBPS_STRETCH
         .Text = ""
         .Alignment = xtpAlignmentCenter
         .TextColor = vbBlack
         .Tooltip = ""
         '.Width = Len(.Pane(0).Text) * (8 * (Screen.Width / Me.ScaleWidth))
         .Handle = mvarMe.ProgressBar.hwnd
      End With
      
      
      .AddPane 59137 'ID_INDICATOR_CAPS
      .AddPane 59138 'ID_INDICATOR_NUM
      .AddPane 59139 'ID_INDICATOR_SCRL
      
      If Sys.xdb.Conectado Then
         dData = Sys.xdb.Sysdate(3)
      Else
         dData = Now()
      End If
      Set StPane = .AddPane(108)  '* DATA
      
      With StPane
         .Style = 0
         .Text = Format(dData, "dd/mm/yyyy")
         .Alignment = xtpAlignmentCenter
         .TextColor = vbBlack
         .Tooltip = "Data do Servidor"
         .Width = Len(.Text) * 8
      End With
      Set StPane = .AddPane(109)  '* DATA
      With StPane
         .Style = 0
         .Text = Format(dData, "hh:mm")
         .Alignment = xtpAlignmentCenter
         .TextColor = vbBlack
         .Tooltip = "Hora do Servidor"
         .Width = Len(.Text) * 8
      End With
   End With
End Sub
Private Sub MontarPanes()
   Dim xPane As Pane
   Dim A As Pane
   Dim B As Pane
   Dim C As Pane
   
   Dim gPanes  As Integer
   gPanes = 2
   
   
   With mvarMe.DockingPaneManager
      .DestroyAll
      If gPanes = 1 Then
         Set A = .CreatePane(1, 150, 40, DockTopOf)
         A.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable

      ElseIf gPanes = 2 Then
            Set A = .CreatePane(1, 150, 120, DockLeftOf, Nothing)
            A.Tag = 1
            A.TabColor = vbRed
            A.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable
            
            Set B = .CreatePane(2, 700, 400, DockRightOf, A)
            B.Tag = 2
            B.TabColor = vbBlue
            B.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable
            
      ElseIf gPanes = 3 Then
         Set A = .CreatePane(1, 150, 120, DockLeftOf, Nothing)
         A.Tag = 1
         
         Set B = .CreatePane(2, 700, 400, DockRightOf, A)
         B.Tag = 2
         
         Set C = .CreatePane(3, 400, 100, DockBottomOf, B)
         C.Tag = 3
         
      ElseIf gPanes = 4 Then
      
      End If
      .Options.HideClient = True
      .PaintManager.ShowCaption = False
      
   End With
   mvarMe.CommandBars.RecalcLayout
End Sub
Public Sub LimpaInstaciaObj()
   If gDebug Then MsgBox "LimpaInstaciaObj"
   On Error GoTo TrataErro
         
   Set MdiMenu = Nothing
         
   Exit Sub
TrataErro:
   MsgBox "Limpa Instacia de Objetos" & vbNewLine & Err.Number & " - " & Err.Description, vbOKOnly + vbCritical, "Atenção!"
End Sub
Public Sub InstaciaObj()
   Dim TpErro  As String
   Dim Erro429 As Boolean
   
   If gDebug Then MsgBox "MyInstaciaObj..."
   
   On Error GoTo TrataErro
   
'   TpErro = "DSActive": Set DsAuto = CriarObjeto("DSACTIVE.AutoInstall")
'   TpErro = "DsDsr": Set DsDsr = CriarObjeto("DSACTIVE.DSR")
'   TpErro = "DSLOAD": Set DSLOAD = CriarObjeto("DSACTIVE.DS_LOAD")
'   TpErro = "DsMsg": Set DsMsg = CriarObjeto("DSACTIVE.MENSAGEM")
   If mvarSys Is Nothing Then
      TpErro = "SysA": Set Sys = CriarObjeto("SysA.SetA")
      With Sys
         .LocalReg = gLocalReg
         .CODSIS = gCODSIS
      End With
      Set mvarSys = Sys
      
   Else
      Set Sys = mvarSys
   End If
   TpErro = "CONEXAO": Set Splash = CriarObjeto("CONEXAO.Splash")
   TpErro = "MdiMenu": Set MdiMenu = CreateObject("Menu.ControlMenu")
   
   
   
   GoTo Saida
TrataErro:
   If Err = 429 Then
      If gDebug Then MsgBox "Não Criou Menu"
      Erro429 = True
      Resume Next
   Else
      MsgBox "Instacia de Objetos [" & TpErro & "]" & vbNewLine & Err.Number & " - " & Err.Description, vbOKOnly + vbCritical, "Atenção!"
   End If
Saida:
   If Erro429 Then
      Err = 429
      MsgBox "Instacia de Objetos [" & TpErro & "]" & vbNewLine & Err.Number & " - " & Err.Description, vbOKOnly + vbCritical, "Atenção!"
   End If
End Sub
Public Sub ExibeSenha(Optional pTrocaConexao As Boolean = False, Optional pxDb As Object, Optional pIDUSU As String)
   Dim bAux As Boolean
   Dim sAux As String
   Dim i As Integer
   Dim nAux As Integer
   
   On Error GoTo TrataErro

   If gDebug Then MsgBox "Em ExibeSenha"
   If Splash Is Nothing Then
      Set Splash = CriarObjeto("CONEXAO.Splash")
   End If
   If Splash Is Nothing Then
      MsgBox "Tela de senha não foi criada.", vbInformation
   Else
      If Not Sys Is Nothing Then
         Sys.ExePath = App.Path
      End If
      
      If gDebug Then MsgBox "Splash Not Nothing"
      With Splash
         Set .Sys = Sys
         .DebugSys = False
         .CODSIS = gCODSIS
         .Alias = gALIAS
         .dbTipo = gDBTipo
         .Server = gSERVER
         .dbName = gDBNAME
         .UID = gDBUSER
         .PWD = gDBPWD
         
         .ExeVersao = LoadVersion
         .ExePath = App.Path & "\"
         .ExeFile = App.EXEName
         If Not MyLoadPicture Is Nothing Then
            Set .ImgTitulo = MyLoadPicture
         End If
         
         If Sys.GetParam("Caption1", gCODSIS) <> "" Then gCaption1 = Sys.GetParam("Caption1", gCODSIS)
         If Sys.GetParam("Caption2", gCODSIS) <> "" Then gCaption2 = Sys.GetParam("Caption2", gCODSIS)
         If Sys.GetParam("Caption3", gCODSIS) <> "" Then gCaption3 = Sys.GetParam("Caption3", gCODSIS)
         .Caption1 = gCaption1
         .Caption2 = gCaption2
         .Caption3 = gCaption3
         
         DoEvents
         Screen.MousePointer = vbDefault
            
         If pTrocaConexao Then
            .IDUSU = Me.StatusBar.Pane(1).Text
         Else
            If Not .VerificaLicenca1 Then
               End
            End If
         End If
         gIDUSU = .IDUSU
         If gDebug Then MsgBox "gIDUSU=" & gIDUSU
         If gIDUSU = "" Then
            If gDebug Then MsgBox "Show"
            'Set MySplash = Nothing
            'If pTrocaConexao Then
               .xdb.SrvDesconecta
            'End If
            .Show
         Else
            If gDebug Then MsgBox "Conectar"
            .IDUSU = gIDUSU
            .Conectar
         End If
         
         Dim dHHIni As Date
         Dim nTimeout As Long
         nTimeout = 5 '*Segundos
         dHHIni = CDate(Format(Now, "hh:mm:ss"))

         If gDebug Then MsgBox "Conectado=" & .xdb.Conectado
         Do While Not .xdb.Conectado
            If .Cancelado Then
               If pTrocaConexao Then
                  .Conectar
                  If .xdb.Conectado Then
                     pIDUSU = .IDUSU
                  End If
                  Exit Do
               Else
                  End
               End If
            End If

            DoEvents
            If gIDUSU <> "" Then '*Se não exibir senha, verificar tempo de conexão
               If DateDiff("s", dHHIni, CDate(Format(Now, "hh:mm:ss"))) > nTimeout Then
                  .Alias = "PRODUCAO"
                  If InStr(UCase(.Server), "\SQLEXPRESS") <> 0 Then
                     .Server = Environ("COMPUTERNAME") & "\SQLEXPRESS"
                  Else
                     .Server = Environ("COMPUTERNAME")
                  End If
                  .dbTipo = 1
                  .dbName = gDBNAME
                  .UID = gDBUSER
                  .PWD = gDBPWD
                  If Not ExisteArquivo(.LocalReg) Then
                     Call .GerarLocalReg
                  End If
                  If Dir(Environ("Programfiles") & "\Microsoft SQL Server\", vbDirectory) <> "" Or Dir(Environ("Programfiles") & "(x86)\Microsoft SQL Server\", vbDirectory) <> "" Then
                     .Conectar
                  End If
                  If .Conectado Then
                     bAux = ExisteArquivo(Sys.PathSetup & gSetupFile)
                     If Not bAux Then
                        If ExisteArquivo(Environ("Programfiles") & "(x86)\ClasseA\Admin\Dll\" & gSetupFile) Then
                           Sys.PathSetup = Environ("Programfiles") & "(x86)\ClasseA\Admin\Dll\"
                        End If
                     End If
                     If bAux Then
                        Call WriteIniFile(Sys.PathSetup & gSetupFile, "Database Format", "SERVER", .Server)
                        If InStr(UCase(Sys.PathSetup), "(X86)") <> 0 Then
                           Call WriteIniFile(Sys.PathSetup & gSetupFile, "AutoInstall Files", "Path", Sys.PathSetup)
                        End If
                        If ExisteArquivo(Sys.PathSetup & "Bak\" & gSetupFile & ".zip") Then
                           Call Unzip(Sys.PathSetup & "Bak\", gSetupFile & ".zip", Sys.PathTmp, False)
                           If ExisteArquivo(Sys.PathTmp & gSetupFile) Then
                              sAux = Trim(ReadIniFile(Sys.PathTmp & gSetupFile, "AutoInstall Files", "Path"))
                              If sAux <> "" Then Call WriteIniFile(Sys.PathSetup & gSetupFile, "AutoInstall Files", "Path", sAux)
                              i = 1
                              sAux = Trim(ReadIniFile(Sys.PathTmp & gSetupFile, "AutoInstall Files", "Path" & i))
                              While sAux <> ""
                                 If sAux <> "" Then Call WriteIniFile(Sys.PathSetup & gSetupFile, "AutoInstall Files", "Path" & i, sAux)
                                 i = i + 1
                                 sAux = Trim(ReadIniFile(Sys.PathTmp & gSetupFile, "AutoInstall Files", "Path" & i))
                              Wend
                              
                              sAux = Trim(ReadIniFile(Sys.PathTmp & gSetupFile, Sys.CODSIS & " AutoInstall Files", "Path"))
                              If sAux <> "" Then Call WriteIniFile(Sys.PathSetup & gSetupFile, Sys.CODSIS & " AutoInstall Files", "Path", sAux)
                              i = 1
                              sAux = Trim(ReadIniFile(Sys.PathTmp & gSetupFile, Sys.CODSIS & " AutoInstall Files", "Path" & i))
                              While sAux <> ""
                                 If sAux <> "" Then Call WriteIniFile(Sys.PathSetup & gSetupFile, Sys.CODSIS & " AutoInstall Files", "Path" & i, sAux)
                                 i = i + 1
                                 sAux = Trim(ReadIniFile(Sys.PathTmp & gSetupFile, Sys.CODSIS & " AutoInstall Files", "Path" & i))
                              Wend
                           End If
                           
                        End If
                     End If
                  Else
                     If vbYes = ExibirPergunta("Tempo de conexão expirado!" & vbNewLine & "Deseja verificar sua conexão?") Then
                        'Set MySplash = Nothing
                        .ShowCon vbModal
                     'End With
                     Else
                        'MsgBox "Tempo de conexão expirado!", vbInformation, "Conexão"
                        End
                     End If
                  End If
               End If
            End If
         Loop
                  
         DoEvents
         Screen.MousePointer = vbHourglass

         If gDebug Then MsgBox "pTrocaConexao=" & pTrocaConexao & " Or Splash.xDb.Conectado=" & .xdb.Conectado
         If .xdb.Conectado Or pTrocaConexao Then
         If gDebug Then MsgBox "pTrocaConexao=" & pTrocaConexao & " Or Splash.xDb.Conectado=" & .xdb.Conectado
            If .Cancelado Then
               Set Sys.xdb = IIf(pTrocaConexao, .xdb, pxDb)
               gIDUSU = pIDUSU
            Else
               Set Sys.xdb = .xdb
               gIDUSU = .IDUSU
            End If
            
            If gDebug Then MsgBox "IDUSU : " & gIDUSU & vbNewLine & "CODSIS : " & gCODSIS
            Call AtualizaBD
            
            With Sys
               'Set .XDb = XDb
               .IDUSU = Trim(gIDUSU)
               .CODSIS = gCODSIS
               .LocalReg = gLocalReg
               '********
               '* Se pasta de Administração alterou re-executa Sistema.
               If gDebug Then MsgBox ".PathSetup : " & .PathSetup & vbNewLine & "ReadIniFile(gLocalReg, 'Setup', 'PATHSETUP', '') : " & ReadIniFile(gLocalReg, "Setup", "PATHSETUP", "") & vbNewLine & " Reexecuta : " & (.PathSetup <> ReadIniFile(gLocalReg, "Setup", "PATHSETUP", "") And .PathSetup <> "")
               If ResolvePathName(.PathSetup) <> ResolvePathName(ReadIniFile(gLocalReg, "Setup", "PATHSETUP", "")) And .PathSetup <> "" Then
                  sAux = ResolvePathName(.GetParam(pCODPARAM:="PATHSETUP", pCODSIS:="GLOBAL", Default:="", pDescricao:="Pasta de Administração"))
                  If Len(Trim(sAux)) > 1 Then
                     Call WriteIniFile(gLocalReg, "Setup", "PATHSETUP", sAux)
                  End If
                  Call .SetRegPathSetup(gLocalReg)
                  Main
                  Exit Sub
               End If
               
               If gIDUSU = "" Then
                  End
               End If
               
               nAux = ReadIniFile(gLocalReg, "Config", "LOJA")
               .GetIniVars pCODSIS:=gCODSIS, pIDLOJA:=nAux, pIniFile:=gSetupFile, pAppPath:=App.Path
               '.SaveIniVars
               
               Screen.MousePointer = vbHourglass
               If gDebug Then MsgBox "Antes MDI.Show"
'xxx               MontaMenu
'               MDI.Show
'               Set MySplash = Nothing
                  
               On Error Resume Next
               Set .DefaultIcon = mvarMe.Icon
               '               'Verifica visualizção da versão
               '               Screen.MousePointer = vbHourglass
               '               s_VerificaVisualVersao
               
            End With
         Else
            End
         End If
      End With
   End If
   Exit Sub
TrataErro:
   If gIDUSU <> "" Then
      gIDUSU = ""
'      Set MySplash = Nothing
      Splash.Show
      On Error Resume Next
   End If
   Call UnloadIni
   'Set XDb = Nothing
   Set Splash = Nothing
   
'MsgBox Err.Number & " - " & Err.Description   'ShowError("Sub Main()")
   Resume Next
End Sub
Private Sub UnloadIni()
   Dim n As Variant
   On Error Resume Next
   Unload Splash
   Set Splash = Nothing
   For Each n In Forms
      If UCase("FrmInicio") = UCase(n.Name) Or UCase("FrmSenha") = UCase(n.Name) Then
         Unload n
         Exit For
      End If
   Next
End Sub

