VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_ShortBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private WithEvents mvarMe  As FrmShortBar
Attribute mvarMe.VB_VarHelpID = -1
Private WithEvents CmmdBars   As XtremeCommandBars.CommandBars
Attribute CmmdBars.VB_VarHelpID = -1
Private WithEvents StatusBar As XtremeCommandBars.StatusBar
Attribute StatusBar.VB_VarHelpID = -1

Private mvarSys As Object

Dim TLAnterior    As Object
Dim bReload As Boolean
Dim TLCalendario  As Object
Dim TLContato     As Object
Dim TLFinanceiro  As Object
Dim TLSuprimento  As Object
Dim TLTarefa      As Object
Dim TLDiario      As Object
Dim bDebug        As Boolean
Public Property Set Sys(ByVal vData As Object)
    Set mvarSys = vData
    Set CmmdBars = mvarSys.MDI.CommandBars
    Set StatusBar = mvarSys.MDI.StatusBar
    gDebug = (ReadIniFile(mvarSys.ExePath & "Debug.txt", "DEBUG", "SHORTBAR", "0") = "1")
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Private Sub CreateShortcutBar()
   Dim Item As ShortcutBarItem
   
   On Error GoTo TrataErro
   
If bDebug Then MsgBox "IniciarCalendario"
   Call IniciarCalendario
If bDebug Then MsgBox "IniciarContato"
   Call IniciarContato
If bDebug Then MsgBox "IniciarFinanceiro"
   Call IniciarFinanceiro
If bDebug Then MsgBox "IniciarSuprimento"
   Call IniciarSuprimento
If bDebug Then MsgBox "IniciarTarefa"
   Call IniciarTarefa
If bDebug Then MsgBox "IniciarDiario"
   Call IniciarDiario
   
   'mvarMe.ScbMain.AddImageList mvarMe.ImageList1
   'mvarMe.ScbMain.AddImageList mvarMe.ImageList24
   mvarMe.ScbMain.Icons = mvarMe.ImgShortcutBar.Icons
If bDebug Then MsgBox "AddItem 1"
   Set Item = mvarMe.ScbMain.AddItem(SHORTCUT.Calendario, "Calendario", TLCalendario.Pane.hwnd)
   'TLCalendario.BarItem = Item
   Item.IconId = 1
If bDebug Then MsgBox "AddItem 2"
   Set Item = mvarMe.ScbMain.AddItem(SHORTCUT.Contatos, "Contatos", TLContato.Pane.hwnd)
   'TLContato.BarItem = Item
   Item.IconId = 2
If bDebug Then MsgBox "AddItem 3"
   Set Item = mvarMe.ScbMain.AddItem(SHORTCUT.Tarefa, "Tarefa", TLTarefa.Pane.hwnd)
   Set TLTarefa.BarItem = Item
   Item.IconId = 5
If bDebug Then MsgBox "AddItem 4"
   Set Item = mvarMe.ScbMain.AddItem(SHORTCUT.Financeiro, "Caixa", TLFinanceiro.Pane.hwnd)
   Set TLFinanceiro.BarItem = Item
   Item.IconId = 3
If bDebug Then MsgBox "AddItem 5"
   Set Item = mvarMe.ScbMain.AddItem(SHORTCUT.Suprimento, "Estoque", TLSuprimento.Pane.hwnd)
   Set TLSuprimento.BarItem = Item
   Item.IconId = 4

If bDebug Then MsgBox "AddItem 6"
   Set Item = mvarMe.ScbMain.AddItem(SHORTCUT.Diario, "Diario", TLDiario.Pane.hwnd)
   Set TLDiario.BarItem = Item
   Item.IconId = 6


   'mvarMe.ScbMain.FindItem(SHORTCUT.Calendario).Selected = True
   Exit Sub
TrataErro:
   Call ExibirStop("[TL_ShortBar][CreateShortcutBar] " & vbNewLine & vbNewLine & Err & " - " & Error, "Atenção!")
End Sub
'Dim MyObj       As Object
'     If Not Sys.EXISTEOBJETO("SHORTBAR") Then
'         Set MyObj = Criarobjeto("SHORTBAR3R.TL_SHORTBAR")
'        Sys.OBJEToS.Add MyObj, "SHORTBAR"
'         Set MyObj.Sys = Sys
'      Else
'         Set MyObj = Sys.OBJEToS("SHORTBAR")
'      End If
'
'     MyObj.Show
'mvarSys.MDI.CommandBars.Icons = mvarMe.ImgShortcutBar.Icons
'Me.CommandBars.FindControl(Type, Id, Visible, Recursive)
Private Sub CreateToolBar()
   Dim ControlPopUp As CommandBarPopup
   Dim Control As CommandBarControl
   Dim ToolBar As CommandBar
   
   
   mvarSys.MDI.CommandBars.Icons = mvarMe.ImgToobar.Icons

   Set ToolBar = CriarToolbar(mvarSys, "Standard")
   With ToolBar
      .Customizable = False
      .Closeable = False
      .DefaultButtonStyle = xtpButtonIcon
      .EnableDocking xtpFlagHideWrap
   
      Set Control = CriarButtonToolbar(ToolBar, XTPControlType.xtpControlButton, 1, pCaption:="Fechar Caixa", pCategory:="BAR", pStyle:=xtpButtonIcon, pBeginGroup:=False, pIconId:=10001)
      If Not Control Is Nothing Then
         With Control
            .Parameter = SetTag(.Parameter, "VBSCRIPT", """UTILITARIO3R.TL_FecharDia""")
         End With
         .Position = 0
         .Visible = True
      End If
   
      Set Control = CriarButtonToolbar(ToolBar, XTPControlType.xtpControlButton, 2, pCaption:="Vendas", pCategory:="BAR", pStyle:=xtpButtonIcon, pBeginGroup:=False, pIconId:=10009)
      If Not Control Is Nothing Then
         With Control
            .Parameter = SetTag(.Parameter, "VBSCRIPT", """FINANC3R.TL_Vendas""")
         End With
         .Position = 0
         .Visible = True
      End If
   End With
   If (xVal(mvarSys.Propriedades("PONTO")) = 1) Then
      Set ToolBar = CriarToolbar(mvarSys, "RH")
      
      Set ControlPopUp = CriarButtonToolbar(ToolBar, XTPControlType.xtpControlSplitButtonPopup, 1, pCaption:="Registrar Hora", pCategory:="RH", pStyle:=xtpButtonIcon, pBeginGroup:=False, pIconId:=10010)
      If Not ControlPopUp Is Nothing Then
         ControlPopUp.Parameter = SetTag(ControlPopUp.Parameter, "VBSCRIPT", """CADASTRO3R.TL_CADRBATIDA""")
         
         Set Control = ControlPopUp.CommandBar.Controls.Add(xtpControlButton, 1, "Banco de Horas")
         If Not Control Is Nothing Then
            Control.Parameter = SetTag(Control.Parameter, "VBSCRIPT", """CADASTRO3R.TL_CADRBANCOHH""")
         End If
      End If
      
      Call DockBarRightOf("RH", "Standard", mvarSys)
      With ToolBar
         .Customizable = False
         .Closeable = False
         .DefaultButtonStyle = xtpButtonIcon
         .EnableDocking xtpFlagAlignTop + xtpFlagAlignRight
         .Position = xtpBarTop
         .Visible = True
      End With
   End If
End Sub
Public Sub Show(Optional Id As Long = 301) '* 301=SHORTCUT.Calendario
   Dim MyRs    As Object
   Dim Sql     As String
   Dim bExiste As Boolean
   
   Dim cWords  As Collection
   Dim n
   
   Dim nIDLOJA       As Integer
   Dim nIDCOLIGADA   As Integer
   Dim sNMCOLIGADA   As String
   
   With mvarSys
      Sql = "Select IDCOLIGADA"
      Sql = Sql & " From USUARIO"
      Sql = Sql & " Where IDUSU=" & SqlStr(.IDUSU)
      If .xDb.Abretabela(Sql, MyRs) Then
         If .IDUSU = "DIO" Then
            Sql = "Select IDCOLIGADA, NMCOLIGADA, TAG "
            Sql = Sql & " From COLIGADA "
            Sql = Sql & " Where IDCOLIGADA=" & ReadIniFile(.LocalReg, "Config", "COLIGADA", "1")
            Sql = Sql & " Order By IDCOLIGADA"
            bExiste = .xDb.Abretabela(Sql, MyRs)
            If Not bExiste Then
               Sql = "Select IDCOLIGADA, NMCOLIGADA, TAG "
               Sql = Sql & " From COLIGADA "
               bExiste = .xDb.Abretabela(Sql, MyRs)
            End If
         Else
            Sql = "Select IDCOLIGADA, NMCOLIGADA, TAG "
            Sql = Sql & " From COLIGADA "
            Sql = Sql & " Where IDCOLIGADA=" & xVal(MyRs("IDCOLIGADA") & "")
            Sql = Sql & " Order By IDCOLIGADA"
            bExiste = .xDb.Abretabela(Sql, MyRs)
         
         End If
      Else
         Sql = "Select IDCOLIGADA, NMCOLIGADA, TAG "
         Sql = Sql & " From COLIGADA "
         Sql = Sql & " Order By IDCOLIGADA"
         bExiste = .xDb.Abretabela(Sql, MyRs)
      End If
      If bExiste Then
         nIDCOLIGADA = xVal(MyRs("IDCOLIGADA") & "")
         sNMCOLIGADA = Trim(MyRs("NMCOLIGADA") & "")
         If sNMCOLIGADA = "XXX" Or sNMCOLIGADA = "" Then
            sNMCOLIGADA = Trim(InputBox("Entre com nome da unidade.", "Unidade"))
            If sNMCOLIGADA = "" Then
               ExibirInformacao "Nome inválido"
               Unload mvarMe
            Else
               Set cWords = GetWords(sNMCOLIGADA)
               sNMCOLIGADA = ""
               For Each n In cWords
                  n = UnFormat(CStr(n))
                  If Len(n) > 3 Then
                     sNMCOLIGADA = sNMCOLIGADA & " " & FormatarNome(CStr(n))
                  Else
                     sNMCOLIGADA = sNMCOLIGADA & " " & n
                  End If
               Next
           End If
           sNMCOLIGADA = UCase(Trim(UnFormat(sNMCOLIGADA)))
           If sNMCOLIGADA = "" Then
               ExibirInformacao "Nome inválido"
               Unload mvarMe
            Else
               Sql = "Update COLIGADA"
               Sql = Sql & " Set NMCOLIGADA=" & SqlStr(sNMCOLIGADA)
               Sql = Sql & " Where IDCOLIGADA=" & SqlNum(nIDCOLIGADA)
               If .xDb.Executa(Sql) Then
                  Sql = "Update OLOJA"
                  Sql = Sql & " Set NOME=" & SqlStr(sNMCOLIGADA)
                  Sql = Sql & " Where IDLOJA = ("
                  Sql = Sql & "  Select Min(O2.IDLOJA) "
                  Sql = Sql & "  From OLOJA O2 "
                  Sql = Sql & "  Where O2.IDCOLIGADA=" & SqlNum(nIDCOLIGADA)
                  Sql = Sql & "   )"
                  If .xDb.Executa(Sql) Then
                     Sql = "Select IDCOLIGADA, NMCOLIGADA, TAG "
                     Sql = Sql & " From COLIGADA "
                     Sql = Sql & " Order By IDCOLIGADA"
                     bExiste = .xDb.Abretabela(Sql, MyRs)
                  Else
                     ExibirInformacao "Nome não foi salvo."
                     Unload mvarMe
                  End If
               
               Else
                  ExibirInformacao "Nome não foi salvo."
                  Unload mvarMe
               End If
            End If
         End If
         If Trim(MyRs("TAG") & "") = "" Then
            Dim MyCad As Object
            Set MyCad = CriarObjeto("CADASTRO3R.TL_EMPRESA")
            Set MyCad.Sys = mvarSys
            Set MyCad = Nothing
         End If
      Else
         Call CadastrarEmpresa
         bExiste = .xDb.Abretabela(Sql, MyRs)
      End If
      If bExiste Then
         If ExisteItem(mvarSys.USER.GRUPOS, "SISTEMA") Or mvarSys.USER.IDGRUPO = "SISTEMA" Then
            .Propriedades("IDCOLIGADA") = xVal(MyRs("IDCOLIGADA") & "")
            .Propriedades("NMCOLIGADA") = Trim(MyRs("NMCOLIGADA") & "")
         Else
            nIDCOLIGADA = xVal(ReadIniFile(mvarSys.LocalReg, "Config", "COLIGADA", -1))
            
            Sql = "Select U.IDUSU, IsNull(C2.IDCOLIGADA,U.IDCOLIGADA) [IDCOLIGADA]"
            Sql = Sql & " , IsNull(C2.NMCOLIGADA,C.NMCOLIGADA) [NMCOLIGADA]"
            Sql = Sql & " From USUARIO U"
            Sql = Sql & " Left Join USUARIO_COLIGADA UC On U.IDUSU=UC.IDUSU"
            Sql = Sql & " Left Join COLIGADA C On C.IDCOLIGADA=U.IDCOLIGADA"
            Sql = Sql & " Left Join COLIGADA C2 On C2.IDCOLIGADA=UC.IDCOLIGADA"
            Sql = Sql & " Where U.IDUSU=" & SqlStr(mvarSys.IDUSU)
            Sql = Sql & " And " & nIDCOLIGADA & " In (C.IDCOLIGADA, C2.IDCOLIGADA)"
            bExiste = .xDb.Abretabela(Sql, MyRs)
            If bExiste Then
               .Propriedades("IDCOLIGADA") = xVal(MyRs("IDCOLIGADA") & "")
               .Propriedades("NMCOLIGADA") = Trim(MyRs("NMCOLIGADA"))
               Call WriteIniFile(mvarSys.LocalReg, "Config", "COLIGADA", MyRs("IDCOLIGADA"))
            Else
               Sql = " Select U.IDUSU"
               Sql = Sql & " , IsNull(Min(UC.IDCOLIGADA),U.IDCOLIGADA) [IDCOLIGADA]"
               Sql = Sql & " , IsNull(C2.NMCOLIGADA,C.NMCOLIGADA) [NMCOLIGADA]"
               Sql = Sql & " From USUARIO U"
               Sql = Sql & " Left Join USUARIO_COLIGADA UC On U.IDUSU=UC.IDUSU"
               Sql = Sql & " Left Join COLIGADA C On C.IDCOLIGADA=U.IDCOLIGADA"
               Sql = Sql & " Left Join COLIGADA C2 On C2.IDCOLIGADA=UC.IDCOLIGADA"
               Sql = Sql & " Where U.IDUSU=" & SqlStr(mvarSys.IDUSU)
               Sql = Sql & " Group By U.IDUSU, U.IDCOLIGADA, C.NMCOLIGADA, C2.NMCOLIGADA"
               bExiste = .xDb.Abretabela(Sql, MyRs)
               If bExiste Then
                  .Propriedades("IDCOLIGADA") = xVal(MyRs("IDCOLIGADA") & "")
                  .Propriedades("NMCOLIGADA") = Trim(MyRs("NMCOLIGADA") & "")
                  Call WriteIniFile(mvarSys.LocalReg, "Config", "COLIGADA", xVal(MyRs("IDCOLIGADA") & ""))
               Else
                  ExibirInformacao "Este usuário não possui coligada associada."
                  Sys.Propriedades("FCOMANDO") = "End"
                  Exit Sub
               End If
            End If
         End If
      End If
      
      nIDCOLIGADA = xVal(.Propriedades("IDCOLIGADA", True))
      nIDLOJA = xVal(.Propriedades("IDLOJA", True))
      If nIDCOLIGADA = 0 Then nIDCOLIGADA = xVal(.Propriedades("COLIGADA", True))
      If nIDLOJA = 0 Then nIDLOJA = xVal(.Propriedades("LOJA", True))
      If nIDCOLIGADA = 0 Then nIDCOLIGADA = 1
      If nIDLOJA = 0 Then nIDLOJA = 1
      'If Sys.xDb.Alias = ReadIniFile(gLocalReg, "Conection " & CStr(IdConection), "ALIAS") Then

      Sql = "Select * "
      Sql = Sql & " From OLOJA "
      Sql = Sql & " Where IDLOJA=" & SqlNum(nIDLOJA)
      Sql = Sql & " And IDCOLIGADA=" & SqlNum(nIDCOLIGADA)
      Sql = Sql & " Order By IDLOJA"
      bExiste = .xDb.Abretabela(Sql, MyRs)
      If Not bExiste Then
         nIDLOJA = 0
         Sql = "Select *"
         Sql = Sql & " From OLOJA "
         Sql = Sql & " Order By IDLOJA"
         If .xDb.Abretabela(Sql, MyRs) Then
            bExiste = True
            nIDLOJA = MyRs("IDLOJA") & ""
         End If
      End If
      
      If Not bExiste Then
         If nIDLOJA = 1 Then
            Call CadastrarEmpresa
            bExiste = .xDb.Abretabela(Sql, MyRs)
         Else
            Sql = "Select * "
            Sql = Sql & " From OLOJA "
            Sql = Sql & " Where IDLOJA=1"
            bExiste = .xDb.Abretabela(Sql, MyRs)
            If Not bExiste Then
               Call CadastrarEmpresa
               bExiste = .xDb.Abretabela(Sql, MyRs)
            End If
         End If
      End If
   
      If bExiste Then nIDLOJA = xVal(MyRs("IDLOJA") & "")
      If nIDLOJA = 0 Then
         ExibirInformacao "O Sistema não pode continuar sem o cadastro da Unidade."
         Exit Sub
      End If
      Call SetPropriedades(nIDLOJA, True)
      If Not bReload Then
         StatusBar.Pane(2).Text = mvarSys.Propriedades("NMLOJA")
         StatusBar.Pane(2).Width = Len(StatusBar.Pane(2).Text) * 7
         StatusBar.Pane(2).ToolTip = "Loja"
      End If
   End With
   Set MyRs = Nothing
         
   If mvarSys.Propriedades("RELOAD") = 1 Then
   'If mvarMe.Visible Then
      Call Reload(nIDLOJA)
      'bReload = True
      'Call mvarMe_ScbMainSelectedChanged(mvarMe.ScbMain.FindItem(SHORTCUT.Calendario))
   Else
      Load mvarMe
      Call TarefasIniciais
   End If
   
   If mvarSys.MDI.DockingPaneManager.Panes(1).Handle <> mvarMe.hwnd Then
      Call SetMDI(mvarMe.hwnd, mvarSys.MDI.hwnd)
      Call AcoplarForm(mvarMe, 1, mvarSys)
      mvarMe.ScbMain.FindItem(Id).Selected = True
   Else
      Id = mvarMe.ScbMain.Selected.Id
      mvarMe.ScbMain.FindItem(Id).Selected = True
   End If
End Sub
Private Sub TarefasIniciais()
   Dim NgUtil As Object
   Dim bLojaPadrao As Boolean
   
   If ReadIniFile(mvarSys.LocalReg, "Config", "LOJAPADRAO", "0") = 0 Then bLojaPadrao = True
   If mvarSys.IDLOJA = ReadIniFile(mvarSys.LocalReg, "Config", "LOJAPADRAO", "0") = 0 Then bLojaPadrao = True
   If mvarSys.xDb.Alias <> "PRODUCAO" Then bLojaPadrao = False
   
   If bLojaPadrao Then
      'If Day(Now) <= 5 Or (Day(Now) >= 16 And Day(Now) <= 20) Then
      Set NgUtil = CriarObjeto("UTILITARIO3R.NG_UTILITARIO")
      Set NgUtil.Sys = mvarSys
      Call NgUtil.F_BaixarNF_FTP(True)
   End If

End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmShortBar
End Sub
Private Sub Class_Terminate()
   On Error Resume Next
   Set mvarMe = Nothing
End Sub

Private Sub CmmdBars_CommandBarKeyDown(CommandBar As XtremeCommandBars.ICommandBar, KeyCode As Long, Shift As Integer)
KeyCode = KeyCode
End Sub

Private Sub CmmdBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
   Dim MyTela As Object
   
   DoEvents
   If Control.Category = "BAR" Then
      Select Case Control.Id
         Case 1
            '.Parameter = SetTag(.Parameter, "VBSCRIPT", """UTILITARIO3R.TL_FecharDia""")
           If gDebug Then MsgBox "CriarObjeto(UTILITARIO3R.TL_FecharDia)"
            Control.Parameter = SetTag(Control.Parameter, "VBSCRIPT", "")
            Set MyTela = CriarObjeto("UTILITARIO3R.TL_FecharDia")
            If Not MyTela Is Nothing Then
               Set MyTela.Sys = mvarSys
               
              If gDebug Then MsgBox "Set MyTela.CalControl = TLCalendario.TLEvento.CalControl"
               Set MyTela.CalControl = TLCalendario.TLEvento.CalControl
               Set MyTela.NgCal = TLCalendario.NgCal
              If gDebug Then MsgBox " MyTela.Show"
               MyTela.Show
               
               On Error Resume Next
               TLCalendario.TLEvento.CalControl.RedrawControl
               TLCalendario.TLEvento.CalControl.Populate
            End If
      End Select
   ElseIf UCase(Control.Caption) = "ATUALIZAÇÃO DO SISTEMA" Then
      'Call DataUpdate
   End If
End Sub
Private Sub DataUpdate()
   Dim nResult As Double
   
   Dim oFtp        As Object
   Dim FtpBak      As String
   Dim FtpBakUID   As String
   Dim FtpBakPWD   As String
   Dim FtpBakPath  As String
   
   Dim sLoja   As String
   Dim sArq    As String
   Dim sArqExe As String
   Dim LocPath As String
   Dim RemPath As String
   
   If IsWebConnected Then
      sLoja = Replace(mvarSys.Propriedades("NMLOJA"), " ", "")
      sArq = mvarSys.Propriedades("IDCOLIGADA") & sLoja & ".sql"
      If UCase(mvarSys.MDI.CommandBars.StatusBar(2).Text) <> UCase(sLoja) Then
         Exit Sub
      End If
   
      FtpBak = Decrypt2(mvarSys.GetParam("FTP", Default:="ftp.classeanet.com.br", pDescricao:="Servidor de Ftp"))
      FtpBakUID = mvarSys.GetParam("FtpUID", Default:="clientedpil", pDescricao:="Usuário do Ftp") 'freguesia
      FtpBakPWD = Decrypt2(mvarSys.GetParam("FtpPWD", Default:=Encrypt2(""), pDescricao:="Senha do Ftp")) 'dpil1010 tresrios10
      FtpBakPath = mvarSys.GetParam("FtpBakPath", Default:="", pDescricao:="Caminho do Ftp de Backup")
      FtpBakPath = "/" & sLoja & IIf(Mid(FtpBakPath, 1, 1) = "/", "", "/") & FtpBakPath
      
                           
      Set oFtp = CriarObjeto("VersaoFTP.TL_VerifVersao")
      With oFtp
         Set .Sys = mvarSys
         If .ConectarFTP(FtpBak, FtpBakUID, FtpBakPWD, False) Then
            'sLoja = mvarSys.Propriedades("NMLOJA")
            RemPath = mvarSys.GetParam("FtpBakPath", Default:="", pDescricao:="Caminho do Ftp de Backup")
            RemPath = "/" & sLoja & IIf(Mid(RemPath, 1, 1) = "/", "", "/") & RemPath
            LocPath = mvarSys.PathTmp
            
            '* Baixar Arquivo
            Call ExcluirArquivo(LocPath & sArq)
            If oFtp.BaixarArquivo(RemPath, sArq, LocPath, sArq) Then
               If ExisteArquivo(LocPath & sArq) Then
                  Call ExecuteScript(mvarSys.xDb, LocPath & sArq)
                  
                  sArqExe = Mid(sArq, 1, Len(sArq) - 4) & "_" & Format(Now(), "yymmdd_mmhh") & Right(sArq, 4)
                  Call CopiarArquivo(LocPath & sArq, LocPath & sArqExe)
                  Call ExcluirArquivo(LocPath & sArq)
                  Call oFtp.ExcluirArquivoRemoto(RemPath & sArq, False)
                  Call oFtp.EnviarArquivo(LocPath, sArqExe, RemPath, sArqExe, False)
                  
               End If
            End If
            oFtp.DesconectarFTP
            Set oFtp = Nothing
         End If
      End With
   End If
End Sub

Private Sub mvarMe_Activate()
If bDebug Then MsgBox "Activate"
   If GetTag(mvarMe, "1Vez", 1) = 1 Then
      Call SetTag(mvarMe, "1Vez", 0)
      Screen.MousePointer = vbDefault
   End If
   'mvarMe.Timer1.Enabled = bBiometria
   mvarSys.Propriedades("RELOAD") = "1"
   mvarMe.ScbMain.Visible = True
End Sub
Private Sub mvarMe_Terminate()
'   Call Terminate_Finger
End Sub

Private Sub mvarMe_Timer1Timer()
   Dim oObj As Object
   Dim bMeioDia As Boolean
   Dim bMeiaNoite As Boolean
   Dim bHora1 As Boolean
   Dim bHora2 As Boolean
   Static Dia As Integer
   
   Dim bProducao  As Boolean
   bProducao = bProducao Or (UCase(mvarSys.xDb.Alias) = "PRODUCAO")
   bProducao = bProducao Or (UCase(mvarSys.xDb.Alias) = "PRODUCÃO")
   bProducao = bProducao Or (UCase(mvarSys.xDb.Alias) = "PRODUÇAO")
   bProducao = bProducao Or (UCase(mvarSys.xDb.Alias) = "PRODUÇÃO")
   If Not bProducao Then
      mvarMe.Timer1.Enabled = False
      Exit Sub
   End If
         
         
   If GetAsyncKeyState(vbKeyF12) = -32767 Then
      Set oObj = CriarObjeto("CADASTRO3R.TL_CADRBATIDA")
      Set oObj.Sys = mvarSys
      Call oObj.Show
      Set oObj = Nothing
   End If
   bMeiaNoite = CDate(Format(Now(), "hh:mm")) = CDate("23:59")
   bMeioDia = CDate(Format(Now(), "hh:mm")) = CDate("12:00")

   bHora1 = CDate(Format(Now(), "hh:mm")) = CDate("10:00")
   bHora2 = CDate(Format(Now(), "hh:mm")) = CDate("16:00")
   
   If bMeioDia Or bMeiaNoite Or bHora1 Or bHora2 Then
      '* Backup e Fechar
      Set oObj = CriarObjeto("UTILITARIO3R.TL_FecharDia")
      Set oObj.Sys = mvarSys
      Call oObj.GerarBackup
      
      Call oObj.ArmazenarBackup
      Set oObj = Nothing
      
      If Dia <> Day(Now()) And Dia <> 0 Then
         Call TarefasIniciais
         Call TLCalendario.F_AcoesAgenda
      End If
      If bMeiaNoite Then Sys.Propriedades("FCOMANDO") = "End"
   End If
   Dia = Day(Now())
End Sub

Private Sub StatusBar_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
   Call SetTag(mvarSys.MDI, "nStatusButton", Button)
End Sub
Private Sub StatusBar_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
   On Error Resume Next
   Select Case xVal(GetTag(mvarSys.MDI, "nStatusButton", 0))
      Case 1 'Left
      Case 2 'Right
         Select Case mvarSys.MDI.StatusBar.HitTest(x, y).Id
           Case 101:   PaneConexao_Click    '* CONEXÃO
'           Case 102:   PaneUsusario_Click   '* USUARIO
           Case 103:   PaneEquipe_Click     '* PERFIL/EQUIPE...
'           Case 104:   PaneStrech_Click     '* STRETCH
'           Case 108:   PaneData_Click       '* DATA
'           Case 109:   PaneHora_Click       '* HORA
           
'           Case 59137: PaneCapsLock_Click   '* ID_INDICATOR_CAPS
'           Case 59138: PaneNum_Click        '* ID_INDICATOR_NUM
'           Case 59139: PaneScrl_Click       '* ID_INDICATOR_SCRL
         End Select
   End Select
End Sub
Private Sub PaneConexao_Click()
   On Error Resume Next
   mvarSys.IDLOJA = -1
   StatusBar.Pane(2).Text = ""
   
   If TLAnterior Is TLCalendario Then
      Call IniciarCalendario
      Call TLCalendario.MontarUnidades
      Call TLCalendario.RefreshCalendar
   End If
   
End Sub
Private Sub PaneEquipe_Click()
   Dim i       As Integer
   Dim Popup   As CommandBar
   Dim Control As CommandBarControl
   Dim nReturn As Integer
   Dim MyRs As Object
   Dim Sql As String
   Dim bExiste As Boolean
   Dim nAux As Integer
   Static nUlt As Integer
   
     
   On Error GoTo TrataErro
   nAux = Year(Now()) + Month(Now()) + Day(Now()) + Hour(Now()) + Minute(Now()) + Second(Now())
   If nAux - nUlt < 3 And nAux - nUlt > 0 Then
      Exit Sub
   End If
   nUlt = Year(Now()) + Month(Now()) + Day(Now()) + Hour(Now()) + Minute(Now()) + Second(Now())
   
   Set Popup = CmmdBars.Add("Popup", xtpBarPopup)
  
   Sql = "Select IDLOJA, NOME From OLOJA"
   If mvarSys.xDb.Abretabela(Sql, MyRs) Then
      While Not MyRs.EOF
         Set Control = Popup.Controls.Add(xtpControlButton, xVal(MyRs("IDLOJA") & ""), MyRs("NOME"))
         Control.Category = "POPUP_EQUIPE"
         Control.Parameter = "|CARREGADO=1|IDLOJA=" & CStr(MyRs("IDLOJA") & "") & "|"
         Control.Checked = (xVal(mvarSys.Propriedades("IDLOJA")) = MyRs("IDLOJA"))
         MyRs.MoveNext
      Wend
   End If
  
   nReturn = Popup.ShowPopup(TPM_RETURNCMD)
   'MsgBox "nReturn:" & mvarMe.paneequipe(nReturn).Parameter
   mvarSys.MDI.MousePointer = vbHourglass
   Screen.MousePointer = vbHourglass
   If nReturn > 0 And nReturn <> mvarSys.IDLOJA Then
      Call Reload(nReturn)
   End If
   
GoTo Saida
TrataErro:
   MsgBox Err & " - " & Error
Saida:
   Screen.MousePointer = vbDefault
   mvarSys.MDI.MousePointer = vbDefault
End Sub
Private Sub Reload(pIDLOJA As Integer)
   Dim iItem As Integer
   
   If mvarSys.IDLOJA = pIDLOJA Then Exit Sub
   
   bReload = True
   mvarSys.Propriedades("RELOAD") = 1
   If mvarMe.ScbMain.Selected Is Nothing Then
      iItem = SHORTCUT.Calendario
   Else
      iItem = mvarMe.ScbMain.Selected.Id
   End If
   
   Call SetPropriedades(pIDLOJA)
   With mvarSys
      StatusBar.Pane(2).Text = .Propriedades("NMLOJA")
      StatusBar.Pane(2).Width = Len(StatusBar.Pane(2).Text) * 7
      
      If mvarSys.MDI.DockingPaneManager.Panes(1).Handle <> mvarMe.hwnd Then
         Call SetMDI(mvarMe.hwnd, mvarSys.MDI.hwnd)
         Call AcoplarForm(mvarMe, 1, mvarSys)
      End If
   End With
   
   DoEvents
   mvarMe.ScbMain.RemoveAllItems
   Call CreateShortcutBar
   
   'mvarMe.ScbMain.FindItem(iItem).Selected = True
   'If Not mvarMe.ScbMain.FindItem(iItem) Is Nothing Then
   '   Call mvarMe_ScbMainSelectedChanged(mvarMe.ScbMain.FindItem(iItem))
   'End If
   
   Call IniciarCalendario
   Call TLCalendario.MontarUnidades
   mvarMe.ScbMain.FindItem(iItem).Selected = True
   
   mvarSys.Propriedades("RELOAD") = 0
   bReload = False
                 
   On Error Resume Next
   Dim vObj As Variant
   For Each vObj In mvarSys.Objetos
      Set vObj.Sys = mvarSys
      vObj.IDLOJA = mvarSys.Propriedades("IDLOJA")
   Next
   
End Sub
Private Sub SetPropriedades(pId As Integer, Optional bSet As Boolean = False)
   Dim Sql As String
   Dim MyRs As Object
   With mvarSys
      If .IDLOJA <> pId Or pId = 0 Or .IDLOJA Or bSet Then
         .IDLOJA = pId
         Sql = "Select Min(E.IDEST) [IDEST], Min(A.IDAGENDA) [IDAGENDA]"
         Sql = Sql & " From SEST_LOJA E, OAGENDA A "
         Sql = Sql & " Where E.IDLOJA=A.IDLOJA"
         Sql = Sql & " And A.IDLOJA=" & .IDLOJA
         If .xDb.Abretabela(Sql, MyRs) Then
            .Propriedades("IDAGENDA") = Val(MyRs("IDAGENDA") & "")
            .Propriedades("IDEST") = Val(MyRs("IDEST") & "")
         End If
         .Propriedades("MAXLENTEL") = .GetParam("MAXLENTEL", Default:=8, pDescricao:="Quantidade de números do campo Telefone", pIDLOJA:=pId)
         .Propriedades("LENIDVENDA") = .GetParam("LENIDVENDA", Default:=6, pDescricao:="Quantidade de dígitos da venda.", pIDLOJA:=pId)
         .Propriedades("SENHAMESTRE") = Decrypt2(.GetParam("SENHAMESTRE", Default:=Encrypt2("123"), pDescricao:="Senha Adm Loja", pIDLOJA:=pId))
         .Propriedades("SENHAGERENTE") = Decrypt2(.GetParam("SENHAGERENTE", Default:=Encrypt2(""), pDescricao:="Senha Gerente", pIDLOJA:=pId))
         .Propriedades("BIOMETRIA") = .GetParam("BIOMETRIA", Default:=0, pDescricao:="Uso de Biometria")
         .Propriedades("PONTO") = .GetParam("PONTO", Default:=0, pDescricao:="Uso de Ponto")
         
         .Propriedades("OCLIENTEUNICO") = .GetParam("OCLIENTEUNICO", Default:=0, pDescricao:="Cad. Cliente Único", pIDLOJA:=pId)
         .Propriedades("OCONTATOUNICO") = .GetParam("OCONTATOUNICO", Default:=0, pDescricao:="Cad. Contato Único", pIDLOJA:=pId)
         .Propriedades("RFUNCIONARIOUNICO") = .GetParam("OCLIENTEUNICO", Default:=0, pDescricao:="Cad. Funcionário Único", pIDLOJA:=pId)
         .Propriedades("FCCORRENTEUNICO") = .GetParam("FCCORRENTEUNICO", Default:=0, pDescricao:="Cad. Conta Único", pIDLOJA:=pId)
      End If
   End With
End Sub

Private Sub mvarMe_Load()
   bDebug = (ReadIniFile(mvarSys.LocalReg, "Config", "DEBUG", "0") = "1")
   Call CreateShortcutBar
If bDebug Then MsgBox "CreateToolbar"
   Call CreateToolBar
   mvarMe.ScbMain.Visible = False
   Screen.MousePointer = vbHourglass
If bDebug Then MsgBox "End Sub Load"
   'Call Init_Finger
   DoEvents
End Sub
Private Sub mvarMe_Resize()
   On Error Resume Next
   mvarMe.ScbMain.Move 50, 50, mvarMe.ScaleWidth - 50, mvarMe.ScaleHeight - 100
   'bBiometria = (xVal(mvarSys.Propriedades("BIOMETRIA")) = 1)
   'mvarMe.Timer1.Enabled = bBiometria
End Sub
Private Sub mvarMe_ScbMainSelectedChanged(ByVal Item As XtremeShortcutBar.IShortcutBarItem)
   DoEvents
   Dim n As Object
   Dim sAux As String
   Dim oToolBar As CommandBars

   On Error GoTo TrataErro

'   mvarMe_Resize
   Set oToolBar = mvarSys.MDI.CommandBars
   With oToolBar
      For Each n In .ActiveMenuBar.Controls 'mvarSys.MDI.CommandBars(1).Controls
         If GetTag(n.Parameter, "MENUCHILD", "") = "S" Then
            n.Visible = False
         End If
      Next
      For Each n In oToolBar
         If Mid(UCase(n.Title), 1, 4) = "BAR_" Then
            n.Visible = False
         End If
      Next
   End With


   mvarMe_Resize
   DoEvents
'   mvarSys.MDI.DockingPaneManager.Panes(2).Handle = 0
   DoEvents
   Select Case Item.Id
      Case SHORTCUT.Calendario
         Set TLAnterior = TLCalendario
         TLCalendario.Show bReload
      
      Case SHORTCUT.Tarefa
         If Not TLCalendario Is Nothing Then Set TLTarefa.TlCal = TLCalendario
         Set TLAnterior = TLTarefa
         TLTarefa.Show bReload
      
      Case SHORTCUT.Contatos
         Set TLAnterior = TLContato
         TLContato.Show bReload
      
      Case SHORTCUT.Financeiro
         Set TLAnterior = TLFinanceiro
         TLFinanceiro.Show bReload
         
      Case SHORTCUT.Suprimento
         Set TLAnterior = TLSuprimento
         TLSuprimento.Show bReload
      
      Case SHORTCUT.Diario
         Set TLAnterior = TLDiario
         TLDiario.Show bReload
      
   End Select
   bReload = False

   Exit Sub
TrataErro:
   Call ExibirStop("[TL_ShortBar][ScbMainSelectedChanged][Item.Id " & Item.Id & "]" & vbNewLine & vbNewLine & Err & " - " & Error, "Atenção!")
End Sub
Private Sub IniciarCalendario()
   On Error GoTo TrataErro
   
   If TLCalendario Is Nothing Then
      If bDebug Then MsgBox "Create"
      Set TLCalendario = CriarObjeto("Calendario3R.TL_CALENDARIO")
   End If
   If bDebug Then MsgBox "Sys"
   If Not TLCalendario Is Nothing Then Set TLCalendario.Sys = mvarSys
   If bDebug Then MsgBox "End"
   
   Exit Sub
TrataErro:
   Call ExibirStop("[TL_ShortBar][Calendario] " & vbNewLine & vbNewLine & Err & " - " & Error, "Atenção!")
End Sub
Private Sub IniciarContato()
   On Error GoTo TrataErro
   
   If TLContato Is Nothing Then Set TLContato = CriarObjeto("Contato3R.TL_CONTATO")
   If Not TLContato Is Nothing Then Set TLContato.Sys = mvarSys

   Exit Sub
TrataErro:
   Call ExibirStop("[TL_ShortBar][Contato] " & vbNewLine & vbNewLine & Err & " - " & Error, "Atenção!")
End Sub
Private Sub IniciarFinanceiro()
   On Error GoTo TrataErro
   
   If TLFinanceiro Is Nothing Then
       Set TLFinanceiro = CriarObjeto("Financ3R.NG_Financeiro")
       mvarSys.Propriedades("IDCONTA") = 0
    End If
   If Not TLFinanceiro Is Nothing Then Set TLFinanceiro.Sys = mvarSys
   
   Exit Sub
TrataErro:
   Call ExibirStop("[TL_ShortBar][Financeiro] " & vbNewLine & vbNewLine & Err & " - " & Error, "Atenção!")
End Sub
Private Sub IniciarSuprimento()
   On Error GoTo TrataErro
   
   If TLSuprimento Is Nothing Then Set TLSuprimento = CriarObjeto("Supri3R.NG_Suprimento")
   If Not TLSuprimento Is Nothing Then Set TLSuprimento.Sys = mvarSys
   
   Exit Sub
TrataErro:
   Call ExibirStop("[TL_ShortBar][Suprimento] " & vbNewLine & vbNewLine & Err & " - " & Error, "Atenção!")
End Sub
Private Sub IniciarTarefa()
   On Error GoTo TrataErro
   
   If TLTarefa Is Nothing Then Set TLTarefa = CriarObjeto("Tarefa3R.NG_Tarefa")
   If Not TLTarefa Is Nothing Then Set TLTarefa.Sys = mvarSys
   
   Exit Sub
TrataErro:
   Call ExibirStop("[TL_ShortBar][Tarefa] " & vbNewLine & vbNewLine & Err & " - " & Error, "Atenção!")
End Sub
Private Sub IniciarDiario()
   On Error GoTo TrataErro
   
   If TLDiario Is Nothing Then Set TLDiario = CriarObjeto("Diario3R.NG_Diario")
   If Not TLDiario Is Nothing Then Set TLDiario.Sys = mvarSys
   
   Exit Sub
TrataErro:
   Call ExibirStop("[TL_ShortBar][Diario] " & vbNewLine & vbNewLine & Err & " - " & Error, "Atenção!")
End Sub
Private Sub CadastrarEmpresa()
   Dim MyCad As Object
   Set MyCad = CriarObjeto("CADASTRO3R.TL_EMPRESA")
   With MyCad
      Set .Sys = Sys
      .Show
   End With
End Sub
