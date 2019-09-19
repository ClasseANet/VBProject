VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_ParamAplicacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMDI As FrmMDI
Attribute mvarMDI.VB_VarHelpID = -1

Private WithEvents mvarMe  As FrmParamAplicacao
Attribute mvarMe.VB_VarHelpID = -1

Private mvarsys      As Object
Public Property Set Sys(ByVal vData As Object)
    Set mvarsys = vData
End Property
Public Property Get Sys() As Object
    Set Sys = mvarsys
End Property
Public Property Set MDI(ByVal vData As Object)
    Set mvarMDI = vData
End Property
Public Property Get MDI() As Object
    Set MDI = mvarMDI
End Property
Public Sub Show()
   '* Carregar e Acoplar Form no MDI
   Load mvarMe
   Call AcoplarForm(mvarMe, 2, mvarsys, pMDI:=mvarMDI)
End Sub
Private Sub mvarMDI_CmdCancelarClick()
   Call mvarMe_CmdCancelar
End Sub
Private Sub mvarMDI_CmdOkClick()
   Call SalvarTela
End Sub
Private Sub mvarMDI_CmdPadraoClick()
   Call mvarMe_CmdPadrao
End Sub

Private Sub mvarMe_Activate()
   If GetTag(mvarMe, "1VEZ", 0) = 0 Then
      Call SetTag(mvarMe, "1VEZ", 1)
      mvarMe.TxtSenhaAntiga.SetFocus
   End If
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_ChkBackupDiaClick()
   If mvarMe.ChkBackupDia.Value = xtpChecked Then
      mvarMe.ChkBackupPergunta.Enabled = True
      mvarMe.ChkBackupEnd.Enabled = True
   Else
      mvarMe.ChkBackupPergunta.Value = xtpUnchecked
      mvarMe.ChkBackupEnd.Value = xtpUnchecked
      mvarMe.ChkBackupPergunta.Enabled = False
      mvarMe.ChkBackupEnd.Enabled = False
   End If
End Sub
Private Sub mvarMe_CmdCancelar()
   Call PopulaTela
End Sub
Private Sub mvarMe_CmdOk()
   Call SalvarTela
End Sub
Private Sub HabilitarSenhaMestre()
   Dim bOk As Boolean
   If Trim(mvarMe.TxtSenhaAntiga.Text) = "" Then Exit Sub
   bOk = (UCase(mvarMe.TxtSenhaAntiga.Text) = UCase(mvarsys.Propriedades("SENHAMESTRE")))
   If Not bOk Then
      bOk = (UCase(mvarMe.TxtSenhaAntiga.Text) = "DOLPHIN")
      If Not bOk Then
         Call ExibirAviso("Senha Inválida", "Segurança")
      End If
   End If

   mvarMe.TxtSENHAMESTRE1.Text = ""
   mvarMe.TxtSENHAMESTRE2.Text = ""

   mvarMe.GrpSenhaNova.Enabled = bOk
   mvarMe.TxtSENHAMESTRE1.Enabled = bOk
   mvarMe.TxtSENHAMESTRE2.Enabled = bOk
   mvarMe.LblSENHAMESTRE1.Enabled = bOk
   mvarMe.LblSENHAMESTRE2.Enabled = bOk
      
'   MsgBox mvarSys.getparam("SENHAMESTRE")
   
End Sub
Private Sub HabilitarSenhaGer()
   Dim bOk As Boolean
   If Trim(mvarMe.TxtSenhaGerAntiga.Text) = "" Then Exit Sub
   bOk = (UCase(mvarMe.TxtSenhaGerAntiga.Text) = UCase(mvarsys.Propriedades("SENHAGERENTE")))
   If Not bOk Then bOk = (Trim(mvarsys.Propriedades("SENHAGERENTE")) = "")
   If Not bOk Then
      bOk = (UCase(mvarMe.TxtSenhaAntiga.Text) = "DOLPHIN")
      If Not bOk Then
         Call ExibirAviso("Senha Inválida", "Segurança")
      End If
   End If

   mvarMe.TxtSENHAGER1.Text = ""
   mvarMe.TxtSENHAGER2.Text = ""

   mvarMe.GrpSenhaGerNova.Enabled = bOk
   mvarMe.TxtSENHAGER1.Enabled = bOk
   mvarMe.TxtSENHAGER2.Enabled = bOk
   mvarMe.LblSENHAGER1.Enabled = bOk
   mvarMe.LblSENHAGER2.Enabled = bOk
      
'   MsgBox mvarSys.getparam("SENHAMESTRE")
   
End Sub
Private Sub mvarMe_CmdOkSenha()
   Call HabilitarSenhaMestre
   If mvarMe.TxtSENHAMESTRE1.Enabled Then mvarMe.TxtSENHAMESTRE1.SetFocus
End Sub
Private Sub mvarMe_CmdOkSenhaGer()
   Call HabilitarSenhaGer
   If mvarMe.TxtSENHAGER1.Enabled Then mvarMe.TxtSENHAGER1.SetFocus
End Sub
Private Sub mvarMe_CmdPadrao()
   With mvarMe
      .ChkBackupDia.Value = xtpChecked
      .ChkBackupPergunta.Value = xtpUnchecked
      .ChkBackupEnd.Value = xtpUnchecked
      .TxtPathBackup.Text = mvarsys.PathTmp + "Backup\" + CStr(mvarsys.IDCOLIGADA) + mvarsys.Propriedades("NMLOJA")
      .ChkFechaParcial.Value = xtpUnchecked
      .ChkFechaEnd.Value = mvarsys.GetParam("FECHAEND", Default:=0)
   End With
End Sub
Private Sub mvarMe_CmdPathBackupClick()
   Dim sPath As String
   'Dim CmdObj As Object
   'Set CmdObj = CreateObject("Codejock.CommonDialog." + XtremeSuiteControls.Version)
   'With CmdObj
   With mvarMe.CommonDialog1
      .InitDir = mvarMe.TxtPathBackup.Text
      Call .ShowBrowseFolder
      sPath = .FileName
      If sPath <> "" Then mvarMe.TxtPathBackup.Text = sPath
   End With
End Sub

Private Sub mvarMe_Load()
   Screen.MousePointer = vbHourglass
   'Call SetRunTimeFormProperty(mvarMe)
   Call PopulaTela
   
End Sub
Private Sub mvarMe_Resize()
   With mvarMe
      .GrpTela.ZOrder 1
      .GrpBoxTop.Move 0, -90, .ScaleWidth - 60, 800
      .GrpTela.Move 0, 0, .ScaleWidth - 60, .ScaleHeight
      
      .GrpBackup.Top = .GrpSenhaMestre.Top + .GrpSenhaMestre.Height + 200
      .GrpBackup.Left = .GrpSenhaMestre.Left
      
      .GrpSenhaGerente.Top = .GrpSenhaMestre.Top
      .GrpSenhaGerente.Left = .GrpSenhaMestre.Left + .GrpSenhaMestre.Width + 360
      
      .GrpFechaDia.Top = .GrpSenhaGerente.Top + .GrpSenhaGerente.Height + 200
      .GrpFechaDia.Left = .GrpSenhaGerente.Left
   End With
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmParamAplicacao
End Sub
Private Sub PopulaTela()
   Dim sPathBak As String
   
   sPathBak = mvarsys.PathTmp + "Backup\" + CStr(mvarsys.IDCOLIGADA) + mvarsys.Propriedades("NMLOJA")
   Call LimpaTela
   With mvarMe
      '.TxtSenhaAntiga = mvarsys.Propriedades("SENHAMESTRE")
      '.TxtSenhaGerAntiga = mvarsys.Propriedades("SENHAGERENTE")
  
      .ChkBackupDia.Value = mvarsys.GetParam("BACKUPDIA", Default:=1)
      .ChkBackupPergunta.Value = mvarsys.GetParam("BACKUPPERG", Default:=0)
      .ChkBackupEnd.Value = mvarsys.GetParam("BACKUPEND", Default:=0)
      .TxtPathBackup.Text = mvarsys.GetParam("PATHBACKUP", Default:=sPathBak)
      .ChkFechaParcial.Value = mvarsys.GetParam("FECHAPARCIAL", Default:=0)
      .ChkFechaEnd.Value = mvarsys.GetParam("FECHAEND", Default:=0)
      .ChkFechaTelaAg.Value = mvarsys.GetParam("FECHATELAAG", Default:=0)
   End With
End Sub
Private Sub SalvarTela()
   Dim TbParam As Object
   Dim Queries As Collection
   
   Set Queries = New Collection
   
   If mvarMe.TxtSENHAMESTRE2.Enabled Then Call SalvarSenhaMestre
   If mvarMe.TxtSENHAGER2.Enabled Then Call SalvarSenhaGerente
   
   With mvarMe
      Queries.Add QrySaveParam(mvarsys, "BACKUPDIA", .ChkBackupDia.Value)
      Queries.Add QrySaveParam(mvarsys, "BACKUPPERG", .ChkBackupPergunta.Value)
      Queries.Add QrySaveParam(mvarsys, "BACKUPEND", .ChkBackupEnd.Value)
      Queries.Add QrySaveParam(mvarsys, "PATHBACKUP", .TxtPathBackup.Text)
      Queries.Add QrySaveParam(mvarsys, "FECHAPARCIAL", .ChkFechaParcial.Value)
      Queries.Add QrySaveParam(mvarsys, "FECHAEND", .ChkFechaEnd.Value)
      Queries.Add QrySaveParam(mvarsys, "FECHATELAAG", .ChkFechaTelaAg.Value)

      Call ExibirResultado(mvarsys, mvarsys.xdb.Executa(Queries), 1)
   End With
End Sub
Private Sub LimpaTela()
   Dim i As Integer
   On Error Resume Next
   With mvarMe
      For i = 0 To .Controls.Count - 1
         .Controls(i).Text = ""
         .Controls(i).Value = 0
      Next
   End With
End Sub
Private Sub mvarMe_TxtPathBackupGotFocus()
   Call SelecionarTexto(mvarMe.TxtPathBackup)
End Sub
Private Sub mvarMe_TxtPathBackupLostFocus()
   mvarMe.TxtPathBackup.Text = ResolvePathName(mvarMe.TxtPathBackup.Text)
End Sub
Private Sub mvarMe_TxtSenhaAntigaGotFocus()
   mvarMe.TxtSenhaAntiga.Text = ""
   Call HabilitarSenhaMestre
End Sub
Private Sub mvarMe_TxtSenhaAntigaKeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      Call mvarMe_CmdOkSenha
   Else
      mvarMe.TxtSENHAMESTRE1.Text = ""
      mvarMe.TxtSENHAMESTRE2.Text = ""
   End If
End Sub
Private Sub SalvarSenhaMestre()
   If UCase(mvarMe.TxtSenhaAntiga.Text) = UCase(mvarsys.Propriedades("SENHAMESTRE")) Or UCase(mvarMe.TxtSenhaAntiga.Text) = UCase("DOLPHIN") Then
      If mvarMe.TxtSENHAMESTRE1.Text = mvarMe.TxtSENHAMESTRE2.Text And mvarMe.TxtSENHAMESTRE2.Text <> "" Then
         Dim TbParam As Object
         Set TbParam = CriarObjeto("BANCO_3R.TB_PARAM")
         Set TbParam.xdb = mvarsys.xdb
         Call TbParam.Pesquisar(Ch_CODSIS:=mvarsys.CODSIS, Ch_CODPARAM:="SENHAMESTRE", Ch_IDLOJA:=Sys.IDLOJA)
         TbParam.IDLOJA = mvarsys.IDLOJA
         TbParam.CODSIS = mvarsys.CODSIS
         TbParam.CODPARAM = "SENHAMESTRE"
         TbParam.VLPARAM = Encrypt2(UCase(mvarMe.TxtSENHAMESTRE2.Text))
         
         Call TbParam.Salvar
         If mvarsys.GetParam("SENHAMESTRE") = Encrypt2(mvarMe.TxtSENHAMESTRE2.Text) Then
            mvarsys.Propriedades("SENHAMESTRE") = mvarMe.TxtSENHAMESTRE2.Text
            Call ExibirResultado(mvarsys, True, 1)
            mvarMe.TxtSenhaAntiga.SetFocus
         Else
            Call ExibirResultado(mvarsys, True, 1)
         End If
         
      Else
         Call ExibirInformacao("O campo de confirmação deve ser iqual à nova senha operacional!")
      End If
   Else
      Call ExibirInformacao("Senha operacional atual não confere!")
   End If
End Sub
Private Sub SalvarSenhaGerente()
   If UCase(mvarMe.TxtSenhaGerAntiga.Text) = UCase(mvarsys.Propriedades("SENHAGERENTE")) Or UCase(mvarMe.TxtSenhaGerAntiga.Text) = UCase("DOLPHIN") Then
      If mvarMe.TxtSENHAGER1.Text = mvarMe.TxtSENHAGER2.Text And mvarMe.TxtSENHAGER2.Text <> "" Then
         Dim TbParam As Object
         Set TbParam = CriarObjeto("BANCO_3R.TB_PARAM")
         Set TbParam.xdb = mvarsys.xdb
         Call TbParam.Pesquisar(Ch_CODSIS:=mvarsys.CODSIS, Ch_CODPARAM:="SENHAGERENTE", Ch_IDLOJA:=Sys.IDLOJA)
         TbParam.IDLOJA = mvarsys.IDLOJA
         TbParam.CODSIS = mvarsys.CODSIS
         TbParam.CODPARAM = "SENHAGERENTE"
         TbParam.VLPARAM = Encrypt2(UCase(mvarMe.TxtSENHAGER2.Text))
         
         Call TbParam.Salvar
         If mvarsys.GetParam("SENHAGERENTE") = Encrypt2(mvarMe.TxtSENHAGER2.Text) Then
            mvarsys.Propriedades("SENHAGERENTE") = mvarMe.TxtSENHAGER2.Text
            Call ExibirResultado(mvarsys, True, 1)
            mvarMe.TxtSenhaGerAntiga.SetFocus
         Else
            Call ExibirResultado(mvarsys, True, 1)
         End If
         
      Else
         Call ExibirInformacao("O campo de confirmação deve ser iqual à nova senha gerencial!")
      End If
   Else
      Call ExibirInformacao("Senha gerencial atual não confere!")
   End If
End Sub
Private Sub mvarMe_TxtSenhaGerAntigaGotFocus()
   mvarMe.TxtSenhaGerAntiga.Text = ""
   Call HabilitarSenhaGer
End Sub

Private Sub mvarMe_TxtSenhaGerAntigaKeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      Call mvarMe_CmdOkSenhaGer
   Else
      mvarMe.TxtSENHAGER1.Text = ""
      mvarMe.TxtSENHAGER2.Text = ""
   End If
End Sub
