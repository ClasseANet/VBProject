VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DS_SEGURANCA"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"TB_ACCESS"
Attribute VB_Ext_KEY = "Member1" ,"TB_GRPUSER"
Attribute VB_Ext_KEY = "Member2" ,"TB_MODULO"
Attribute VB_Ext_KEY = "Member3" ,"TB_USER"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Const TabelaGRP = "GRPUSU"
Const CampoIDGRP = "IDGRUPO"
Const CampoDSRGRP = "DSCGRUPO"

Const TabelaUSU = "USUARIO"
Const CampoID = "IDUSU"
Const CampoGRP = "IDGRUPO"
Const CampoPWD = "SENHAUSU"
Const CampoNOME = "NMUSU"
Const CampoUsuCAD = "USUCADASTRO"



Private WithEvents mvarSeg As FrmSeg
Attribute mvarSeg.VB_VarHelpID = -1
Private WithEvents mvarSenha As FrmSenha
Attribute mvarSenha.VB_VarHelpID = -1
Private WithEvents mvarSenhaNova As FrmSenhaNova
Attribute mvarSenhaNova.VB_VarHelpID = -1

Private mvarDtMask As String
Private mvarSepData As String

Private mvarIDUSU As String
Private mvarPWDUSU As String
Private mvarIDGRUPO As String

Private mvarUsuSistema As String
Private mvarGrpSistema As String
Private mvarNmSistema As String
Private mvarPwdSistema As String

Private mvarMaster As String
Private mvarGrpMaster As String
Private mvarNmMaster As String
Private mvarPwdMaster As String

Private mvarAcesso As String

Private mvarIdioma As Integer
Private mvarFundoTela As String
Private mvarCODSYS As String

Private mvarRetorna As Boolean
Private mvarTentativas As Integer

Private mvarAplic As Object
Private mvarDBase As Object

Private mvarTB_GRUPO As Object
Private mvarTB_MODULO As Object
Private mvarTB_USUARIO As Object

Private Vet_Access
Public Property Get TB_USUARIO() As Object
   Set TB_USUARIO = mvarTB_USUARIO
End Property
Public Property Set TB_USUARIO(vData As Object)
    Set mvarTB_USUARIO = vData
End Property
Public Property Get TB_MODULO() As Object
    Set TB_MODULO = mvarTB_MODULO
End Property
Public Property Set TB_MODULO(vData As Object)
    Set mvarTB_MODULO = vData
End Property
Public Property Get TB_GRUPO() As Object
    Set TB_GRUPO = mvarTB_GRUPO
End Property
Public Property Set TB_GRUPO(vData As Object)
    Set mvarTB_GRUPO = vData
End Property
Public Sub Show()
   mvarSenha.Show vbModal
End Sub
Public Sub ShowNew()
   mvarSenhaNova.Show vbModal
End Sub
Public Sub ShowSeg()
   mvarSeg.Show vbModal
End Sub
Public Property Let CodSys(ByVal vData As String)
    mvarCODSYS = vData
End Property
Public Property Get CodSys() As String
    CodSys = mvarCODSYS
End Property
Public Property Let FundoTela(ByVal vData As String)
    mvarFundoTela = vData
End Property
Public Property Get FundoTela() As String
    FundoTela = mvarFundoTela
End Property
Public Property Let Idioma(ByVal vData As Integer)
    mvarIdioma = vData + eIdioma.PORTUGUES
End Property
Public Property Get Idioma() As Integer
    Idioma = mvarIdioma - eIdioma.PORTUGUES
End Property
Public Property Let Acesso(ByVal vData As String)
    mvarAcesso = vData
End Property
Public Property Get Acesso() As String
    Acesso = mvarAcesso
End Property
Public Property Let Tentativas(ByVal vData As Integer)
    mvarTentativas = vData
End Property
Public Property Get Tentativas() As Integer
    Tentativas = mvarTentativas
End Property
Public Property Let Retorna(ByVal vData As Boolean)
    mvarRetorna = vData
End Property
Public Property Get Retorna() As Boolean
    Retorna = mvarRetorna
End Property
Public Property Let PWDUSU(ByVal vData As String)
    mvarPWDUSU = vData
End Property
Public Property Get PWDUSU() As String
    PWDUSU = mvarPWDUSU
End Property
Public Property Let IDUSU(ByVal vData As String)
    mvarIDUSU = vData
End Property
Public Property Get IDUSU() As String
    IDUSU = mvarIDUSU
End Property
Public Property Let IDGRUPO(ByVal vData As String)
    mvarIDGRUPO = vData
End Property
Public Property Get IDGRUPO() As String
    IDGRUPO = mvarIDGRUPO
End Property
'Public Property Set DBase(ByVal vData As Object)
'    Set mvarDBase = vData
'End Property
Public Property Let Dbase(ByVal vData As Object)
    Set mvarDBase = vData
End Property
Public Property Get Dbase() As Object
    Set Dbase = mvarDBase
End Property
Public Property Let Aplic(ByVal vData As Object)
    Set mvarAplic = vData
End Property
Public Property Get Aplic() As Object
    Set Aplic = mvarAplic
End Property
Public Property Let SepData(ByVal vData As String)
    mvarSepData = vData
End Property
Public Property Get SepData() As String
    SepData = mvarSepData
End Property
Public Property Get UsuSistema() As String
    UsuSistema = mvarUsuSistema
End Property
Public Property Get GrpSistema() As String
    GrpSistema = mvarGrpSistema
End Property
Public Property Get PwdSistema() As String
    PwdSistema = mvarPwdSistema
    PwdSistema = ""
End Property
Public Property Get Master() As String
    Master = mvarMaster
End Property
Public Property Get GrpMaster() As String
    GrpMaster = mvarGrpMaster
End Property
Public Property Get NmMaster() As String
    NmMaster = mvarNmMaster
End Property
Public Property Get PwdMaster() As String
    PwdMaster = mvarPwdMaster
End Property
Private Sub Class_Initialize()
   mvarDtMask = "dd/MM/yyyy"             '* Máscara de Data Padrão
   mvarSepData = "/"                     '* Separador de data Padrão
   mvarTentativas = 3                    '* Número de Tentativas de Entrada
   
   Set mvarSenha = New FrmSenha          '* Tela de Senha
   Set mvarSenhaNova = New FrmSenhaNova  '* Tela de Nova Senha
   Set mvarSeg = New FrmSeg              '* Tela de Segurança
   
   mvarUsuSistema = "DIO"                '* Código do Usuário Sistema
   mvarGrpSistema = "SISTEMA"            '* Grupo do Usuário Sistema
   mvarNmSistema = "DIOGENES"            '* Senha do Usuário Sistema
   mvarPwdSistema = "DOLPHIN"            '* Senha do Usuário Sistema
   
   mvarMaster = "MAR"                    '* Código do Master
   mvarGrpMaster = "MASTER"                 '* Grupo do Master
   mvarNmMaster = "MASTER"               '* Nome do Master
   mvarPwdMaster = "SHIPP"               '* Senha do Master
   
   mvarIdioma = 5000 + 5000              '* Idoma Padrão Potuguês
End Sub
Private Sub Class_Terminate()
  Set mvarTB_USUARIO = Nothing
  Set mvarTB_MODULO = Nothing
  Set mvarTB_GRUPO = Nothing
  Set mvarSenha = Nothing
  Set mvarSenhaNova = Nothing
  Set mvarSeg = Nothing
End Sub
Private Sub mvarSenha_CmdNovaSenha()
   Dim Sql$
   On Error GoTo Fim
   Screen.MousePointer = vbHourglass
   If Len(mvarSenha.TxtUsuario) = 0 Then
      Call ClsMsg.ExibirAviso(ClsMsg.LoadMsg(63), mvarSenha.Caption)
      Call ClsCtrl.Set_Focus(mvarSenha.TxtUsuario)
      Exit Sub
'   ElseIf Len(mvarSenha.TxtSenha) = 0 Then
'      Call clsmsg.ExibirAviso(ClsMSG.LoadMsg(63), mvarSenha.Caption)
'      Call ClsCtrl.Set_Focus(mvarSenha.TxtSenha)
'      Exit Sub
   Else
      If mvarDBase.ExisteChave(TabelaUSU, CampoID, mvarSenha.TxtUsuario, eTipoDeDados.Caracter) Then
         mvarPWDUSU = mvarDBase.RSAux(CampoPWD)
         mvarIDUSU = mvarDBase.RSAux(CampoID)
         mvarIDGRUPO = mvarDBase.RSAux(CampoGRP)
      Else
         Call ClsMsg.ExibirAviso(ClsMsg.LoadMsg(58), ClsMsg.LoadMsg(57))
         Call ClsCtrl.Set_Focus(mvarSenha.TxtUsuario)
         Exit Sub
      End If
   End If
   'mvarDBase.Dys.Close
'   mvarIDUSU = ""
'   mvarPWDUSU = ""
'   mvarIDGRUPO = ""
   
'   mvarIDUSU = mvarDBase.RsAux(CampoID)
   If mvarDBase.RSAux(CampoPWD) = mvarSenha.TxtSENHA Then
      mvarPWDUSU = mvarDBase.RSAux(CampoPWD)
   Else
      Call ClsMsg.ExibirAviso(ClsMsg.LoadMsg(64), mvarSenha.Caption)
      Call ClsCtrl.Set_Focus(mvarSenha.TxtSENHA)
      Exit Sub
   End If
   mvarIDGRUPO = mvarDBase.RSAux(CampoGRP)
   
   DataInicioSistema = CVDate(mvarSenha.MskData.FormattedText)
   mvarRetorna = False
   Call ShowNew
   DoEvents
   If mvarRetorna Then
      Unload mvarSenha
   Else
      mvarSenha.TxtUsuario.SetFocus
   End If
      Screen.MousePointer = vbDefault
Exit Sub
Fim:
   If InStr(UCase(Error), "LOGIN FAILED") Then
      Call ClsMsg.ExibirAviso(ClsMsg.LoadMsg(7), ClsMsg.LoadMsg(57))
   End If
   Call ClsMsg.ShowError
End Sub
Private Sub mvarSenha_CmdOk()
    Screen.MousePointer = vbHourglass
    mvarTentativas = mvarTentativas - 1
    If Valida_Senha Then
       mvarIDUSU = UCase(mvarSenha.TxtUsuario.Text)
       Unload mvarSenha
    Else
       If mvarTentativas <= 0 Then
          Call ClsMsg.ExibirAviso(ClsMsg.LoadMsg(7), ClsMsg.LoadMsg(57))
          mvarRetorna = True
          Unload mvarSenha
       End If
    End If
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarSenha_CmdReturn()
    mvarRetorna = True
    Unload mvarSenha
End Sub
Private Sub mvarSenha_Load()
   Screen.MousePointer = vbHourglass
   Call F_DefIdiomaSenha
   Retorna = False
   mvarSenha.MskData = ClsDsr.Data_Padrao_Windows(Day(Now), Month(Now), Year(Now))
   If mvarIDUSU <> "CADASTRO" Then
      mvarSenha.TxtUsuario.Text = mvarIDUSU
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub mvarSenha_Timer()
   On Error Resume Next
   mvarSenha.Timer.Tag = CStr(Val(mvarSenha.Timer.Tag) + 1)
   If Val(mvarSenha.Timer.Tag) = 5 Then
      mvarIDUSU = "FIM"
      Unload mvarSenha
   End If
End Sub
Private Sub mvarSenha_Unload()
    If mvarRetorna Then
        mvarIDUSU = "FIM"
    End If
    Screen.MousePointer = vbDefault
End Sub
Friend Function Valida_Senha() As Boolean
   Dim Aux As Integer
   Dim Aux_senha As String
   Dim Txt As String, Titulo As String
   
   Dim NL
   
   Screen.MousePointer = vbDefault
   
   mvarSenha.TxtUsuario.Text = UCase(Trim(mvarSenha.TxtUsuario.Text))
   mvarSenha.TxtSENHA.Text = UCase(Trim(mvarSenha.TxtSENHA.Text))

   Valida_Senha = False
   cTitulo = ClsMsg.LoadMsg(57) '"Controle de Acesso"
   NL = vbLf
   
   '* ClsMSG.LoadMsg(27) = Campo inválido
   If Len(mvarSenha.MskData) = 0 Or Len(mvarSenha.MskData) <> 10 Or Not IsDate(mvarSenha.MskData.FormattedText) Then
      Txt = "( " & ClsDsr.EliminarString(mvarSenha.PnlDt, "&") & " )" & vbLf & vbLf & ClsMsg.LoadMsg(27)
      Call ClsMsg.ExibirAviso(Txt, cTitulo)
      Call ClsCtrl.Set_Focus(mvarSenha.MskData)
      Exit Function
   End If
   If Len(mvarSenha.TxtUsuario.Text) = 0 Then
      Txt = "( " & ClsDsr.EliminarString(mvarSenha.PnlUser, "&") & " )" & NL & NL & ClsMsg.LoadMsg(27)
      Call ClsMsg.ExibirAviso(Txt, cTitulo)
      mvarSenha.TxtUsuario = ""
      Call ClsCtrl.Set_Focus(mvarSenha.TxtUsuario)
      Exit Function
   End If
'   If Len(mvarSenha.TxtSenha.Text) = 0 Then
'      Txt = "( " & EliminarString(mvarSenha.PnlPwd, "&") & " )" & NL & NL & ClsMSG.LoadMsg(27)
'      Call clsmsg.ExibirAviso(Txt$, cTitulo)
'      mvarSenha.TxtSenha = ""
'      Call ClsCtrl.Set_Focus(mvarSenha.TxtSenha)
'      Exit Function
'   End If
   
   If Not Valida_DB Then
      Exit Function
   End If
   Valida_Senha = True
End Function
Friend Function Valida_DB() As Boolean
   Dim sGrupoAux     As String
   Dim SenhaSistema  As Boolean
   
   Screen.MousePointer = vbHourglass
   Valida_DB = False

   On Error GoTo Fim
  
   '***************
   '* Se Usuário for o Master ou o de Sistema
   '* não necessita Verificar a existência
   If ClsDsr.InArray(mvarSenha.TxtUsuario, Array(mvarUsuSistema, mvarMaster)) Then
      If Not mvarDBase.ExisteChave(TabelaUSU, CampoID, mvarSenha.TxtUsuario, eTipoDeDados.Caracter) Then
         Call Incluir_Usuario_Grupo_Sistema
      End If

      If ClsDsr.InArray(mvarSenha.TxtSENHA, Array(mvarPwdSistema, mvarPwdMaster)) Then
         mvarIDGRUPO = IIf(mvarPwdSistema = mvarSenha.TxtSENHA, mvarGrpSistema, mvarGrpMaster)
         Valida_DB = True
         Screen.MousePointer = vbDefault
         Exit Function
      End If
   End If
   
   '***************
   '* Verifica se o usuário existe
   If Not mvarDBase.ExisteChave(TabelaUSU, CampoID, mvarSenha.TxtUsuario, eTipoDeDados.Caracter) Then
'   If Not mvarTB_USUARIO.PESQUISAR(mvarSenha.TxtUsuario) Then
      Screen.MousePointer = vbDefault
      Call ClsMsg.ExibirAviso(ClsMsg.LoadMsg(58), ClsMsg.LoadMsg(57))
      Call ClsCtrl.Set_Focus(mvarSenha.TxtUsuario)
      Exit Function
   Else
      '***************
      '* Verifica se o senha está correta
      If mvarDBase.RSAux(CampoPWD) <> mvarSenha.TxtSENHA Then
         
         '***************
         '* Verifica se o senha Pertence ao usuário de Sistema
         sGrupoAux = mvarDBase.RSAux(CampoGRP) & ""
         If mvarDBase.ExisteChave(TabelaUSU, CampoPWD, mvarSenha.TxtSENHA, eTipoDeDados.Caracter) Then
            SenhaSistema = (mvarDBase.RSAux(CampoGRP) & "" = mvarGrpSistema)
         End If
         If SenhaSistema Then
            mvarIDGRUPO = sGrupoAux
         Else
            Screen.MousePointer = vbDefault
            Call ClsMsg.ExibirAviso(ClsMsg.LoadMsg(59), ClsMsg.LoadMsg(57))
            mvarSenha.TxtSENHA = ""
            Call ClsCtrl.Set_Focus(mvarSenha.TxtSENHA)
            Exit Function
         End If
      Else
         mvarIDGRUPO = mvarDBase.RSAux(CampoGRP) & ""
      End If
   End If
   
   mvarIDUSU = mvarSenha.TxtUsuario
   
   Valida_DB = True
   Screen.MousePointer = vbDefault
   Exit Function
Fim:
   If InStr(UCase(Error), "LOGIN FAILED") Then
      Call ClsMsg.ExibirAviso(ClsMsg.LoadMsg(7), ClsMsg.LoadMsg(57))
   End If
   Call ClsMsg.ShowError
End Function
Friend Sub F_DefIdiomaSenha()
   Dim IdiomaSenha As Long
   
'   IdiomaSenha = mvarIdioma + 5000
   
   With mvarSenha
      .Caption = ClsMsg.LoadMsg(mvarIdioma + 1)
      .PnlDt.Caption = ClsMsg.LoadMsg(mvarIdioma + 2)
      .PnlUser.Caption = ClsMsg.LoadMsg(mvarIdioma + 3)
      .PnlPwd.Caption = ClsMsg.LoadMsg(mvarIdioma + 4)
      .cmdOK.Caption = ClsMsg.LoadMsg(mvarIdioma + 5)
      .CmdReturn.Caption = ClsMsg.LoadMsg(mvarIdioma + 6)
      .CmdNovaSenha.Caption = ClsMsg.LoadMsg(mvarIdioma + 7)
      .Refresh
   End With
End Sub
Friend Sub F_DefIdiomaSenhaNova()
   Dim IdiomaSenha As Long
   
'   IdiomaSenha = mvarIdioma + 5000
   
   With mvarSenhaNova
      .Caption = ClsDsr.EliminarString(ClsMsg.LoadMsg(mvarIdioma + 7), ".") & " " & ClsMsg.LoadMsg(mvarIdioma + 1)
      .Caption = ClsDsr.EliminarString(.Caption, "&")
      .PnlOld.Caption = ClsMsg.LoadMsg(mvarIdioma + 1)
      .PnlNew.Caption = ClsMsg.LoadMsg(mvarIdioma + 8)
      .PnlVerify.Caption = ClsMsg.LoadMsg(mvarIdioma + 9)
      .CmdOper(0).Caption = ClsMsg.LoadMsg(mvarIdioma + 5)
      .CmdOper(1).Caption = ClsMsg.LoadMsg(mvarIdioma + 6)
      .Refresh
   End With
End Sub
Friend Function ValidaNew()
   Dim Sql$, Txt$
   ValidaNew = False
   If Len(mvarSenhaNova.TxtSENHA.Text) = 0 Then
      Call ClsMsg.ExibirAviso(ClsMsg.LoadMsg(64), ClsMsg.LoadMsg(57))
      Call ClsCtrl.Set_Focus(mvarSenhaNova.TxtSENHA)
      Exit Function
   End If
'   Sql$ = "select PWDUSER from TB_USER "
'   Sql$ = Sql$ & "where IDUSER ='" & mvarIDUSU & "'"
'   Sql$ = Sql$ & " and PWDUSER ='" & Trim(mvarSenhaNova.TxtSenha) & "'"
'   Call mvarDBase.AbreTabela(Sql)
   If Trim(mvarSenhaNova.TxtSENHA) <> mvarPWDUSU Then
'   If mvarDBase.CodeSql = NOTFOUND Then
      Call ClsMsg.ExibirAviso(ClsMsg.LoadMsg(59), ClsMsg.LoadMsg(57))
      Call ClsCtrl.Set_Focus(mvarSenhaNova.TxtSENHA)
      Screen.MousePointer = vbDefault
      Exit Function
   End If
   If Len(mvarSenhaNova.TxtNovaSenha.Text) = 0 Then
      Call ClsMsg.ExibirAviso(ClsMsg.LoadMsg(65), ClsMsg.LoadMsg(57))
      Call ClsCtrl.Set_Focus(mvarSenhaNova.TxtNovaSenha)
      Exit Function
   End If
   If Len(mvarSenhaNova.TxtConfirmacao) = 0 Then
      Call ClsMsg.ExibirAviso(ClsMsg.LoadMsg(64), ClsMsg.LoadMsg(57))
      Call ClsCtrl.Set_Focus(mvarSenhaNova.TxtConfirmacao)
      Exit Function
   ElseIf mvarSenhaNova.TxtNovaSenha <> mvarSenhaNova.TxtConfirmacao Then
      Call ClsMsg.ExibirAviso(ClsMsg.LoadMsg(59), ClsMsg.LoadMsg(57))
       mvarSenhaNova.TxtNovaSenha = ""
       mvarSenhaNova.TxtConfirmacao = ""
       Call ClsCtrl.Set_Focus(mvarSenhaNova.TxtNovaSenha)
       Exit Function
   End If
   ValidaNew = True
'   mvarDBase.Dys.Close
End Function

Private Sub mvarSenhaNova_ActiveSenhaNova()
  Screen.MousePointer = vbDefault
  If mvarPWDUSU <> "" Then
     mvarSenhaNova.TxtSENHA.Text = mvarPWDUSU
     mvarSenhaNova.TxtNovaSenha.SetFocus
  Else
     mvarSenhaNova.TxtSENHA.SetFocus
  End If
End Sub

Private Sub mvarSenhaNova_CmdOperNew(Index As Integer)
   Dim Sql$
   On Error GoTo Fim
   Select Case Index
      Case 0
         If ValidaNew() Then
            If mvarSenhaNova.TxtNovaSenha.Text <> mvarPWDUSU Then
               Sql$ = "update " & TabelaUSU
               Sql = Sql & " set " & CampoPWD & " ='" & Trim(mvarSenhaNova.TxtNovaSenha.Text) & "'"
               Sql = Sql & " where " & CampoID & " ='" & Trim(mvarIDUSU) & "'"
               Call mvarDBase.Executa(Sql)
            End If
            Unload mvarSenhaNova
         End If
      Case 1
         Unload mvarSenhaNova
   End Select
Exit Sub
Fim:
   Call ClsMsg.ShowError
   Unload mvarSenhaNova
End Sub
Private Sub mvarSeg_ChkAcessoClick(Index As Integer)
   Dim Acs$
   Dim i%, Grp%
   On Error GoTo Fim
   If Mid(mvarSeg.CmbGrpUser(1), 1, 3) = mvarGrpSistema Then
      mvarSeg.ChkAcesso(Index).Value = vbChecked
   End If
   For i = mvarSeg.ChkAcesso.LBound To mvarSeg.ChkAcesso.UBound
      If mvarSeg.ChkAcesso(i).Value = vbChecked Then
         Select Case i
            Case 0: Acs$ = Acs$ + "L"
            Case 1: Acs$ = Acs$ + "E"
            Case 2: Acs$ = Acs$ + "I"
            Case 3: Acs$ = Acs$ + "A"
         End Select
      End If
   Next
   Grp = Val(Mid(mvarSeg.CmbGrpUser(1), 1, 3))
   Vet_Access(Grp, mvarSeg.TreModu.SelectedItem.Index, 1) = Acs$
   mvarSeg.TreModu.SelectedItem.Tag = Acs$
   Exit Sub
Fim:
  On Error GoTo 0
  For i = mvarSeg.ChkAcesso.LBound To mvarSeg.ChkAcesso.UBound
      mvarSeg.ChkAcesso(i).Value = vbUnchecked
  Next
End Sub

Private Sub mvarSeg_CmbGrpUserClick(Index As Integer)
   Dim Sql$
   Dim i%, Grp%
   Select Case Index
      Case 1
         If mvarSeg.TreModu.Nodes.Count = UBound(Vet_Access, 2) Then
            For i = 2 To mvarSeg.TreModu.Nodes.Count
               Grp = Val(Mid(mvarSeg.CmbGrpUser(1), 1, 3))
               mvarSeg.TreModu.Nodes(i).Tag = Vet_Access(Grp, i, 1)
            Next
            AtualizaChkAcesso
         End If
   End Select
End Sub

Private Sub mvarSeg_CmdSegClick(Index As Integer)
   Screen.MousePointer = vbHourglass
   Select Case Index
      Case 0: Call F_SALVAR
      Case 1: Call F_EXCLUIR
      Case 2: Unload mvarSeg
   End Select
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarSeg_LblIdClick(Index As Integer)
   Dim Arrid, Tabela$
   Dim Cod As Control, Dsc As Control
   If Index = 2 Then Exit Sub
   mvarSeg.Suja = True
   Select Case Index
      Case 0
         Tabela$ = "TB_GRPUSU"
         Set Cod = mvarSeg.MskIdGrpUser
         Set Dsc = mvarSeg.TxtDscGrpUser
      Case 1
         Tabela$ = "TB_USU"
         Set Cod = mvarSeg.TxtIdUser
         Set Dsc = mvarSeg.TxtNmUser
    End Select
    Arrid = F_LOV(Tabela)
   '=======================
   If UBound(Arrid) < 0 Then Exit Sub
   '=======================
   Cod = Arrid(0)
   Cod.SetFocus
   Dsc.SetFocus
End Sub

Private Sub mvarSeg_Load()
   Dim Sql$
   ReDim Vet_Access(1, 1, 1)
'* Montar árvore de Módulos do programa
   Sql = "select IDMODU, DSCMODU, NVMODU "
   Sql = Sql & " from TB_MODUlO "
   Sql = Sql & " where SITMODU='A' "
   Sql = Sql & " and   CODSYS='" & mvarCODSYS & "'"
   Sql = Sql & " order by IDMODU"
   Call ClsCtrl.MontarTreeView(mvarDBase, mvarSeg.TreModu, Sql$, True)
'* Montar Combo de Grupo de usuários
   Sql = "select IDGRPUSER,DSCGRPUSU "
   Sql = Sql & " from TB_GRPUSU "
   Sql = Sql & " where IDGRPUSU >='" & mvarIDGRUPO & "'"
   Sql = Sql & " order by IDGRPUSU"
   Call ClsCtrl.MontarDbCombo(mvarDBase, mvarSeg.CmbGrpUser(0), Sql$, "DSCGRPUSU", "IDGRPUSU", True)
   Call ClsCtrl.MontarDbCombo(mvarDBase, mvarSeg.CmbGrpUser(1), Sql$, "DSCGRPUSU", "IDGRPUSU", True)
   If mvarSeg.CmbGrpUser(1).ListCount = 0 Then
      mvarSeg.CmbGrpUser(1).ListIndex = -1
   Else
      mvarSeg.CmbGrpUser(1).ListIndex = 0
   End If
'* Definir Banco de Dados Das tabelas utilizadas no Form
   On Error Resume Next
   TB_USER.Dbase = mvarDBase
   TB_GRPUSER.Dbase = mvarDBase
   TB_ACCESS.Dbase = mvarDBase
   mvarTB_GRPUSER.CodSys = mvarCODSYS
   mvarTB_ACCESS.CodSys = mvarCODSYS
'* Configurar Tela Padrão
   Call ClsCtrl.ConfigForm(mvarSeg, , mvarFundoTela)
   mvarSeg.Refresh
End Sub

Private Sub mvarSeg_MskIdGrpUserLostFocus()
   Dim Grp%
   Grp = Val(mvarSeg.MskIdGrpUser)
   If mvarSeg.TabSeg.Tab = 0 Then
      mvarSeg.MskIdGrpUser = ClsDsr.StrZero(CStr(Grp), 3)
      Call PopulaGrpUser
      mvarSeg.CmdSeg(1).Enabled = Not (Grp >= 0 And Grp <= 5)
   End If
End Sub
Private Sub mvarSeg_TabSegClick(PreviousTab As Integer)
   Dim Sql$
   On Error Resume Next
   mvarSeg.CmdSeg(0).Caption = "&Incluir"
   mvarSeg.CmdSeg(0).Enabled = True
   mvarSeg.CmdSeg(1).Enabled = True
   Select Case mvarSeg.TabSeg.Tab
      Case 0
         If mvarSeg.MskIdGrpUser.Enabled = False Then
            mvarSeg.TxtDscGrpUser.SetFocus
            Exit Sub
         End If
         mvarSeg.MskIdGrpUser.SetFocus
      Case 1
         Call ClsCtrl.MontarDbCombo(mvarDBase, mvarSeg.CmbGrpUser(0), Sql$, "DSCGRPUSU", "IDGRPUSU", True)
         If mvarSeg.TxtIdUser.Enabled = False Then
            mvarSeg.TxtNmUser.SetFocus
            Exit Sub
         End If
         mvarSeg.Frme(2).Refresh
         mvarSeg.Frme(3).Refresh
         Call ClsCtrl.Set_Focus(mvarSeg.TxtIdUser)
      Case 2
         Call ClsCtrl.MontarDbCombo(mvarDBase, mvarSeg.CmbGrpUser(1), Sql$, "DSCGRPUSU", "IDGRPUSU", True)
      '* Montar vetor de Acessos
         If mvarSeg.TreModu.Nodes.Count <> UBound(Vet_Access, 2) Then
            Vet_Access = mvarTB_ACCESS.MontaAcesso(Vet_Access)
            For i = 2 To mvarSeg.TreModu.Nodes.Count
               Grp = Val(Mid(mvarSeg.CmbGrpUser(1), 1, 3))
               mvarSeg.TreModu.Nodes(i).Tag = Vet_Access(Grp, i, 1)
            Next
            AtualizaChkAcesso
         End If
         mvarSeg.CmdSeg(0).Caption = "&Alterar"
         mvarSeg.CmdSeg(1).Enabled = False
   End Select
   Screen.MousePointer = vbDefault
   mvarSeg.Refresh
End Sub
Friend Function F_LOV(Tabela$)
   Dim Sql As Variant 'Coluna de ordenação (Padrao = 2 )
                      'ou Query de acesso
   Dim Cab, IdCampo, Tit$
   Dim MyClass As New LOV
   
   Select Case UCase(Tabela)
      Case "TB_USER"
         Cab = Array("Código", "IDUSER", 10, vbLeftJustify, _
                     "Usuário", "NMUSER", 30, vbLeftJustify)
         IdCampo = Array("IDUSER")
         Tit$ = "Usuários"
      Case "TB_GRPUSER"
         Cab = Array("Código", "IDGRPUSER", 6, vbLeftJustify, _
                      "Grupo", "DSCGRPUSER", 30, vbLeftJustify)
         IdCampo = Array("IDGRPUSER")
         Tit$ = "Grupos de Usuários"
      Case Else: Arrid = Array(""): Exit Function
   End Select
   With MyClass
'      DSR100.Sistema = UCase(AppName$ + " - " + AppTitle$)
      .Idioma = mvarIdioma
      .Aplic = mvarAplic
      .FundoTela = mvarFundoTela
      .Tipo = "LOV"
      
      .Versao = "XX" 'AppVer
      .Empresa = "EMPRESA" 'SYS_Nome_Empresa$
      
      .Dbase = mvarDBase
      .Table = UCase(Tabela)
'      .Query = Sql
      .Cab = Cab
      .IdField = IdCampo
           
      .Caption = Tit$
      .Show
      F_LOV = .Id
   End With
   Set MyClass = Nothing
End Function
Friend Sub PopulaUser()
   Dim Id$, CmD As Control
   If Trim(mvarSeg.TxtIdUser) = "" Then Exit Sub
   Set CmD = mvarSeg.CmdSeg(0)
   With mvarTB_USER
      Select Case .GetSelect(Trim(mvarSeg.TxtIdUser))
         Case ALTERACAO
            CmD.Caption = "&Alterar"
         '* Recupera grupo para controle de acesso
            Call mvarTB_GRPUSER.GetSelect(Trim(.IDGRPUSER))
            mvarSeg.CmbGrpUser(0).Visible = False
            mvarSeg.LblGrp = .IDGRPUSER & " - " & mvarTB_GRPUSER.DSCGRPUSER
            mvarSeg.CmbGrpUser(0).Visible = (mvarIDGRUPO <= .IDGRPUSER)
         '* Popula Tela
            mvarSeg.TxtIdUser = .IDUSER
            Call ClsCtrl.LocalizarCombo(mvarSeg.CmbGrpUser(0), .IDGRPUSER)
            If mvarIDGRUPO = mvarGrpSistema Then
               mvarSeg.TxtPwdUser(0).PasswordChar = ""
               mvarSeg.TxtPwdUser(1).PasswordChar = ""
            Else
               mvarSeg.TxtPwdUser(0).PasswordChar = "*"
               mvarSeg.TxtPwdUser(1).PasswordChar = "*"
            End If
            mvarSeg.TxtPwdUser(0) = .PWDUSER
            mvarSeg.TxtPwdUser(1) = .PWDUSER
            mvarSeg.TxtNmUser = .NMUSER
         Case INCLUSAO
            CmD.Caption = "&Incluir"
            Id$ = mvarSeg.TxtIdUser
            Call ClsDsr.LimparTela(mvarSeg)
            mvarSeg.TxtIdUser = Id$
         Case ERRO
      End Select
   End With
End Sub
Friend Sub PopulaGrpUser()
   Dim Id$, CmD As Control
   If Trim(mvarSeg.MskIdGrpUser) = "" Then Exit Sub
   Set CmD = mvarSeg.CmdSeg(0)
   With mvarTB_GRPUSER
      Select Case .GetSelect(Trim(mvarSeg.MskIdGrpUser))
         Case ALTERACAO
            CmD.Caption = "&Alterar"
            mvarSeg.MskIdGrpUser = .IDGRPUSER
            mvarSeg.TxtDscGrpUser = .DSCGRPUSER
         Case INCLUSAO
            CmD.Caption = "&Incluir"
            Id$ = mvarSeg.MskIdGrpUser
            Call ClsDsr.LimparTela(mvarSeg)
            mvarSeg.MskIdGrpUser = Id$
         Case ERRO
      End Select
   End With
End Sub
Friend Sub F_EXCLUIR()
  
   If Not VerificaAcesso(mvarAcesso, eEXCLUSAO) Then Exit Sub
   Select Case mvarSeg.TabSeg.Tab
      Case 0
         If mvarIDGRUPO > mvarSeg.MskIdGrpUser Then Exit Sub
         If mvarTB_GRPUSER.Excluir = vbYes Then Call ClsDsr.LimparTela(mvarSeg)
         Call ClsCtrl.Set_Focus(mvarSeg.MskIdGrpUser)
      Case 1
         If mvarIDGRUPO > Mid(mvarSeg.CmbGrpUser(0), 1, 3) Then Exit Sub
         If mvarTB_USER.Excluir = vbYes Then Call ClsDsr.LimparTela(mvarSeg)
         Call ClsCtrl.Set_Focus(mvarSeg.TxtIdUser)
      'Case 2
   End Select
End Sub
Friend Sub F_SALVAR()
   Dim Sql$, Id%
   If Not ValidaCampos Then Exit Sub
   Screen.MousePointer = vbHourglass
   Select Case mvarSeg.TabSeg.Tab
      Case 0: F_SALVAR_GRPUSER
      Case 1: F_SALVAR_USER
      Case 2: F_SALVAR_ACESSO
   End Select
   Screen.MousePointer = vbDefault
End Sub
Friend Sub F_SALVAR_ACESSO()
   Dim Sql$, Grp$, MODU$
   Dim i%, j%, Ind%
   Ind% = mvarSeg.CmbGrpUser(1).ListIndex
   
   If mvarIDUSU <> mvarUsuSistema Then If Not VerificaAcesso(mvarAcesso, eALTERACAO) Then Exit Sub

   For j% = 0 To mvarSeg.CmbGrpUser(1).ListCount - 1
      mvarSeg.CmbGrpUser(1).ListIndex = j
      Grp$ = Mid(mvarSeg.CmbGrpUser(1), 1, 3)
      For i = 2 To mvarSeg.TreModu.Nodes.Count
         If mvarIDGRUPO < Grp Or Grp = mvarGrpSistema Or Grp = mvarGrpMaster Then
            With mvarTB_ACCESS
               .IDMODU = mvarSeg.TreModu.Nodes(i).Key
               .CODACCESS = IIf((Grp = mvarGrpSistema Or Grp = mvarGrpMaster), "LEIA", mvarSeg.TreModu.Nodes(i).Tag)
               .CodSys = mvarCODSYS
               .IDGRPUSER = Grp$
               .Alterar
            End With
         End If
      Next
   Next
'* Recuperar Acesso da própria Tela
'   If mvarTB_ACCESS.GetSelect(mvarCODSYS, mvarIDGRUPO, "0400") <> FOUND Then Exit Sub
'   mvarAcesso = mvarTB_ACCESS.CODACCESS
'**********************************
   
   mvarSeg.CmbGrpUser(1).ListIndex = Ind%
   mvarSeg.Suja = False
'   Call clsmsg.ExibirAviso(LoadMsg(34), ClsMSG.LoadMsg(57))
End Sub
Friend Sub F_SALVAR_USER()
   Dim Sql$, Cod$, Grp$, Id%
   If mvarIDGRUPO > Mid(mvarSeg.LblGrp, 1, 3) Then
      Call ClsMsg.ExibirStop(ClsMsg.LoadMsg(7), ClsMsg.LoadMsg(57))
      Exit Sub
   End If
   
   With mvarTB_USER
      If Not VerificaAcesso(mvarAcesso, .Existe) Then Exit Sub
      .IDUSER = mvarSeg.TxtIdUser
      .IDGRPUSER = Mid(mvarSeg.CmbGrpUser(0), 1, 3)
      .NMUSER = mvarSeg.TxtNmUser
      .PWDUSER = mvarSeg.TxtPwdUser(0)
      .LOGON = mvarIDUSU
      .Gravar
      mvarSeg.Suja = False
      If .Existe = INCLUSAO Then
         Call ClsDsr.LimparTela(mvarSeg)
      End If
   End With
   Call ClsCtrl.Set_Focus(mvarSeg.TxtIdUser)
End Sub
Friend Sub F_SALVAR_GRPUSER()
   Dim Sql$, Cod$, Grp$, Id%

   If mvarIDGRUPO > mvarSeg.MskIdGrpUser Then Exit Sub
   
   With mvarTB_GRPUSER
      If Not VerificaAcesso(mvarAcesso, .Existe) Then Exit Sub
      .IDGRPUSER = mvarSeg.MskIdGrpUser
      .DSCGRPUSER = mvarSeg.TxtDscGrpUser
      .Gravar mvarCODSYS
      mvarSeg.Suja = False
      If .Existe = INCLUSAO Then
         Id% = Val(mvarSeg.MskIdGrpUser) + 1
         Call ClsDsr.LimparTela(mvarSeg)
         mvarSeg.MskIdGrpUser = ClsDsr.StrZero(Id%, 3)
      End If
   End With
   Call ClsCtrl.Set_Focus(mvarSeg.MskIdGrpUser)
End Sub
Private Sub mvarSeg_TreModuNodeClick(ByVal Node As ComctlLib.Node)
   Call AtualizaChkAcesso
End Sub
Private Sub mvarSeg_TxtIdUserLostFocus()
   Call PopulaUser
End Sub
Private Sub mvarSeg_Unload(Cancel As Integer)
    Set mvarTB_GRPUSER = Nothing
    Set mvarTB_USER = Nothing
    Set mvarTB_ACCESS = Nothing
End Sub
Friend Sub AtualizaChkAcesso()
   Dim Acs$
   Dim i%
   On Error Resume Next
   Acs$ = mvarSeg.TreModu.SelectedItem.Tag
   For i = mvarSeg.ChkAcesso.LBound To mvarSeg.ChkAcesso.UBound
      mvarSeg.ChkAcesso(i).Value = vbUnchecked
      mvarSeg.ChkAcesso(i).Enabled = True
   Next
   If mvarSeg.TreModu.SelectedItem.Index = 1 Then
      For i = mvarSeg.ChkAcesso.LBound To mvarSeg.ChkAcesso.UBound
         mvarSeg.ChkAcesso(i).Value = vbChecked
         mvarSeg.ChkAcesso(i).Enabled = False
      Next
      Exit Sub
   End If
   If Mid(mvarSeg.CmbGrpUser(1), 1, 3) = mvarGrpSistema Then Acs$ = "LEIA"
   If InStr(Acs$, "L") <> 0 Then mvarSeg.ChkAcesso(0).Value = vbChecked
   If InStr(Acs$, "E") <> 0 Then mvarSeg.ChkAcesso(1).Value = vbChecked
   If InStr(Acs$, "I") <> 0 Then mvarSeg.ChkAcesso(2).Value = vbChecked
   If InStr(Acs$, "A") <> 0 Then mvarSeg.ChkAcesso(3).Value = vbChecked
End Sub
Friend Function ValidaCampos()
   ValidaCampos = True
   Select Case mvarSeg.TabSeg.Tab
      Case 0
         If Val(mvarSeg.MskIdGrpUser) = 0 Or Trim$(mvarSeg.TxtDscGrpUser) = "" Then
            Call ClsMsg.ExibirAviso(ClsMsg.LoadMsg(27), ClsMsg.LoadMsg(1))
            If Val(mvarSeg.MskIdGrpUser) = 0 Then Call ClsCtrl.Set_Focus(mvarSeg.MskIdGrpUser)
            If Trim$(mvarSeg.MskIdGrpUser) = "" Then Call ClsCtrl.Set_Focus(mvarSeg.TxtDscGrpUser)
            Exit Function
         End If
      Case 1
         If Trim(mvarSeg.TxtIdUser) = "" Or Trim$(mvarSeg.TxtNmUser) = "" Or Trim$(mvarSeg.TxtPwdUser(0)) = "" Then
            Call ClsMsg.ExibirAviso(ClsMsg.LoadMsg(27), ClsMsg.LoadMsg(1))
            If Trim$(mvarSeg.TxtIdUser) = "" Then Call ClsCtrl.Set_Focus(mvarSeg.TxtIdUser)
            If Trim$(mvarSeg.TxtNmUser) = "" Then Call ClsCtrl.Set_Focus(mvarSeg.TxtNmUser)
            If Trim$(mvarSeg.TxtPwdUser(0)) = "" Then Call ClsCtrl.Set_Focus(mvarSeg.TxtPwdUser(0))
            Exit Function
         End If
   End Select
   ValidaCampos = True
End Function
Public Function VerificaAcesso(pAcesso$, pOper%)
   Dim Ac$
   If mvarIDGRUPO = mvarGrpSistema Then
      VerificaAcesso = True
      Exit Function
   End If
   Select Case pOper
      Case LEITURA: Ac$ = "L"
      Case EXCLUSAO: Ac$ = "E"
      Case INCLUSAO: Ac$ = "I"
      Case ALTERACAO: Ac$ = "A"
   End Select
   VerificaAcesso = (InStr(pAcesso$, Ac$) <> 0)
   If ClsDsr.InArray(Ac$, Array("I", "A")) Then VerificaAcesso = (InStr(pAcesso$, "L") <> 0)
   If Not VerificaAcesso Then Call ClsMsg.ExibirStop(ClsMsg.LoadMsg(7), ClsMsg.LoadMsg(57))
End Function
Friend Sub Incluir_Usuario_Grupo_Sistema()
   Dim Sql As String
  
   On Error Resume Next
   
   '*********
   '* Insere Usuário de Sistema
   If Not mvarDBase.ExisteChave(TabelaUSU, CampoID, mvarUsuSistema, eTipoDeDados.Caracter) Then
      '*********
      '* Insere Grupo de Sistema
      If Not mvarDBase.ExisteChave(TabelaGRP, CampoIDGRP, mvarGrpSistema, eTipoDeDados.Caracter) Then
         Sql = "insert into " & TabelaGRP
         Sql = Sql & "(" & CampoIDGRP
         Sql = Sql & "," & CampoDSRGRP
         Sql = Sql & ")"
         Sql = Sql & " values "
         Sql = Sql & " ('SISTEMA'"
         Sql = Sql & ",'USUARIO DE SISTEMA'"
         Sql = Sql & ")"
         Call mvarDBase.Executa(Sql, True)
      End If
   
      Sql = "insert into " & TabelaUSU
      Sql = Sql & " (" & CampoID
      Sql = Sql & "," & CampoGRP
      Sql = Sql & "," & CampoPWD
      Sql = Sql & "," & CampoNOME
      Sql = Sql & "," & CampoUsuCAD
      Sql = Sql & ")"
      Sql = Sql & " values "
      Sql = Sql & " ('" & mvarUsuSistema & "'"
      Sql = Sql & ",'" & mvarGrpSistema & "'"
      Sql = Sql & ",'" & mvarPwdSistema & "'"
      Sql = Sql & ",'" & mvarNmSistema & "'"
      Sql = Sql & ",'" & mvarUsuSistema & "'"
      Sql = Sql & ")"
      Call mvarDBase.Executa(Sql, True)
   
      '*********
      '* Insere Usuário Master
      If Not mvarDBase.ExisteChave(TabelaUSU, CampoID, mvarMaster, eTipoDeDados.Caracter) Then
         '*********
         '* Insere Grupo Master
         If Not mvarDBase.ExisteChave(TabelaGRP, CampoIDGRP, "MASTER", eTipoDeDados.Caracter) Then
            Sql = "insert into " & TabelaGRP
            Sql = Sql & "(" & CampoIDGRP
            Sql = Sql & "," & CampoDSRGRP
            Sql = Sql & ")"
            Sql = Sql & " values "
            Sql = Sql & " ('MASTER'"
            Sql = Sql & ",'USUARIO MASTER'"
            Sql = Sql & ")"
            Call mvarDBase.Executa(Sql, True)
         End If
         
         Sql = "insert into " & TabelaUSU
         Sql = Sql & "(" & CampoID
         Sql = Sql & "," & CampoGRP
         Sql = Sql & "," & CampoPWD
         Sql = Sql & "," & CampoNOME
         Sql = Sql & "," & CampoUsuCAD
         Sql = Sql & ")"
         Sql = Sql & " values "
         Sql = Sql & " ('" & mvarMaster & "'"
         Sql = Sql & ",'" & mvarGrpMaster & "'"
         Sql = Sql & ",'" & mvarPwdMaster & "'"
         Sql = Sql & ",'" & mvarNmMaster & "'"
         Sql = Sql & ",'" & mvarMaster & "'"
         Sql = Sql & ")"
         Call mvarDBase.Executa(Sql, True)
      End If
   End If
End Sub
Public Sub RecuperarMenu(SysMdi As Object)
Attribute RecuperarMenu.VB_Description = "Recupera o menu do sistema exceto o de nome ""MNUMOUSE"" e grava-os na tabela ""TB_MODULOS"""
   Dim Sql$, Cod$, Nome$, Nivel$, Grp$
   Dim i%, j%
   Dim RcAcesso As Recordset, RcModu As Recordset
   
'* Abrir Tabela de Grupo de Ususarios
'   Sql = "select IDGRPUSER from TB_GRPUSER"
'   Call mvarDBase.AbreTabela(Sql, RcAcesso)
'   If Not mvarDBase.CodeSql Then Exit Sub
   
   '* Tornar Inativo
   Sql = "update TB_MODULO set SITMODU='I'"
   Sql = Sql & " where CODSYS='" & UCase(mvarCODSYS) & "'"
   Call mvarDBase.Executa(Sql)
   For i = 0 To SysMdi.Controls.Count - 1
      If TypeName(SysMdi.Controls(i)) = "Menu" And UCase(Mid(SysMdi.Controls(i).Name, 1, 8)) <> "MNUMOUSE" Then
         Nome = ClsDsr.EliminarString(SysMdi.Controls(i).Caption, "&")
         Nivel = CStr(Len(Mid(SysMdi.Controls(i).Name, 4)) / 2)
         If Val(Nivel) = 0 Then
            Cod = ClsDsr.StrZero(SysMdi.Controls(i).Index, 2)
         Else
            Cod = Mid(SysMdi.Controls(i).Name, 4) & ClsDsr.StrZero(SysMdi.Controls(i).Index, 2)
         End If
         Sql = "select IDMODU,DSCMODU from TB_MODULO "
         Sql = Sql & " where IDMODU='" & Cod & "'"
         Sql = Sql & " and   DSCMODU='" & Nome & "'"
         Sql = Sql & " and   CODSYS='" & UCase(mvarCODSYS) & "'"
         'Se o Módulo Existe com o mesmo Código e Descrição
         If mvarDBase.ExisteReg(Sql) Then
            'Existe Item de Menu
            Sql = "update TB_MODULO set DSCMODU='" & Nome & "'"
            Sql = Sql & ", SITMODU='A'"
            Sql = Sql & " where IDMODU='" & Cod & "'"
            Sql = Sql & " and   CODSYS='" & UCase(mvarCODSYS) & "'"
         Else
            '* Não Existe item de Menu ou Foi alterado Índice e/ou Descrição
            Sql = "select IDMODU,DSCMODU from TB_MODULO "
            Sql = Sql & " where DSCMODU='" & Nome & "'"
            Sql = Sql & " and   CODSYS='" & UCase(mvarCODSYS) & "'"

            If mvarDBase.AbreTabela(Sql, RcModu) Then
               '* Se um Item com o mesmo nome realmente existir no Menu,
               '* a posibilidade de ídice alterado é errada e o Item atual
               '* deve ser incluído
               ExisteNoMenuAtual = False
               For j = 0 To SysMdi.Controls.Count - 1
                  '* Verificar Todso os Itens de Menu
                  If TypeName(SysMdi.Controls(j)) = "Menu" And UCase(Mid(SysMdi.Controls(j).Name, 1, 8)) <> "MNUMOUSE" Then
                     RcModu.MoveFirst
                     While Not RcModu.EOF
                        '* Se o controle não for o Item pesquisado e a Descrição
                        '* for a mesma. Existe outro Item com o mesmo nome no menu.
                        If SysMdi.Controls(j).Name & ClsDsr.StrZero(SysMdi.Controls(j).Index, 2) <> "Mnu" & Cod Then
                           If ClsDsr.EliminarString(SysMdi.Controls(j).Caption, "&") = Nome Then
                              ExisteNoMenuAtual = True
                              Exit For
                           End If
                        End If
                        RcModu.MoveNext
                     Wend
                  End If
               Next
               If ExisteNoMenuAtual Then
                  'Não Existe
                  Sql = "insert into TB_MODULO "
                  Sql = Sql & " (CODSYS,IDMODU, DSCMODU,SITMODU,NVMODU)"
                  Sql = Sql & " values "
                  Sql = Sql & " ("
                  Sql = Sql & "'" & (UCase(mvarCODSYS)) & "'"
                  Sql = Sql & "," & SqlStr(Cod)
                  Sql = Sql & "," & SqlStr(Nome)
                  Sql = Sql & ",'A'"
                  Sql = Sql & "," & CStr(Int(Len(Cod) / 2)) + ")"
               Else
                  '* Existe Item de Menu mas foi alterado o índice
                  Sql = "update TB_MODULO"
                  Sql = Sql & " set IDMODU=" & SqlStr(Cod)
                  Sql = Sql & ", SITMODU='A'"
                  Sql = Sql & " where DSCMODU=" & SqlStr(Nome)
                  Sql = Sql & " and   CODSYS=" & SqlStr(UCase(mvarCODSYS))
               End If
            Else
               '* Não Existe ou Existe Item mas foi alterada a Descrição
               Sql = "select IDMODU,DSCMODU from TB_MODULO "
               Sql = Sql & " where IDMODU=" & SqlStr(Cod)
               Sql = Sql & " and   CODSYS=" & SqlStr(UCase(mvarCODSYS))
               If mvarDBase.ExisteReg(Sql) Then
                  'Existe Item de Menu mas foi alterada a Descrição
                  Sql = "update TB_MODULO set DSCMODU=" & SqlStr(Nome)
                  Sql = Sql & ", SITMODU='A'"
                  Sql = Sql & " where IDMODU=" & SqlStr(Cod)
                  Sql = Sql & " and   CODSYS=" & SqlStr(UCase(mvarCODSYS))
               Else
                  'Não Existe
                  Sql = "insert into TB_MODULO "
                  Sql = Sql & " (CODSYS,IDMODU, DSCMODU,SITMODU,NVMODU)"
                  Sql = Sql & " values "
                  Sql = Sql & " ("
                  Sql = Sql & SqlStr(UCase(mvarCODSYS))
                  Sql = Sql & "," & SqlStr(Cod)
                  Sql = Sql & "," & SqlStr(Nome)
                  Sql = Sql & ",'A'"
                  Sql = Sql & "," & CStr(Int(Len(Cod) / 2)) + ")"
               End If
            End If
         End If
         If Nome <> "-" Then
            Call mvarDBase.Executa(Sql)
            '*************
            '*** Atualizar Tabela de Acesso ***
            '*************
            RcAcesso.MoveFirst
            While Not RcAcesso.EOF
               Grp = CStr(RcAcesso("IDGRPUSER"))
               Sql = "select CODSYS from TB_ACCESS "
               Sql = Sql & " where CODSYS=" & SqlStr(UCase(mvarCODSYS))
               Sql = Sql & " and IDGRPUSER =" & SqlStr(Grp$)
               Sql = Sql & " and IDMODU =" & SqlStr(Cod)
               If Not mvarDBase.ExisteReg(Sql) Then
                  Sql = "insert into TB_ACCESS (CODSYS,IDGRPUSER,IDMODU,CODACCESS)"
                  Sql = Sql & " values "
                  Sql = Sql & "(" & SqlStr(UCase(mvarCODSYS))
                  Sql = Sql & "," & SqlStr(Grp$)
                  Sql = Sql & "," & SqlStr(Cod$)
                  Sql = Sql & ", 'LEIA')"
                  Call mvarDBase.Executa(Sql)
               End If
               RcAcesso.MoveNext
            Wend
         End If
      End If
   Next
   RcAcesso.Close
   Sql = "delete from TB_MODULO where SITMODU='I'"
   Call mvarDBase.Executa(Sql)
End Sub
Private Sub mvarSenhaNova_LoadSenhaNova()
   Call F_DefIdiomaSenhaNova
End Sub
Private Sub mvarSenhaNova_TxtSenhaLostFocus()
   'If mvarSenhaNova.ActiveControl Is mvarSenhaNova.cmdOper(1) Then Exit Sub
   'mvarPWDUSU = TxtSenha.Text
   'If Not Valida_Senha Then
   '   mvarSenhaNova.TxtSenha.SetFocus
   'End If
End Sub
Public Function Encrypt(Password As String, Optional Chave As String) As Variant
    Dim v_sqlerrm    As String
    Dim SenhaCript   As String
    Dim var1         As String
    Dim to_text      As String
    Const MIN_ASC = 32
    Const MAX_ASC = 126
    Const NUM_ASC = MAX_ASC - MIN_ASC + 1
    
    'Chave = 2001 ''qualquer nº para montar o algorítimo da criptografia
    Dim offset As Long
    Dim str_len As Integer
    Dim i As Integer
    Dim Ch As Integer
        
    to_text = ""
    offset = NumericPassword(Chave)
    Rnd -1
    Randomize offset
    str_len = Len(Password)
    For i = 1 To str_len
        Ch = Asc(Mid$(Password, i, 1))
        If Ch >= MIN_ASC And Ch <= MAX_ASC Then
            Ch = Ch - MIN_ASC
            offset = Int((NUM_ASC + 1) * Rnd)
            Ch = ((Ch + offset) Mod NUM_ASC)
            Ch = Ch + MIN_ASC
            to_text = to_text & Chr$(Ch)
        End If
    Next i
    
    Encrypt = to_text
End Function

Public Function Decrypt(Password As String, Optional Chave As String) As Variant
   Dim v_sqlerrm  As String
   Dim SenhaCript As String
   Dim var1       As String
   Dim to_text    As String

   Const MIN_ASC = 32  ' Space.
   Const MAX_ASC = 126 ' ~.
   Const NUM_ASC = MAX_ASC - MIN_ASC + 1
   
   'Chave = 2001 ''qualquer nº para montar o algorítimo da criptografia
   Dim offset As Long
   Dim str_len As Integer
   Dim i As Integer
   Dim Ch As Integer
 
   to_text = ""
   offset = NumericPassword(Chave)
   Rnd -1
   Randomize offset
   str_len = Len(Password)
   For i = 1 To str_len
       Ch = Asc(Mid$(Password, i, 1))
       If Ch >= MIN_ASC And Ch <= MAX_ASC Then
           Ch = Ch - MIN_ASC
           offset = Int((NUM_ASC + 1) * Rnd)
           Ch = ((Ch - offset) Mod NUM_ASC)
           If Ch < 0 Then Ch = Ch + NUM_ASC
           Ch = Ch + MIN_ASC
           to_text = to_text & Chr$(Ch)
       End If
   Next i
   
   Decrypt = to_text
End Function
Public Function NumericPassword(ByVal Password As String) As Long
    Dim Value As Long
    Dim Ch As Long
    Dim shift1 As Long
    Dim shift2 As Long
    Dim i As Integer
    Dim str_len As Integer

    str_len = Len(Password)
    For i = 1 To str_len
        ' Adiciona a próxima letra
        Ch = Asc(Mid$(Password, i, 1))
        Value = Value Xor (Ch * 2 ^ shift1)
        Value = Value Xor (Ch * 2 ^ shift2)

        ' Change the shift offsets.
        shift1 = (shift1 + 7) Mod 19
        shift2 = (shift2 + 13) Mod 23
    Next i
    NumericPassword = Value
End Function
Public Function Encrypt2(ByVal Password As String, Optional Key As String) As String
   Dim P    As String
   Dim b    As String
   Dim S    As String
   Dim i    As Integer
   Dim j    As Integer
   Dim A1   As Integer
   Dim A2   As Integer
   Dim A3   As Integer
   
   Dim sAux As String
   
   If Len(Trim(Password)) < 2 Then
      sAux = "PI"
   Else
      sAux = Right$(Password, 1) & Left$(Password, 1)
   End If
   Password = sAux & Password
   If Trim(Key) = "" Then Key = Encrypt2("231072150500", "DIO")
   j = 1
   For i = 1 To Len(Key)
     P = P & Asc(Mid$(Key, i, 1))
   Next
    
   For i = 1 To Len(Password)
     A1 = Asc(Mid$(P, j, 1))
     j = j + 1
     If j > Len(P) Then j = 1
     A2 = Asc(Mid$(Password, i, 1))
     A3 = A1 Xor A2
     b = Hex$(A3)
     If Len(b$) < 2 Then b$ = "0" + b
     S = S + b
   Next
   Encrypt2 = S
End Function
Public Function Decrypt2(ByVal Password As String, Optional Key As String) As String
   Dim P    As String
   Dim b    As String
   Dim S    As String
   Dim i    As Integer
   Dim j    As Integer
   Dim A1   As Integer
   Dim A2   As Integer
   Dim A3   As Integer
   
   If Trim(Key) = "" Then Key = Encrypt2("231072150500", "DIO")
   j = 1
   For i = 1 To Len(Key)
     P = P & Asc(Mid$(Key, i, 1))
   Next
   For i = 1 To Len(Password) Step 2
     A1 = Asc(Mid$(P, j, 1))
     j = j + 1
     If j > Len(P) Then j = 1
     b = Mid$(Password, i, 2)
     A3 = Val("&H" + b)
     A2 = A1 Xor A3
     S = S + Chr$(A2)
   Next
   Decrypt2 = Mid(S, 3)
End Function
