VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_EMPRESA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private WithEvents mvarMe As FrmEmpresa
Attribute mvarMe.VB_VarHelpID = -1
Private mvarsys As Object
Private mvarIDLOJA As Integer
Public Property Set Sys(ByVal vData As Object)
   Set mvarsys = vData
   If VerificaColigada Then
      mvarIDLOJA = Val(mvarsys.Propriedades("IDLOJA"))
   Else
      mvarsys.Propriedades("FCOMANDO") = "End"
   End If
End Property
Public Property Get Sys() As Object
    Set Sys = mvarsys
End Property
Public Property Get Frm() As Object
    Set Frm = mvarMe
End Property
Public Property Get IDLOJA() As Integer
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDLOJA(ByVal vNewValue As Integer)
    mvarIDLOJA = vNewValue
End Property
Public Sub Show()
   If mvarsys.Propriedades("FCOMANDO") <> "End" Then
      mvarMe.Show vbModal
   End If
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmEmpresa
End Sub
Private Sub Class_Terminate()
   On Error Resume Next
   Unload mvarMe
   Set mvarMe = Nothing
End Sub
Private Function VerificaColigada() As Boolean
   Dim sMsg    As String
   Dim sNome   As String
   Dim TbCol   As Object
   Dim nIDCOL  As Integer
   Dim sTag    As String
   Dim bExiste As Boolean
   
   bExiste = False
   nIDCOL = xVal(mvarsys.Propriedades("IDCOLIGADA"))
   If nIDCOL = 0 Then
      If mvarsys.xDb.AbreTabela("Select Min(IDCOLIGADA) [ID] From COLIGADA") Then
         nIDCOL = xVal(mvarsys.xDb.RsAux("ID") & "")
      End If
      Set TbCol = CriarObjeto("BANCO_3R.TB_COLIGADA")
      Set TbCol.xDb = mvarsys.xDb
      bExiste = TbCol.Pesquisar(Ch_IDCOLIGADA:=nIDCOL)
      
      sMsg = "Entre com o nome do responsável pela unidade."
      sNome = ""
      sNome = UCase(InputBox(sMsg, "Responásvel", TbCol.NMCOLIGADA))
      DoEvents
      If sNome = "" Then
         VerificaColigada = False
      Else
         VerificaColigada = True
      End If
   Else
      VerificaColigada = True
   End If
   
   If TbCol Is Nothing Then
      Set TbCol = Nothing
      Set TbCol = CriarObjeto("BANCO_3R.TB_COLIGADA")
      Set TbCol.xDb = mvarsys.xDb
   End If
   If Not TbCol.Pesquisar(Ch_IDCOLIGADA:=nIDCOL) Then
      TbCol.NMCOLIGADA = sNome
      If TbCol.isDirt Then
         If TbCol.Salvar Then
            If mvarsys.xDb.AbreTabela("Select Min(IDCOLIGADA) [ID] From COLIGADA") Then
               nIDCOL = xVal(mvarsys.xDb.RsAux("ID") & "")
            End If
         End If
      End If
   End If
   If TbCol.Pesquisar(Ch_IDCOLIGADA:=nIDCOL) Then
      sTag = Decrypt2(TbCol.Tag)
      If sTag = "" Then
         sTag = SetTag(sTag, "IDUSU", "DIO")
         sTag = SetTag(sTag, "NUMLINC", "1")
         sTag = SetTag(sTag, "DTLINC", Now() + 45)
         TbCol.Tag = Encrypt2(sTag)
      End If
      If IsDate(GetTag(sTag, "DTLINC")) Then
         If DateDiff("d", Now() + 1, CDate(GetTag(sTag, "DTLINC"))) > 45 Then
            If Dir(mvarsys.ExePath & TbCol.NMCOLIGADA & ".lic") = "" Then
               sTag = SetTag(sTag, "DTLINC", "01/01/0000")
               TbCol.Tag = Encrypt2(sTag)
            End If
         End If
      Else
         If Dir(mvarsys.ExePath & TbCol.NMCOLIGADA & ".lic") = "" Then
            sTag = SetTag(sTag, "DTLINC", "01/01/0000")
            TbCol.Tag = Encrypt2(sTag)
         End If
      End If
      If TbCol.isDirt Then
         TbCol.Salvar
      End If
   End If
End Function

Private Sub mvarMe_Activate()
   If GetTag(mvarMe.Tag, "1VEZ", "1") = "1" Then
'      Call MontaTela
'      Call PopulaTela
'      mvarMe_CmdNovoClick
     
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub mvarMe_ChkAtivoEmpClick()
   Dim bOk As Boolean
   bOk = (mvarMe.ChkAtivoEmp.Value = xtpChecked)
   mvarMe.TxtNOME.Enabled = bOk
   mvarMe.TxtCNPJ.Enabled = bOk
   
   mvarMe.TxtCodServFederal.Enabled = bOk
   mvarMe.TxtCodServMunic.Enabled = bOk
   mvarMe.TxtFantasia.Enabled = bOk
   mvarMe.TxtRazao.Enabled = bOk
   mvarMe.TxtRzabrev.Enabled = bOk
   
   mvarMe.TxtInscEst.Enabled = bOk
   mvarMe.TxtInscMunic.Enabled = bOk
   
   
   mvarMe.TxtDTOPERACAO.Enabled = bOk
   mvarMe.TxtEndereco.Enabled = bOk
   mvarMe.TxtBairro.Enabled = bOk
   mvarMe.TxtCidade.Enabled = bOk
   mvarMe.TxtESTADO.Enabled = bOk
   mvarMe.TxtCEP.Enabled = bOk
   mvarMe.TxtTELEFONE1.Enabled = bOk
   mvarMe.TxtTELEFONE2.Enabled = bOk
   
   
   'mvarMe.ChkAtivoEmp.Enabled = bOk
   mvarMe.ChkMatriz.Enabled = bOk
   mvarMe.ChkSimples.Enabled = bOk
   
   'mvarMe.ChkRegEspecialTrib.Enabled = bok
   'mvarMe.ChkTributacao.Enabled = bok
   'mvarMe.ChkIncentivoCult.Enabled = bok
End Sub

Private Sub mvarMe_CmdExcluirClick()
   Dim Node As Object
   Dim sKey As String
   Dim bOk  As Boolean
   
   Set Node = mvarMe.TreeView1.SelectedItem
   sKey = Node.Key
   If GetTag(Node, "TIPO", "") = "LOJA" Then
      'Call ExibirInformacao("No momento você não poderá excluir a estrutura.")
      Dim mvarIDCOLIGADA As Integer
      mvarIDCOLIGADA = xVal(GetTag(mvarMe.TreeView1.SelectedItem, "IDCOLIGADA"))
      bOk = ExcluirLoja(mvarIDCOLIGADA, mvarIDLOJA)
      
   ElseIf GetTag(Node, "TIPO", "") = "SALA" Then
      bOk = ExcluirSala(mvarIDLOJA)
   End If
   If bOk Then
      Call PopulaTela
      On Error Resume Next
      mvarMe.TreeView1.Nodes(sKey).EnsureVisible
      mvarMe.TreeView1.Nodes(sKey).Selected = True
   End If
   
End Sub
Private Sub mvarMe_CmdNovoClick()
   Dim Sql As String
   
   Call LimpaTela
   Set Node = mvarMe.TreeView1.SelectedItem
   If GetTag(Node, "TIPO", "") = "COLIGADA" Then
      If mvarMe.TxtNMCOLIGADA.Visible Then
         mvarMe.TxtNMCOLIGADA.SetFocus
      End If
   
   ElseIf GetTag(Node, "TIPO", "") = "LOJA" Then
      If mvarMe.TxtNOME.Visible Then
         mvarMe.TxtNOME.SetFocus
      End If
   ElseIf GetTag(Node, "TIPO", "") = "SALA" Then
      If mvarMe.TxtDTOPERSALA.Visible Then
         Dim nMax As Integer
         Sql = "Select Max(CODSALA) [COD] "
         Sql = Sql & " From OSALA "
         Sql = Sql & " Where IDLOJA = " & mvarIDLOJA
         If mvarsys.xDb.AbreTabela(Sql) Then
            nMax = xVal(mvarsys.xDb.RsAux("COD") & "")
         End If
         mvarMe.TxtCODSALA.Text = Right("00" & nMax + 1, 2)
         mvarMe.TxtDTOPERSALA.SetFocus
      End If
   End If
End Sub
Private Sub mvarMe_CmdSairClick()
   Unload mvarMe
End Sub
Private Sub mvarMe_CmdSalvarClick()
   Dim sKey As String
   Dim bOk  As Boolean
   Dim Node As Object
   If Not mvarMe.TreeView1.SelectedItem Is Nothing Then
      sKey = mvarMe.TreeView1.SelectedItem.Key
      Set Node = mvarMe.TreeView1.SelectedItem
   End If
   
   If GetTag(Node, "TIPO", "") = "COLIGADA" Then
      bOk = SalvarColigada
   ElseIf GetTag(Node, "TIPO", "") = "LOJA" Then
      bOk = SalvarLoja
   ElseIf GetTag(Node, "TIPO", "") = "SALA" Then
      bOk = SalvarSala(mvarIDLOJA)
   End If
   If bOk Then
      Call PopulaTela
      mvarMe.TreeView1.Nodes(sKey).EnsureVisible
      mvarMe.TreeView1.Nodes(sKey).Selected = True
   End If
   Call ExibirResultado(mvarsys, bOk)
End Sub
Private Function SalvarColigada() As Boolean
   Dim TbCol     As Object
   Dim nCOLIGADA As Integer
      
   If mvarMe.TxtNMCOLIGADA.Text = "" Then
      Call ExibirAviso("Nome Inválido.")
      If mvarMe.TxtNMCOLIGADA.Visible Then
         mvarMe.TxtNMCOLIGADA.SetFocus
      End If
      Exit Function
   End If
   
   Set TbCol = CriarObjeto("BANCO_3R.TB_COLIGADA")
   With TbCol
      .xDb = mvarsys.xDb
      nCOLIGADA = xVal(mvarMe.TxtIDCOLIGADA.Text)
      nCOLIGADA = IIf(nCOLIGADA = 0, 1, nCOLIGADA)
      'Call .Pesquisar(Ch_Where:="NOME='" & mvarMe.TxtNMCOLIGADA.Text & "'")
      If .Pesquisar(Ch_IDCOLIGADA:=nCOLIGADA) Then
         '.IDCOLIGADA = .IDCOLIGADA
         .NMCOLIGADA = mvarMe.TxtNMCOLIGADA.Text
      
'         .CNPJ = mvarMe.TxtCNPJ.Text
'         .DTOPERACAO = mvarMe.TxtDTOPERACAO.Text
'         .ENDERECO = mvarMe.TxtEndereco.Text
'         .BAIRRO = mvarMe.TxtBairro.Text
'         .CIDADE = mvarMe.TxtCidade.Text
'         .ESTADO = mvarMe.TxtESTADO.Text
'         .CEP = mvarMe.TxtCEP.Text
'         .TELEFONE1 = mvarMe.TxtTELEFONE1.Text
'         .TELEFONE2 = mvarMe.TxtTELEFONE2.Text
'         .ATIVO = mvarMe.ChkAtivoEmp.Value
      End If
      If .Salvar Then
         mvarMe.TreeView1.SelectedItem.Text = .NMCOLIGADA
         SalvarColigada = True
      End If
   End With
   Set TbCol = Nothing
End Function
Private Function SalvarLoja() As Boolean
   Dim TbLoja     As Object
   Dim Queries As Collection
   
   Set Queries = New Collection
   
   If mvarMe.TxtNOME.Text = "" Then
      Call ExibirAviso("Nome Inválido.")
      If mvarMe.TxtNOME.Visible Then
         mvarMe.TxtNOME.SetFocus
      End If
      Exit Function
   End If
   If mvarMe.TxtDTOPERACAO.Text = "" Then
      Call ExibirAviso("Data Inválida.")
      If mvarMe.TxtDTOPERACAO.Visible Then
         mvarMe.TxtDTOPERACAO.Text = Format(Now(), "dd/mm/yyyy")
         mvarMe.TxtDTOPERACAO.SetFocus
      End If
      Exit Function
   End If
   
   Set TbLoja = CriarObjeto("BANCO_3R.TB_OLOJA")
   With TbLoja
      .xDb = mvarsys.xDb
      'Call .Pesquisar(Ch_Where:="NOME='" & mvarMe.txtNome.Text & "'")
      Call .Pesquisar(Ch_IDLOJA:=mvarIDLOJA)

      .IDCOLIGADA = xVal(GetTag(mvarMe.TreeView1.SelectedItem, "IDCOLIGADA"))
      .NOME = mvarMe.TxtNOME.Text
      .CNPJ = mvarMe.TxtCNPJ.Text
      .DTOPERACAO = mvarMe.TxtDTOPERACAO.Text
      .ENDERECO = mvarMe.TxtEndereco.Text
      .BAIRRO = mvarMe.TxtBairro.Text
      .CIDADE = mvarMe.TxtCidade.Text
      .ESTADO = mvarMe.TxtESTADO.Text
      .CEP = mvarMe.TxtCEP.Text
      .TELEFONE1 = mvarMe.TxtTELEFONE1.Text
      .TELEFONE2 = mvarMe.TxtTELEFONE2.Text
      .EMAIL = mvarMe.TxtEMAIL.Text
   
      If mvarsys.Propriedades("MailUID") <> mvarMe.TxtEMAIL.Text And mvarMe.TxtEMAIL.Text <> "" Then
         Call QrySaveParam(mvarsys, "MailUID", mvarMe.TxtEMAIL.Text, Queries)
      End If
      If mvarsys.Propriedades("FromDisplayName") <> mvarMe.TxtFromDisplayName.Text And mvarMe.TxtFromDisplayName.Text <> "" Then
         Call QrySaveParam(mvarsys, "FromDisplayName", mvarMe.TxtFromDisplayName.Text, Queries)
      End If
      
                  
      .CODSERVFEDERAL = mvarMe.TxtCodServFederal.Text
      .CODSERVMUNIC = mvarMe.TxtCodServMunic.Text
      .NMFANTASIA = mvarMe.TxtFantasia.Text
      .RZABREV = mvarMe.TxtRzabrev.Text
         
      .INSCEST = mvarMe.TxtInscEst.Text
      .INSCMUNIC = mvarMe.TxtInscMunic.Text
      
      .ATIVO = mvarMe.ChkAtivoEmp.Value
      .FLGFILIAL = IIf(mvarMe.ChkMatriz.Value = 0, 1, 0)
      .SIMPLES = mvarMe.ChkSimples.Value
      
      .REGESPECIALTRIB = 0
      .TRIBUTACAO = 1
      INCENTIVOCULT = 0
      
      If .Salvar Then
         SalvarLoja = True
         mvarIDLOJA = .IDLOJA
         If mvarMe.TxtCODSALA.Text <> "" Then
            SalvarLoja = SalvarSala(.IDLOJA)
         End If
         Call SalvarCCorrente
         Call SalvarContato
         Call SalvarConexao
         
         If mvarsys.xDb.Executa(Queries) Then
            If mvarMe.TxtEMAIL.Text <> "" Then mvarsys.Propriedades("MailUID") = mvarMe.TxtEMAIL.Text
            If mvarMe.TxtFromDisplayName.Text <> "" Then mvarsys.Propriedades("FromDisplayName") = mvarMe.TxtFromDisplayName.Text
         End If
         On Error Resume Next
         mvarsys.xDb.Executa ("Insert USUARIO(IDUSU, IDGRUPO, SENHAUSU, NMUSU, DTCADASTRO, DTEXPIRAUSU, DTEXPIRASENHA, TPUSU, SITUSU, USUCADASTRO, FEXPIRASENHA, MATRUSU, ICORREIO, ECORREIO, ASSINATURA, TimeOut, IDSETOR, IDEMPR, LOCALTRAB, NIVELOS, NMAD, IDPESSOA, IDCOLIGADA, DTVALIDAUSU, TROCARSENHA) VALUES('DPIL','LOJA','','Usuário Loja','2010-03-08', NULL, NULL,'','A', 'DIO', 0, '', '', '', NULL,  0, 0, NULL, NULL, 0, '', 0,   NULL, NULL, 0)")
         
      End If
   End With
   Set TbLoja = Nothing
End Function
Private Function SalvarSala(pIDLOJA As Integer) As Boolean
   Dim TbSala     As Object
   Dim TbAgenda   As Object
   Dim TbMaq      As Object
   Dim TbSalaMaq  As Object
   Dim TbEstoque  As Object
   Dim TbEstLoja  As Object
   Dim TbLoja     As Object
   Dim sCODSALA   As String
   Dim bExiste    As Boolean
   Dim nIDSala    As Integer
   
   If mvarMe.TabEmpresa.Item(2).Visible Then
      If mvarMe.TxtDTOPERSALA.Text = "" Or Not IsDate(mvarMe.TxtDTOPERSALA.Text) Then
         Call ExibirAviso("Data Inválida.")
         mvarMe.TabEmpresa.Item(1).Selected = True
         mvarMe.TxtDTOPERSALA.Text = Format(Now(), "dd/mm/yyyy")
         If mvarMe.TxtDTOPERSALA.Visible And mvarMe.TxtDTOPERSALA.Enabled Then
            mvarMe.TxtDTOPERSALA.SetFocus
         End If
         Exit Function
      End If
      sCODSALA = mvarMe.TxtCODSALA.Text
   Else
      sCODSALA = "01"
   End If
   
   Set TbLoja = CriarObjeto("BANCO_3R.TB_OLOJA", False)
   With TbLoja
      .xDb = mvarsys.xDb
      Call .Pesquisar(Ch_IDLOJA:=pIDLOJA)
   End With
   '*********
   '* Sala
   Set TbSala = CriarObjeto("BANCO_3R.TB_OSALA", False)
'   nIDSala = xVal(mvarMe.TxtIDSALA.Text)
'   nIDSala = IIf(nIDSala = 0, 1, nIDSala)
   With TbSala
      .xDb = mvarsys.xDb
      bExiste = .Pesquisar(Ch_IDLOJA:=pIDLOJA, Ch_IDSALA:=xVal(mvarMe.TxtIDSALA.Text))    ', Ch_Where:=" CODSALA=" & SqlStr(sCODSALA))
      .IDLOJA = pIDLOJA
      .CODSALA = sCODSALA
      '.IDSALA = xVal(mvarMe.TxtIDSALA.Text)
      .ATIVO = mvarMe.ChkAtivo.Value
      If bExiste Then
         If mvarMe.TabEmpresa.Item(2).Visible Then
            .DTOPERACAO = mvarMe.TxtDTOPERSALA.Text
         End If
      Else
         If mvarMe.TabEmpresa.Item(1).Visible Then
            If IsDate(mvarMe.TxtDTOPERSALA.Text) Then
               .DTOPERACAO = mvarMe.TxtDTOPERSALA.Text
            End If
         End If
      
      End If
      If .Salvar Then
         '*********
         '* Agenda
         Set TbAgenda = CriarObjeto("BANCO_3R.TB_OAGENDA", False)
         With TbAgenda
            .xDb = mvarsys.xDb
            If Not .Pesquisar(Ch_IDLOJA:=pIDLOJA, Ch_Where:=" IDSALA=" & TbSala.IDSALA) Then
               .IDLOJA = pIDLOJA
               .IDSALA = TbSala.IDSALA
               .CODAGENDA = TbSala.CODSALA
               
               Call .Salvar
            End If
         End With
         '*********
         '* Máquina
         Set TbSalaMaq = CriarObjeto("BANCO_3R.TB_OSALA_MAQUINA", False)
         With TbSalaMaq
            .xDb = mvarsys.xDb
            If Not .Pesquisar(Ch_IDLOJA:=pIDLOJA, Ch_IDSALA:=TbSala.IDSALA) Then
            
               Set TbMaq = CriarObjeto("BANCO_3R.TB_OMAQUINA")
               With TbMaq
                  .xDb = mvarsys.xDb
                  .IDTPMAQ = 2
                  .CODMAQUINA = "M80_" & TbSala.CODSALA
                  .DTOPERACAO = Format(TbSala.DTOPERACAO, "dd/mm/yyyy")
                  .SITMAQUINA = 1
                  Call .Salvar
               End With
               
               .IDLOJA = pIDLOJA
               .IDSALA = TbSala.IDSALA
               .IDMAQUINA = TbMaq.IDMAQUINA
               .DTINICIO = Format(TbMaq.DTOPERACAO, "dd/mm/yyyy")
               .Salvar
            End If
         End With
         '*********
         '* Estoque
         Set TbEstLoja = CriarObjeto("BANCO_3R.TB_SEST_LOJA", False)
         With TbEstLoja
            .xDb = mvarsys.xDb
            If Not .Pesquisar(Ch_IDLOJA:=pIDLOJA) Then
               Set TbEstoque = CriarObjeto("BANCO_3R.TB_SESTOQUE")
               With TbEstoque
                  .xDb = mvarsys.xDb
                  .DSCEST = TbLoja.NOME
                  .ATIVO = 1
                  Call .Salvar
               End With
               
               .IDLOJA = pIDLOJA
               .IDEST = TbEstoque.IDEST
               .Salvar
            End If
         End With
         SalvarSala = True
      End If
   End With
   
   Set TbSala = Nothing
   Set TbAgenda = Nothing
   Set TbMaq = Nothing
   Set TbSalaMaq = Nothing
   Set TbEstoque = Nothing
   Set TbEstLoja = Nothing
   Set TbLoja = Nothing
End Function
Private Sub SalvarCCorrente()
   Dim pTbConta As Object
   
   Set pTbConta = CriarObjeto("BANCO_3R.TB_FCCORRENTE")
   Set pTbConta.xDb = mvarsys.xDb
   With pTbConta
      Sql = "TPCONTA='D'"
      Sql = Sql + " And  EVENDA=1"
      Sql = Sql + " And  IDLOJA=" & mvarIDLOJA
      If Not .Pesquisar(Ch_Where:=Sql) Then
         .IDLOJA = mvarIDLOJA
         .DSCCONTA = "Caixa"
         .TPCONTA = "D"
         .EVENDA = 1
         .Salvar
      End If
   End With
   
   Set pTbConta = CriarObjeto("BANCO_3R.TB_FCCORRENTE")
   Set pTbConta.xDb = mvarsys.xDb
   With pTbConta
      Sql = "TPCONTA='B'"
      Sql = Sql + " And  EVENDA=1"
      Sql = Sql + " And  IDLOJA=" & mvarIDLOJA
      If Not .Pesquisar(Ch_Where:=Sql) Then
         .IDLOJA = mvarIDLOJA
         .DSCCONTA = "Corrente"
         .TPCONTA = "B"
         .EVENDA = 1
         .Salvar
      End If
   End With
   Set pTbConta = Nothing
'INSERT  INTO FCCORRENTE (IDLOJA, IDCONTA, DSCCONTA, TPCONTA, EVENDA) VALUES
'(1, 1, 'Caixa', 'D',1 );
'INSERT  INTO FCCORRENTE (IDLOJA, IDCONTA, DSCCONTA, TPCONTA, EVENDA) VALUES
'(1, 2, 'Corrente', 'B', 1);
End Sub
Private Sub SalvarContato()
   Dim pTbContato As Object
   
   Set pTbContato = CriarObjeto("BANCO_3R.TB_OCONTATO")
   Set pTbContato.xDb = mvarsys.xDb
   With pTbContato
      Sql = "NOME='Cliente'"
      Sql = Sql + " And  IDLOJA=" & mvarIDLOJA
      If Not .Pesquisar(Ch_Where:=Sql) Then
         .IDLOJA = mvarIDLOJA
         .IDCOLIGADA = mvarsys.Propriedades("IDCOLIGADA")
         .DTCADASTRO = Now()
         .PJ = 0
         .NOME = "Cliente"
         .EMPRESA = ""
         .Salvar
      End If
   End With
   Set pTbContato = Nothing

   Set pTbContato = CriarObjeto("BANCO_3R.TB_OCONTATO")
   Set pTbContato.xDb = mvarsys.xDb
   With pTbContato
      Sql = "NOME='Dpil'"
      Sql = Sql + " And  IDLOJA=" & mvarIDLOJA
      If Not .Pesquisar(Ch_Where:=Sql) Then
         .IDLOJA = mvarIDLOJA
         .IDCOLIGADA = mvarsys.Propriedades("IDCOLIGADA")
         .DTCADASTRO = Now()
         .PJ = 1
         .NOME = "Dpil"
         .EMPRESA = "Dpil Brasil"
         .Salvar
      End If
   End With
   Set pTbContato = Nothing

'INSERT INTO OCONTATO (IDLOJA, IDCONTATO, DTCADASTRO, PJ, NOME, EMPRESA) VALUES (1, 1, GetDate(), 0, 'Cliente', '');
'INSERT INTO OCONTATO (IDLOJA, IDCONTATO, DTCADASTRO, PJ, NOME, EMPRESA) VALUES (1, 2, GetDate(), 1, 'Dpil', 'Dpil Brasil');
End Sub
Private Sub SalvarConexao()
   Dim i          As Integer
   Dim sServer    As String
   Dim sUID       As String
   Dim sPWD       As String
   Dim sAlias     As String
   Dim bExiste    As Boolean
   Dim sUnidade   As String

   i = 0
   If mvarIDLOJA > 1 Then
      sServer = ReadIniFile(mvarsys.LocalReg, "Conection " & i, "SERVER", "")
      sUID = ReadIniFile(mvarsys.LocalReg, "Conection " & i, "UID", "")
      sPWD = ReadIniFile(mvarsys.LocalReg, "Conection " & i, "PWD", "")
      sUnidade = mvarMe.TxtNOME.Text
      sAlias = ReadIniFile(mvarsys.LocalReg, "Conection " & i, "ALIAS", "")
      While sAlias <> ""
         i = i + 1
         sAlias = ReadIniFile(mvarsys.LocalReg, "Conection " & i, "ALIAS", "")
         bExiste = (sAlias = mvarMe.TxtNOME.Text)
         If bExiste Then sAlias = ""
      Wend
      If Not bExiste Then
         Call WriteIniFile(mvarsys.LocalReg, "Conection " & i, "ALIAS", sUnidade)
         Call WriteIniFile(mvarsys.LocalReg, "Conection " & i, "isODBC", "False")
         Call WriteIniFile(mvarsys.LocalReg, "Conection " & i, "DBTIPO", "1")
         Call WriteIniFile(mvarsys.LocalReg, "Conection " & i, "isADO", "True")
         Call WriteIniFile(mvarsys.LocalReg, "Conection " & i, "DBNAME", "G3R_" & sUnidade)
         
         Call WriteIniFile(mvarsys.LocalReg, "Conection " & i, "SERVER", sServer)
         Call WriteIniFile(mvarsys.LocalReg, "Conection " & i, "UID", sUID)
         Call WriteIniFile(mvarsys.LocalReg, "Conection " & i, "PWD", sPWD)
      End If
   End If
End Sub
Private Function ExcluirLoja(pIDCOLIGADA As Integer, pIDLOJA As Integer) As Boolean
   Dim Sql     As String
   Dim TbLoja  As Object
   Dim TbContato As Object
   Dim TbCorrente As Object
   Dim Queries As Collection
   Dim bExiste As Boolean
   
   bExiste = False
   If GetTag(Node, "TIPO", "") = "LOJA" Then
   
      Sql = "Select *"
      Sql = Sql & " From OSALA"
      Sql = Sql & " Where IDLOJA=" & pIDLOJA
      If mvarsys.xDb.ExisteReg(Sql) Then
         bExiste = True
      Else
         Sql = "Select *"
         Sql = Sql & " From OCONTATO"
         Sql = Sql & " Where IDLOJA=" & pIDLOJA
         Sql = Sql & " And IDCONTATO>2"
         If mvarsys.xDb.ExisteReg(Sql) Then
            bExiste = True
         End If
      End If
      If bExiste Then
         Call ExibirAviso("Loja possui itens cadastradas e não pode ser apagada.")
      Else
         Set Queries = New Collection
         Set TbLoja = CriarObjeto("BANCO_3R.TB_OLOJA")
         Set TbContato = CriarObjeto("BANCO_3R.TB_OCONTATO")
         Set TbCorrente = CriarObjeto("BANCO_3R.TB_FCCORRENTE")
         Set TbMeta = CriarObjeto("BANCO_3R.TB_PMETA")
         
         Queries.Add TbContato.QryDelete(Ch_IDLOJA:=pIDLOJA)
         Queries.Add TbCorrente.QryDelete(Ch_IDLOJA:=pIDLOJA)
         Queries.Add TbMeta.QryDelete(Ch_IDLOJA:=pIDLOJA)
         Queries.Add TbLoja.QryDelete(Ch_IDLOJA:=pIDLOJA, Ch_Where:=" IDCOLIGADA = " & pIDCOLIGADA)
         ExcluirLoja = mvarsys.xDb.Executa(Queries)
         With TbLoja
         '   Set .xDb = mvarSys.xDb
         '   If .Pesquisar(Ch_IDLOJA:=pIDLOJA) Then
         '      Queries.Add .QryDelete(Ch_IDLOJA:=.IDLOJA)
         '
         '      ExcluirLoja = mvarSys.xDb.Executa(Queries)
         '   End If
         End With
         Set TbLoja = Nothing
      End If
   End If
End Function
Private Function ExcluirSala(pIDLOJA As Integer) As Boolean
   Dim TbSala     As Object
   Dim TbAgenda   As Object
   Dim TbSalaMaq  As Object
   Dim TbLoja     As Object
   Dim Queries    As Collection
   
   If mvarMe.TabEmpresa.Item(2).Visible Then
      Set TbSala = CriarObjeto("BANCO_3R.TB_OSALA")
      With TbSala
         Set .xDb = mvarsys.xDb
         If .Pesquisar(Ch_IDLOJA:=pIDLOJA, Ch_IDSALA:=xVal(mvarMe.TxtIDSALA.Text)) Then
            Set Queries = New Collection
            Set TbSalaMaq = CriarObjeto("BANCO_3R.TB_OSALA_MAQUINA")
            Set TbAgenda = CriarObjeto("BANCO_3R.TB_OAGENDA")
            'Set TbAgenda.xDb = mvarSys.xDb
            'If TbAgenda.Pesquisar(Ch_IDLOJA:=.IDLOJA, Ch_Where:="IDSALA=" & .IDSALA) Then  '& " And IDAGENDA<>1") Then
            '   While Not TbAgenda.Rs.EOF
            '      Queries.Add TbAgenda.QryDelete(Ch_IDAGENDA:=TbAgenda.Rs("IDAGENDA"), Ch_IDSALA:=.IDSALA)
            '      TbAgenda.Rs.MoveNext
            '   Wend
            'End If
            
            Queries.Add TbAgenda.QryDelete(Ch_IDLOJA:=.IDLOJA, Ch_Where:="IDSALA=" & .IDSALA)
            Queries.Add TbSalaMaq.QryDelete(Ch_IDLOJA:=.IDLOJA, Ch_IDSALA:=.IDSALA)
            Queries.Add TbSala.QryDelete(Ch_IDLOJA:=.IDLOJA, Ch_IDSALA:=.IDSALA)
            
            ExcluirSala = mvarsys.xDb.Executa(Queries)
         End If
      End With
   End If
   
End Function
Private Sub mvarMe_Load()
   Call MontaTela
   Call PopulaTela
End Sub
Private Sub MontaTela()
   With mvarMe
      .PictBottom.BorderStyle = 0
      .GrpBoxBottom.Left = 0
      .GrpBoxBottom.Top = 0
      .GrpBoxBottom.Width = .ScaleWidth
      .GrpBoxBottom.Height = .PictBottom.Height
      
      .TabPgBotton.Left = 0
      .TabPgBotton.Top = 120
      .TabPgBotton.Width = .ScaleWidth
      .TabPgBotton.Height = .PictBottom.Height
   End With
End Sub
Private Sub PopulaTela()
   Dim TbCol As Object
   Dim TbLoj As Object
   Dim TbSal As Object
   Dim Rs    As Object
   Dim xNode1 As TreeViewNode
   Dim xNode2 As TreeViewNode
   Dim xNode3 As TreeViewNode
   Dim i As Integer
   
   Call LimpaTela
   With mvarMe.TreeView1
      On Error Resume Next
      .Nodes.Clear
      .Nodes.Clear
      .FullRowSelect = True
      .LabelEdit = xtpTreeViewLabelManual
      .ShowLines = xtpTreeViewShowLines
      .SingleSel = True
      .ShowPlusMinus = True
      .ShowLines = xtpTreeViewShowLinesAtRoot
      .UseVisualStyle = True
   End With
   
   Set TbCol = CriarObjeto("BANCO_3R.TB_COLIGADA")
   With TbCol
      Set .xDb = mvarsys.xDb
      If .Pesquisar(Ch_Where:="IDCOLIGADA<>0") Then
         While Not .Rs.EOF
            .Popula
            Set xNode1 = mvarMe.TreeView1.Nodes.Add(, , "k" & .IDCOLIGADA, .NMCOLIGADA)
            Call SetTag(xNode1, "TIPO", "COLIGADA")
            Call SetTag(xNode1, "IDCOLIGADA", Mid(xNode1.Key, 2))
            xNode1.Bold = True
            xNode1.Expanded = True
            xNode1.Selected = True
            
            Set TbLoj = CriarObjeto("BANCO_3R.TB_OLOJA")
            With TbLoj
               Set .xDb = mvarsys.xDb
               If .Pesquisar(Ch_Where:="IDCOLIGADA=" & TbCol.IDCOLIGADA) Then
                  While Not .Rs.EOF
                     .Popula
                     Set xNode2 = mvarMe.TreeView1.Nodes.Add(xNode1.Key, xtpTreeViewChild, "l" & .IDLOJA, .NOME)
                     Call SetTag(xNode2, "TIPO", "LOJA")
                     Call SetTag(xNode2, "IDCOLIGADA", .IDCOLIGADA)
                     Call SetTag(xNode2, "IDLOJA", .IDLOJA)
                     Call SetTag(xNode2, "NMLOJA", .NOME)
                     xNode2.Bold = False
                     xNode1.Expanded = True
                     
                     Set TbSal = CriarObjeto("BANCO_3R.TB_OSALA")
                     With TbSal
                        Set .xDb = mvarsys.xDb
                        If .Pesquisar(Ch_IDLOJA:=TbLoj.IDLOJA) Then
                           While Not .Rs.EOF
                              .Popula
                              Set xNode3 = mvarMe.TreeView1.Nodes.Add(xNode2.Key, xtpTreeViewChild, "l" & .IDLOJA & "s" & .IDSALA, "Sala: " & .CODSALA)
                              Call SetTag(xNode3, "TIPO", "SALA")
                              Call SetTag(xNode3, "IDCOLIGADA", TbLoj.IDCOLIGADA)
                              Call SetTag(xNode3, "IDLOJA", .IDLOJA)
                              Call SetTag(xNode3, "IDSALA", .IDSALA)
                              Call SetTag(xNode3, "NMLOJA", TbLoj.NOME)
                              xNode3.Bold = False
                              xNode2.Expanded = True
                              .Rs.MoveNext
                           Wend
                        End If
                     End With
                     .Rs.MoveNext
                  Wend
               End If
            End With
            .Rs.MoveNext
         Wend
      End If
   End With
End Sub

Private Sub mvarMe_TreeView1DblClick()
   Dim sMsg    As String
   Dim sNome   As String
   Dim TbCol   As Object
   Dim nIDCOL  As Integer

   If GetTag(mvarMe.TreeView1.SelectedItem, "TIPO") = "COLIGADA" Then
      nIDCOL = xVal(Mid(mvarMe.TreeView1.SelectedItem.Key, 2))
      
      Set TbCol = CriarObjeto("BANCO_3R.TB_COLIGADA")
      Set TbCol.xDb = mvarsys.xDb
      If TbCol.Pesquisar(Ch_IDCOLIGADA:=nIDCOL) Then
         sMsg = "Entre com o nome do responsável pela unidade."
         sNome = InputBox(sMsg, "Responásvel", TbCol.NMCOLIGADA)
         If sNome <> "" Then
            TbCol.NMCOLIGADA = sNome
            If TbCol.Salvar Then
               mvarMe.TreeView1.SelectedItem.Text = sNome
            End If
         End If
      End If
   End If
End Sub

Private Sub mvarMe_TreeView1NodeClick(ByVal Node As XtremeSuiteControls.TreeViewNode)
   Dim i As Integer
   mvarIDLOJA = xVal(GetTag(Node, "IDLOJA"))
      
   For i = 0 To mvarMe.TabEmpresa.ItemCount - 1
      mvarMe.TabEmpresa.Item(i).Visible = False
   Next
   mvarMe.CmdExcluir.Enabled = True
   mvarMe.CmdNovo.Enabled = True
   Select Case GetTag(Node, "TIPO", "")
      Case "COLIGADA"
         mvarMe.CmdExcluir.Enabled = False
         mvarMe.CmdNovo.Enabled = False
         
         mvarMe.TabEmpresa.Item(0).Visible = True
         mvarMe.TabEmpresa.Item(0).Selected = True
         If mvarMe.TxtNMCOLIGADA.Text <> "" And mvarMe.TxtNMCOLIGADA.Text = Node.Text Then
            Call mvarMe_CmdSalvarClick
            'CALL mvarMe_CmdNovoClick
         End If
         Call PopulaColigada
      Case "LOJA"
         mvarMe.TabEmpresa.Item(1).Visible = True
         mvarMe.TabEmpresa.Item(3).Visible = True
         mvarMe.TabEmpresa.Item(1).Selected = True
         
         If mvarMe.TxtNOME.Text <> "" And mvarMe.TxtNOME.Text = Node.Text Then
            Call mvarMe_CmdSalvarClick
         End If
         Call PopulaLoja
      Case "SALA":
         mvarMe.TabEmpresa.Item(2).Visible = True
         mvarMe.TabEmpresa.Item(2).Selected = True
         If mvarMe.TxtCODSALA.Text <> "" And mvarMe.TxtCODSALA.Text = Node.Text Then
            Call mvarMe_CmdSalvarClick
         End If
         Call PopulaSala
   End Select
'   mvarMe.TabEmpresa.
   DoEvents
   Node.Expanded = True
End Sub
Private Sub PopulaLoja()
   Dim TbLoja     As Object
   Dim nIDLOJA As Integer
   
   nIDLOJA = xVal(GetTag(mvarMe.TreeView1.SelectedItem, "IDLOJA"))
   
   Call LimpaTela
   
   Set TbLoja = CriarObjeto("BANCO_3R.TB_OLOJA")
   With TbLoja
      .xDb = mvarsys.xDb
      If .Pesquisar(Ch_IDLOJA:=nIDLOJA) Then
         '.IDCOLIGADA = xVal(GetTag(mvarMe.TreeView1.SelectedItem, "IDCOLIGADA"))
         mvarIDLOJA = nIDLOJA
         mvarMe.TxtNOME.Text = .NOME
         mvarMe.TxtCNPJ.Text = .CNPJ
         
         mvarMe.TxtCodServFederal.Text = .CODSERVFEDERAL
         mvarMe.TxtCodServMunic.Text = .CODSERVMUNIC
         mvarMe.TxtFantasia.Text = .NMFANTASIA
         mvarMe.TxtRazao.Text = .RAZAO
         mvarMe.TxtRzabrev.Text = .RZABREV
         
         mvarMe.TxtInscEst.Text = .INSCEST
         mvarMe.TxtInscMunic.Text = .INSCMUNIC
         
         mvarMe.TxtDTOPERACAO.Text = .DTOPERACAO
         mvarMe.TxtEndereco.Text = .ENDERECO
         mvarMe.TxtBairro.Text = .BAIRRO
         mvarMe.TxtCidade.Text = .CIDADE
         mvarMe.TxtESTADO.Text = .ESTADO
         mvarMe.TxtCEP.Text = .CEP
         mvarMe.TxtTELEFONE1.Text = .TELEFONE1
         mvarMe.TxtTELEFONE2.Text = .TELEFONE2
         
         If .EMAIL = "" Then
            .EMAIL = mvarsys.Propriedades("MailUID", True)
         End If
         mvarMe.TxtEMAIL.Text = .EMAIL
         mvarMe.TxtFromDisplayName.Text = mvarsys.Propriedades("FromDisplayName", True)
         
         mvarMe.ChkAtivoEmp.Value = .ATIVO
         mvarMe.ChkMatriz.Value = IIf(.FLGFILIAL = 0, 1, 0)
         mvarMe.ChkSimples.Value = .SIMPLES
         
         'mvarMe.ChkRegEspecialTrib.Value = .REGESPECIALTRIB
         'mvarMe.ChkTributacao.Value = .TRIBUTACAO
         'mvarMe.ChkIncentivoCult.Value = .INCENTIVOCULT
         
      End If
   End With
   Set TbLoja = Nothing
End Sub
Private Sub PopulaSala()
   Dim TbSala     As Object
   Dim nIDLOJA As Integer
   Dim nIDSala As Integer
   
   nIDLOJA = xVal(GetTag(mvarMe.TreeView1.SelectedItem, "IDLOJA"))
   nIDSala = xVal(GetTag(mvarMe.TreeView1.SelectedItem, "IDSALA"))
   
   Call LimpaTela
   
   'mvarMe.TabEmpresa.Item(0).Caption = "LOJA: " & GetTag(mvarMe.TreeView1.SelectedItem, "NMLOJA")
   
   Set TbSala = CriarObjeto("BANCO_3R.TB_OSALA")
   With TbSala
      .xDb = mvarsys.xDb
      If .Pesquisar(Ch_IDLOJA:=nIDLOJA, Ch_IDSALA:=nIDSala) Then
         mvarIDLOJA = nIDLOJA
         mvarMe.TxtIDSALA.Text = .IDSALA
         mvarMe.TxtCODSALA.Text = .CODSALA
         mvarMe.TxtDTOPERSALA.Text = .DTOPERACAO
         mvarMe.ChkAtivo.Value = .ATIVO
      End If
   End With
   Set TbLoja = Nothing
End Sub
Private Sub PopulaColigada()
   Dim TbCol     As Object
   Dim nIDCOL As Integer
   Dim nIDSala As Integer
   
   nIDCOL = xVal(GetTag(mvarMe.TreeView1.SelectedItem, "IDCOLIGADA"))
   Call LimpaTela
      
   Set TbCol = CriarObjeto("BANCO_3R.TB_COLIGADA")
   With TbCol
      .xDb = mvarsys.xDb
      If .Pesquisar(Ch_IDCOLIGADA:=nIDCOL) Then
         mvarIDLOJA = nIDCOL
         mvarMe.TxtIDCOLIGADA.Text = .IDCOLIGADA
         mvarMe.TxtNMCOLIGADA.Text = .NMCOLIGADA
      End If
   End With
   Set TbCol = Nothing
End Sub
Private Sub LimpaTela()
   Dim n As Object
   
   'mvarIDLOJA = 0
   On Error Resume Next
   For Each n In mvarMe.Controls
      n.Text = ""
   Next
   mvarMe.TabEmpresa.Item(0).Caption = "COLIGADA"
   mvarMe.TabEmpresa.Item(1).Caption = "LOJA"
   mvarMe.TabEmpresa.Item(2).Caption = "SALA"
End Sub
Private Sub mvarMe_TxtDTOPERACAOLostFocus()
   mvarMe.TxtDTOPERACAO.Text = FormatarData(mvarMe.TxtDTOPERACAO.Text, True)
End Sub
Private Sub mvarMe_TxtDTOPERSALALostFocus()
   mvarMe.TxtDTOPERSALA.Text = FormatarData(mvarMe.TxtDTOPERSALA, True)
End Sub
