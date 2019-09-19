VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_CADSPRODUTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'** Palavras a serem substituidas
'** CADSPRODUTO
'** SPRODUTO
'** IDPROD
'** NMPROD

Option Explicit
Const TbName = "SPRODUTO"
Const IdName = "IDPROD"
Const Campo1 = "NMPROD"
Const PergExc = "Confirma a exclusão do PRODUTO?"
Const Titulo = "PRODUTOs"

Private WithEvents mvarMe  As FrmCADSPRODUTO
Attribute mvarMe.VB_VarHelpID = -1

Private mvarShortBar       As TL_ShortBar
Private mvarRptRec         As ReportRow
Private TbCad              As Object
Private mvarSys            As Object

Private mvarID As Integer
Private mvarCOD As String
Private mvarIDLOJA As Integer
Public Property Set TLCadastro(ByRef vData As Object)
   Dim Sql  As String
   Dim sAux As String
   With vData
      .Titulo = Titulo
      
      Sql = "Select " & IdName & " [Id.]"
      Sql = Sql & ", CODPROD [Código]"
      Sql = Sql & ", " & Campo1 & " [Descrição]"
      Sql = Sql & ", Case ESERVICO When 1 Then 'Serviço' Else 'Material' End  [Tipo]"
      Sql = Sql & ", Case EVENDA When 1 Then 'Venda' Else 'Consumo' End  [Objetivo]"
      Sql = Sql & ", UNIDCONTROLE [Unid.]"
      Sql = Sql & ", 'R$ '+Replace(Cast(VLVENDA as varchar), '.', ',') [Valor]"
      'Sql = Sql & ", VLULTCOMPRA [Valor]"
      'Sql = Sql & ", VLMEDIO [Valor]"
      Sql = Sql & " From " & TbName
      Sql = Sql & " Where IDLOJA= " & mvarIDLOJA
      Sql = Sql & " AND FLGDEL<>'X'"
      Sql = Sql & " Order By EVENDA Desc, ESERVICO Desc, NMPROD, IDPROD"
      
      .Consulta = Sql
      
      sAux = ""
      sAux = sAux & "|CAMPO=" & IdName & "|CHAVE=S|GRUPO=0|VISIBLE=S:"
      .TAGCAMPOS = sAux
   End With
End Property
Public Property Let ID(ByVal vData As String)
   'If mvarID <> vData And Not mvarSys Is Nothing Then
   '   SQL = "Select CODPROD From SPRODUTO Where IDPROD = " & vData
   '   If mvarSys.xDb.Abretabela(SQL) Then
   '      mvarCOD = mvarSys.xDb.RsAux("CODPROD")
   '   End If
   'End If
   mvarID = vData
End Property
Public Property Get ID() As String
   ID = mvarID
End Property
Public Property Let IDLOJA(ByVal vData As Integer)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Integer
   IDLOJA = mvarIDLOJA
End Property
Public Property Set RptRec(ByVal vData As ReportRow)
   Set mvarRptRec = vData
'   mvarIDLOJA = RptRec.Record.Item(0).Value
   If Not RptRec Is Nothing Then
      If Not RptRec.Record Is Nothing Then
         ID = RptRec.Record.Item(0).Value
      End If
   End If
End Property
Public Property Get RptRec() As ReportRow
   Set RptRec = mvarRptRec
End Property
Public Property Set ShortBar(ByVal vData As TL_ShortBar)
   Set mvarShortBar = vData
End Property
Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
   mvarIDLOJA = Val(mvarSys.Propriedades("IDLOJA"))
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Public Sub Show()
   mvarMe.Show vbModal
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmCADSPRODUTO
End Sub
Private Sub mvarMe_Activate()
   If GetTag(mvarMe, "1VEZ", 1) = 1 Then
      Call SetTag(mvarMe, "1VEZ", 0)
      'Call MontaTela
      'Call PopulaTela
      'If mvarID = 0 Then
         mvarMe.TxtNMPROD.SetFocus
      'End If
   End If
End Sub
Private Sub InstanciaClasse()
   Set TbCad = Nothing
   Set TbCad = CriarObjeto("BANCO_3R.TB_" & TbName)
   Set TbCad.xDb = mvarSys.xDb
End Sub
Private Sub PopulaTela(Optional sCOD As String = "")
   Dim bAchou As Boolean
   Dim nQTDMIN As Currency
   Dim Sql As String
   
   If TbCad Is Nothing Then
      Call InstanciaClasse
   End If
   
   With TbCad
      If sCOD <> "" Then
         bAchou = .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="CODPROD=" & sCOD)
         If bAchou Then
            If xVal(mvarMe.TxtID.Text) <> .IDPROD And xVal(mvarMe.TxtID.Text) <> 0 Then
               Call ExibirInformacao("Código pertence a outro produto e não pode ser reutilizado")
               mvarCOD = ""
               bAchou = .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:=IdName & "=" & mvarID)
            End If
         Else
            If xVal(mvarMe.TxtID.Text) <> 0 Then
               If vbNo = ExibirPergunta("Código não existe." & vbNewLine & vbNewLine & "Deseja trocar código do produto?", pDefaultYes:=False) Then
                  mvarCOD = ""
                  bAchou = .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:=IdName & "=" & mvarID)
               End If
            End If
         End If
      Else
         mvarCOD = ""
         bAchou = .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:=IdName & "=" & mvarID)
      End If
      
      
      If bAchou Then
         mvarMe.CmdExcluir.Enabled = True
         mvarCOD = .CODPROD
         
         mvarMe.ChkATIVO.Value = IIf(.SITPROD = "A", xtpChecked, xtpUnchecked)
         mvarMe.TxtID.Text = .IDPROD
         mvarMe.TxtCodProd.Text = .CODPROD
         mvarMe.TxtNMPROD.Text = .NMPROD
         mvarMe.OptESERVICO(.ESERVICO).Value = True
         mvarMe.OptEVENDA(.EVENDA).Value = True
         mvarMe.TxtVLVENDA.Text = ValBr(.VLVENDA)
         mvarMe.TxtVLCOMPRA.Text = ValBr(.VLULTCOMPRA)
         mvarMe.TxtVLMEDIO.Text = ValBr(.VLMEDIO)
         mvarMe.TxtQtdMin.Text = ValBr(.QTDMIN)
         mvarMe.TxtQtdCompra.Text = ValBr(.QTDCOMPRA)
         Call LocalizarCombo(mvarMe.CmbUNIDCONTROLE, .UNIDCONTROLE)
         
         nQTDMIN = 0
         Sql = "Select Sum(QTDITEM) [SALDO]"
         Sql = Sql & " From SMOVEST"
         Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
         Sql = Sql & " And IDPROD=" & mvarID
         If mvarSys.xDb.AbreTabela(Sql) Then
            nQTDMIN = xVal(mvarSys.xDb.RsAux("SALDO") & "")
         End If
         mvarMe.TxtSldDisponivel.Text = ValBr(nQTDMIN)
      Else
         If mvarCOD = "" Then
            Call LimpaTela
         End If
      End If
   End With
End Sub
Public Function Excluir() As Boolean
   Dim Queries As Collection
   
   
   If mvarID <> 0 Then
      If VerificarDependencias() Then
         If ExibirPergunta("Produto: '" & mvarMe.TxtNMPROD & "'" & vbNewLine & vbNewLine & "Confirma Exclusão do Produto?", "Atenção", False) = vbYes Then
            Call TbCad.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:=IdName & "=" & xVal(mvarMe.TxtID))
            
            Set Queries = New Collection
            Call ExcluirDependencias(Queries)
            Queries.Add TbCad.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDPROD:=xVal(mvarMe.TxtID))
            
            If mvarSys.xDb.Executa(Queries) Then
               Call ExibirResultado(mvarSys, True)
               
               'mvarIsDirt = True
               
               Call LimpaTela
               If Not mvarShortBar Is Nothing Then
                  mvarShortBar.TLCadastro.F_Refresh
               End If
               If mvarMe.TxtID.Visible And mvarMe.TxtCodProd.Enabled Then
                  mvarMe.TxtCodProd.SetFocus
               End If
               Excluir = True
               
            Else
               Call ExibirResultado(mvarSys, False)
            End If
            Set Queries = Nothing
         End If
      End If
   End If
End Function
Private Sub ExcluirDependencias(ByRef pQueries As Collection)
   'Dim Tb01 As Object
   
   'Set Tb01 = Nothing
   'Set Tb01 = CriarObjeto("BANCO_3R.TB_OSALA_MAQUINA")
   'Set Tb01.XDb = mvarSys.XDb
            
   'Queries.Add Tb01.QryDelete(Ch_IDMAQUINA:=SqlNum(mvarMe.TxtID))
End Sub
Private Function ValidaCampos() As Boolean
   ' *******************
   Dim bResult As Boolean
   
   bResult = True
   If Trim(mvarMe.TxtCodProd.Text) = "" Then
      'ExibirAviso "Código inválido."
      'bResult = False
      'mvarMe.TxtCODPROD.SetFocus
      'GoTo Fim
   End If
   If Trim(mvarMe.TxtNMPROD.Text) = "" Then
      ExibirAviso "Descrição inválida."
      bResult = False
      mvarMe.TxtNMPROD.SetFocus
      GoTo Fim
   End If
   If Trim(mvarMe.CmbUNIDCONTROLE.Text) = "" Then
      ExibirAviso "Unidade de Medida inválida."
      bResult = False
      mvarMe.CmbUNIDCONTROLE.SetFocus
      GoTo Fim
   End If

Fim:
   ValidaCampos = bResult
End Function
Private Sub LimpaTela()
   Dim n As Object
   On Error Resume Next
   mvarCOD = ""
   With mvarMe
      For Each n In mvarMe.Controls
         n.Text = ""
      Next
   End With
End Sub
Private Sub MontaTela()
   Dim MyTb As Object
   
   With mvarMe.CmbUNIDCONTROLE
      .Clear
      .AddItem "l"
      .AddItem "pct"
      .AddItem "pç"
      .AddItem "un"
   End With
End Sub
Private Sub PopulaClasse(ByRef Queries As Collection)
   Dim oTbSLMaq As Object
   Dim bExiste As Boolean
   
   If TbCad Is Nothing Then Call InstanciaClasse
   
   With TbCad
      bExiste = .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:=IdName & "=" & xVal(mvarMe.TxtID.Text))
      .SITPROD = IIf(mvarMe.ChkATIVO.Value = xtpChecked, "A", "I")
      .IDLOJA = mvarIDLOJA
      .NMPROD = mvarMe.TxtNMPROD.Text
      If Trim(mvarMe.TxtCodProd.Text) = "" Then
         If .IDPROD = 0 Then
            If mvarSys.xDb.AbreTabela("Select Max(IDPROD)+1 [CODPROD] From SPRODUTO") Then
               .CODPROD = mvarSys.xDb.RsAux("CODPROD")
            End If
         Else
            .CODPROD = .IDPROD
         End If
         mvarMe.TxtCodProd.Text = .CODPROD
      End If
      .CODPROD = mvarMe.TxtCodProd.Text
      .UNIDCONTROLE = mvarMe.CmbUNIDCONTROLE.Text
      .UNIDCOMPRA = mvarMe.CmbUNIDCONTROLE.Text
      .UNIDVENDA = mvarMe.CmbUNIDCONTROLE.Text
      .ESERVICO = IIf(mvarMe.OptESERVICO(1).Value, 1, 0)
      .EVENDA = IIf(mvarMe.OptEVENDA(1).Value, 1, 0)
      .VLVENDA = xVal(mvarMe.TxtVLVENDA.Text)
      .VLULTCOMPRA = xVal(mvarMe.TxtVLCOMPRA.Text)
      .VLMEDIO = IIf(.VLMEDIO = 0, xVal(mvarMe.TxtVLVENDA.Text), .VLMEDIO)
      .QTDMIN = xVal(mvarMe.TxtQtdMin.Text)
      .QTDCOMPRA = xVal(mvarMe.TxtQtdCompra.Text)
      .DSCPROD = mvarMe.TxtNMPROD.Text
      
      .EPRODACABADO = 1
      .IDUSU = mvarSys.IDUSU

      If .IDPROD = 0 Then .DTCADASTRO = mvarSys.xDb.SysDate()
      'If .isDirt Or Not bExiste Then   Queries.Add .QrySave
   End With
End Sub
Private Function Salvar() As Boolean
   Dim bOk     As Boolean
   Dim Queries As Collection
   
   If ValidaCampos() Then
      Set Queries = New Collection
      Call PopulaClasse(Queries)
      
      bOk = True
      If TbCad.isDirt Then
         bOk = TbCad.Salvar()
         If bOk Then
            mvarID = TbCad.IDPROD
            mvarMe.TxtID.Text = mvarID
         End If
         If bOk And Queries.Count > 0 Then
            bOk = mvarSys.xDb.Executa(Queries)
         End If
         Call ExibirResultado(mvarSys, bOk, IIf(bOk, 1, 2))
      End If
   End If
   Salvar = bOk
End Function
Private Function VerificarDependencias()
   Dim bResult As Boolean
   Dim sMsg    As String
   Dim Sql     As String
   
   Sql = "Select * " & vbNewLine
   Sql = Sql & " From SMOVEST " & vbNewLine
   Sql = Sql & " Where IDLOJA=" & mvarIDLOJA & vbNewLine
   Sql = Sql & " And IDPROD=" & mvarID & vbNewLine
   If mvarSys.xDb.ExisteReg(Sql) Then
      sMsg = "Este item já possui movimentação de estoque e não pode ser excluído."
      Call ExibirAviso(sMsg)
      GoTo Saida
   End If
   
'   Dim Tb01    As Object
'   Set Tb01 = CriarObjeto("BANCO_3R.TB_OSERVICOEVT")
'   With Tb01
'      Set .Xdb = mvarSys.Xdb
'      If .Pesquisar(Ch_Where:=IdName & "=" & SqlNum(mvarMe.TxtID)) Then
'         sMsg = "Este item já possui compromisso associado e não pode ser excluído."
'         Call ExibirAviso(sMsg)
'         GoTo Saida
'      End If
'   End With
   
   bResult = True
   
   VerificarDependencias = bResult
Saida:
'   Set Tb01 = Nothing
End Function
Private Sub mvarMe_CmdExcluirClick()
   If mvarID <> 0 Then
      If mvarMe.Visible Then
         Call PopulaTela
      End If
      Call Excluir
   End If
End Sub
Private Sub mvarMe_CmdNovoClick()
   If xVal(mvarMe.TxtID) = 0 And Trim(mvarMe.TxtCodProd) = "" Then
      Call LimpaTela
      mvarMe.TxtNMPROD.SetFocus
      'mvarMe.TxtCODPROD.Enabled = True
      'mvarMe.TxtCODPROD.SetFocus
      'mvarMe.TxtID.Enabled = True
      'mvarMe.TxtID.SetFocus
   Else
      If Salvar Then
         Call LimpaTela
         Call InstanciaClasse
      End If
      mvarMe.TxtNMPROD.SetFocus
   End If
End Sub
Private Sub mvarMe_CmdSairClick()
   Unload mvarMe
End Sub
Private Sub mvarMe_CmdSalvarClick()
   If Salvar Then
      If Not mvarShortBar Is Nothing Then
         mvarShortBar.TLCadastro.F_Refresh
      End If
   End If
End Sub
Private Sub mvarMe_Load()
   DoEvents
   Call MontaTela
   Call PopulaTela
   
   mvarMe.TxtCodProd.Enabled = False
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_OptEVENDAClick(Index As Integer)
   mvarMe.TxtVLVENDA.Enabled = mvarMe.OptEVENDA(1).Value
   mvarMe.LblVLVENDA.Enabled = mvarMe.OptEVENDA(1).Value
End Sub
Private Sub mvarMe_TxtCODPRODLostFocus()
   'mvarCOD = xVal(mvarMe.TxtCODPROD.Text)
   'Call PopulaTela
End Sub
Private Sub mvarMe_TxtIDLostFocus()
   If mvarID <> xVal(mvarMe.TxtID.Text) Then
      mvarID = xVal(mvarMe.TxtID.Text)
      Call PopulaTela
   End If
End Sub
Private Sub mvarMe_TxtVLCOMPRAKeyPress(KeyAscii As Integer)
   If Not InArray(KeyAscii, Array(8, 35, 36, 37, 39, 44, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 188)) Then
   'If Not InArray(KeyAscii, Array(8, 35, 36, 44, 46, 37, 39, 110, 188, 96, 97, 98, 99, 100, 101, 102, 103, 104, 105, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57)) Then
      DoEvents
      KeyAscii = 0
   End If
End Sub
Private Sub mvarMe_TxtVLCOMPRALostFocus()
   mvarMe.TxtVLCOMPRA.Text = ValBr(xVal(mvarMe.TxtVLCOMPRA.Text))
End Sub
Private Sub mvarMe_TxtVLVENDAKeyPress(KeyAscii As Integer)
   If Not InArray(KeyAscii, Array(8, 35, 36, 37, 39, 44, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 188)) Then
      DoEvents
      KeyAscii = 0
   End If
End Sub
Private Sub mvarMe_TxtVLVENDALostFocus()
   mvarMe.TxtVLVENDA.Text = ValBr(xVal(mvarMe.TxtVLVENDA.Text))
End Sub
