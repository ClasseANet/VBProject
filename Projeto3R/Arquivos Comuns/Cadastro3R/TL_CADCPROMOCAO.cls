VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_CADCPROMOCAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'** Palavras a serem substituidas
'** CADCPROMOCAO
'** CPROMOCAO
'** IDPROMO
'** DSCPROMO

Option Explicit
Const TbName = "CPROMOCAO"
Const IdName = "IDPROMO"
Const Campo1 = "DSCPROMO"
Const PergExc = "Confirma a exclusão do desconto/promoção?"
Const Titulo = "Descontos / Promoções"

Private WithEvents mvarMe  As FrmCADCPROMOCAO
Attribute mvarMe.VB_VarHelpID = -1

Private mvarShortBar       As TL_ShortBar
Private mvarRptRec         As ReportRow
Private TbCad              As Object
Private mvarSys            As Object

Private mvarID     As Integer
Private mvarIDLOJA As Integer
Public Property Set TLCadastro(ByRef vData As Object)
   Dim SQL  As String
   Dim sAux As String
   With vData
      .Titulo = Titulo
            
      SQL = ""
      SQL = SQL & "Select P.IDPROMO [Id.], P.DSCPROMO [Descrição]"
      SQL = SQL & ", 'R$ '+Cast(Cast(Sum(R.QTDPROD*R.VLUNIT) as Decimal(9,2)) as Varchar) [Total]"
      SQL = SQL & ", 'R$ '+Cast(P.VLDESC as Varchar)+ ' / '+Cast(Cast(P.VLDESC/Sum(R.QTDPROD*R.VLUNIT)*100 as Decimal(9,2)) as Varchar)+'%' [Desconto]"
      SQL = SQL & ", 'R$ '+Cast(P.VALOR as Varchar) [A Pagar]"
      SQL = SQL & ", [Ativo]=Case P.ATIVO when 1 then 'S' else 'N' end"
      SQL = SQL & " From CPROMOCAO P"
      SQL = SQL & " Left Join CPROMO_PROD R On P.IDLOJA=R.IDLOJA And P.IDPROMO=R.IDPROMO"
      SQL = SQL & " Where P.IDLOJA=" & mvarIDLOJA
      SQL = SQL & " Group By P.IDPROMO, P.DSCPROMO, P.VLDESC, P.VALOR, P.ATIVO"
      .Consulta = SQL
      
      sAux = ""
      sAux = sAux & "|CAMPO=" & IdName & "|CHAVE=S|GRUPO=0|VISIBLE=S:"
      .TAGCAMPOS = sAux
   End With
End Property
Public Property Let Id(ByVal vData As String)
   mvarID = vData
End Property
Public Property Get Id() As String
   Id = mvarID
End Property
Public Property Set RptRec(ByVal vData As ReportRow)
   Set mvarRptRec = vData
'   mvarIDLOJA = RptRec.Record.Item(0).Value
   If Not mvarRptRec Is Nothing Then mvarID = mvarRptRec.Record.Item(0).Value
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
   Set mvarMe = New FrmCADCPROMOCAO
End Sub
Private Sub mvarMe_Activate()
   If GetTag(mvarMe, "1VEZ", 1) = 1 Then
      Call SetTag(mvarMe, "1VEZ", 0)
      'Call MontaTela
      'Call PopulaTela
   End If
End Sub
Private Sub InstanciaClasse()
   Set TbCad = Nothing
   Set TbCad = CriarObjeto("BANCO_3R.TB_" & TbName)
   Set TbCad.xDb = mvarSys.xDb
End Sub
Private Sub PopulaTela()
   If TbCad Is Nothing Then
      Call InstanciaClasse
   End If
   
   With TbCad
      If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:=IdName & "=" & mvarID) Then
         mvarMe.CmdExcluir.Enabled = True
         
         mvarMe.TxtID.Text = .IDPROMO
         mvarMe.TxtDSCPRO.Text = .DSCPROMO
         mvarMe.ChkATIVO.Value = .ATIVO
         mvarMe.TxtDTINI.Text = .DTINI
         mvarMe.TxtDTFIM.Text = .DTFIM
         mvarMe.TxtDTINIV.Text = .DTINIV
         mvarMe.TxtDTFIMV.Text = .DTFIMV
         Call LocalizarCombo(mvarMe.CmbNPARCELA, .NPARCELA, True, True)
         mvarMe.ChkFlgServ.Value = .FLGSERV
         mvarMe.ChkFlgTrat.Value = .FLGTRAT
         mvarMe.chkFlgArea.Value = .FLGAREA
               
         Call PopulaLst(.SERVIN, mvarMe.LstTPSERVICO)
         Call PopulaLst(.TRATIN, mvarMe.LstTPTRATAMENTO)
         Call PopulaLst(.AREAIN, mvarMe.LstAREA)
         
         mvarMe.TxtVLTOTAL.Tag = "1"
            mvarMe.TxtVLDESC.Text = ValBr(.VLDESC)
            mvarMe.TxtVALOR.Text = ValBr(.Valor)
            Call PopulaGrdProd
            Call CalculaPromo
         mvarMe.TxtVLTOTAL.Tag = "0"
      Else
         Call LimpaTela
      End If
   End With
End Sub
Private Sub PopulaLst(pIN As String, pList As Object)
   Dim sLst As String
   Dim i As Integer
   
   sLst = pIN
   sLst = "|" & Replace(sLst, ", ", "|") & "|"
   For i = 1 To pList.ListItems.Count
      pList.ListItems(i).Checked = (InStr(sLst, "|" & i & "|") <> 0)
   Next
End Sub
Public Function Excluir() As Boolean
   Dim Queries As Collection
   
   
   If mvarID <> 0 Then
      If VerificarDependencias() Then
         If ExibirPergunta(PergExc, "Atenção", False) = vbYes Then
            Call TbCad.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:=IdName & "=" & xVal(mvarMe.TxtID))
            
            Set Queries = New Collection
            Call ExcluirDependencias(Queries)
            Queries.Add TbCad.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDPROMO:=xVal(mvarMe.TxtID))
            
            If mvarSys.xDb.Executa(Queries) Then
               Call ExibirResultado(mvarSys, True)
               
               'mvarIsDirt = True
               
               Call LimpaTela
               If Not mvarShortBar Is Nothing Then
                  mvarShortBar.TLCadastro.F_Refresh
               End If
               If mvarMe.TxtID.Visible And mvarMe.TxtDSCPRO.Enabled Then
                  mvarMe.TxtDSCPRO.SetFocus
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
   Dim Tb01 As Object
   
   Set Tb01 = Nothing
   Set Tb01 = CriarObjeto("BANCO_3R.TB_CPROMO_PROD")
   Set Tb01.xDb = mvarSys.xDb
            
   pQueries.Add Tb01.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDPROMO:=SqlNum(mvarMe.TxtID))
End Sub
Private Function ValidaCampos() As Boolean
   ' *******************
   Dim bResult As Boolean
   
   bResult = True
   If Trim(mvarMe.TxtDSCPRO.Text) = "" Then
      ExibirAviso "Descrição inválida."
      bResult = False
      mvarMe.TxtDSCPRO.SetFocus
      GoTo Fim
   End If
   If Not IsDate(mvarMe.TxtDTINIV.Text) Then
      ExibirAviso "Data inválida."
      bResult = False
      mvarMe.TxtDTINIV.SetFocus
      GoTo Fim
   End If
   If Not IsDate(mvarMe.TxtDTFIMV.Text) Then
      ExibirAviso "Data inválida."
      bResult = False
      mvarMe.TxtDTFIMV.SetFocus
      GoTo Fim
   End If
   If Not IsDate(mvarMe.TxtDTINI.Text) Then
      ExibirAviso "Data inválida."
      bResult = False
      mvarMe.TxtDTINI.SetFocus
      GoTo Fim
   End If
   If Not IsDate(mvarMe.TxtDTFIM.Text) Then
      ExibirAviso "Data inválida."
      bResult = False
      mvarMe.TxtDTFIM.SetFocus
      GoTo Fim
   End If
   If Trim(mvarMe.TxtVALOR.Text) = "" Then
      ExibirAviso "Valor inválido."
      bResult = False
      mvarMe.TxtVALOR.SetFocus
      GoTo Fim
   End If
Fim:
   ValidaCampos = bResult
End Function
Private Sub LimpaTela()
   Dim n As Object
   Dim i As Integer
   
   On Error Resume Next
   With mvarMe
      For Each n In mvarMe.Controls: n.Text = ""
      Next

      .ChkATIVO.Value = xtpChecked
      .TxtVALOR.Text = "0,00"
      .TxtVLDESC.Text = "0,00"
      .TxtVLTOTAL.Text = "0,00"
      .LblDesc.Caption = "0%"
      
      .CmbNPARCELA.ListIndex = -1
      If .CmbNPARCELA.ListCount > 0 Then .CmbNPARCELA.ListIndex = 0
      .GrdProd.Clear False
      Call IncluiGrdLinhaProd
      For i = 1 To .LstTPSERVICO.ListItems.Count
         .LstTPSERVICO.ListItems.Item(i).Checked = False
      Next
      For i = 1 To .LstTPTRATAMENTO.ListItems.Count
         .LstTPTRATAMENTO.ListItems.Item(i).Checked = False
      Next
      For i = 1 To .LstAREA.ListItems.Count
         .LstAREA.ListItems.Item(i).Checked = False
      Next
      .ChkFlgServ.Value = xtpChecked
      .ChkFlgTrat.Value = xtpChecked
      .chkFlgArea.Value = xtpChecked
   End With
End Sub
Private Sub MontaTela()
   Dim MyTb As Object
   Dim MyLst As XtremeSuiteControls.ListView
   Dim MyItem As XtremeSuiteControls.ListViewItem
   
   mvarMe.TabControl.Item(1).Visible = False
   
   Call MontaGrdProd
   
   Set MyLst = mvarMe.LstTPSERVICO
   Call ConfigLst(MyLst)
   MyLst.ListItems.Clear
   
   Set MyTb = CriarObjeto("BANCO_3R.TB_OTPSERVICO")
   Set MyTb.xDb = mvarSys.xDb
   With MyTb
      If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="ATIVO=1") Then
         If .Rs.recordcount > 6 Then MyLst.ColumnHeaders(1).Width = MyLst.Width - 360
         While Not .Rs.EOF
            .Popula
            Set MyItem = MyLst.ListItems.Add(, .IDTPSERVICO, .DSCSERVICO)
            .Rs.MoveNext
         Wend
      End If
   End With
   
   Set MyLst = mvarMe.LstTPTRATAMENTO
   Call ConfigLst(MyLst)
   MyLst.ListItems.Clear
   
   Set MyTb = CriarObjeto("BANCO_3R.TB_OTPTRATAMENTO")
   Set MyTb.xDb = mvarSys.xDb
   With MyTb
      If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="ATIVO=1") Then
         If .Rs.recordcount > 6 Then MyLst.ColumnHeaders(1).Width = MyLst.Width - 360
         While Not .Rs.EOF
            .Popula
            Set MyItem = MyLst.ListItems.Add(, .IDTPTRATAMENTO, .DSCTRATAMENTO)
            .Rs.MoveNext
         Wend
      End If
   End With
      
   Set MyLst = mvarMe.LstAREA
   Call ConfigLst(MyLst)
   MyLst.ListItems.Clear
   
   Set MyTb = CriarObjeto("BANCO_3R.TB_OAREA")
   Set MyTb.xDb = mvarSys.xDb
   With MyTb
      If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="ATIVO=1") Then
         If .Rs.recordcount > 6 Then MyLst.ColumnHeaders(1).Width = MyLst.Width - 360
         While Not .Rs.EOF
            .Popula
            Set MyItem = MyLst.ListItems.Add(, .IDAREA, .DSCAREA)
            .Rs.MoveNext
         Wend
      End If
   End With
   
   mvarMe.CmbNPARCELA.Clear
   mvarMe.CmbNPARCELA.ListIndex = -1
   Set MyTb = CriarObjeto("BANCO_3R.TB_CCONDPGTO")
   Set MyTb.xDb = mvarSys.xDb
   With MyTb
      If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="ATIVO=1") Then
         While Not .Rs.EOF
            .Popula
            mvarMe.CmbNPARCELA.AddItem .CONDPGTO
            mvarMe.CmbNPARCELA.ItemData(mvarMe.CmbNPARCELA.NewIndex) = .IDCONDPGTO
            .Rs.MoveNext
         Wend
      End If
   End With
   If mvarMe.CmbNPARCELA.ListCount > 0 Then mvarMe.CmbNPARCELA.ListIndex = 0
End Sub
Private Sub ConfigLst(pLst As Object)
   With pLst
      .ListItems.Clear
      .Checkboxes = True
      .AllowColumnReorder = False
      .FullRowSelect = True
      .HideSelection = False
      .MultiSelect = True
      .Sorted = True
      .FlatScrollBar = True
      .View = xtpListViewReport
      .ColumnHeaders.Clear
      .ColumnHeaders.Add 1, "DSC", "Descrição"
      .ColumnHeaders(1).Width = .Width
   End With
End Sub
Private Sub PopulaClasse(ByRef Queries As Collection)
   Dim oTbSLMaq As Object
   Dim bExiste As Boolean
   Dim i As Integer
   Dim sServIn As String
   Dim sTratIn As String
   Dim sAreaIn As String
   
   If TbCad Is Nothing Then Call InstanciaClasse
   
   With TbCad
      bExiste = .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDPROMO:=mvarMe.TxtID.Text)
      .IDLOJA = mvarIDLOJA
      .DSCPROMO = mvarMe.TxtDSCPRO.Text
      .ATIVO = mvarMe.ChkATIVO.Value
      .DTINI = mvarMe.TxtDTINI.Text
      .DTFIM = mvarMe.TxtDTFIM.Text
      .DTINIV = mvarMe.TxtDTINIV.Text
      .DTFIMV = mvarMe.TxtDTFIMV.Text
      .VLDESC = mvarMe.TxtVLDESC.Text
      .Valor = mvarMe.TxtVALOR.Text
      .VLTOTAL = xVal(Replace(mvarMe.TxtVLTOTAL.Text, "R$", ""))
      If mvarMe.CmbNPARCELA.ListIndex >= 0 Then
         .NPARCELA = mvarMe.CmbNPARCELA.ItemData(mvarMe.CmbNPARCELA.ListIndex)
      End If
      .FLGSERV = mvarMe.ChkFlgServ.Value
      .FLGTRAT = mvarMe.ChkFlgTrat.Value
      .FLGAREA = mvarMe.chkFlgArea.Value
      sServIn = ""
      sTratIn = ""
      sAreaIn = ""
      For i = 1 To mvarMe.LstTPSERVICO.ListItems.Count
         If mvarMe.LstTPSERVICO.ListItems(i).Checked Then
            sServIn = sServIn & IIf(sServIn = "", "", ", ") & mvarMe.LstTPSERVICO.ListItems(i).Key
         End If
      Next
      For i = 1 To mvarMe.LstTPTRATAMENTO.ListItems.Count
         If mvarMe.LstTPTRATAMENTO.ListItems(i).Checked Then
            sTratIn = sTratIn & IIf(sTratIn = "", "", ", ") & mvarMe.LstTPTRATAMENTO.ListItems(i).Key
         End If
      Next
      For i = 1 To mvarMe.LstAREA.ListItems.Count
         If mvarMe.LstAREA.ListItems(i).Checked Then
            sAreaIn = sAreaIn & IIf(sAreaIn = "", "", ", ") & mvarMe.LstAREA.ListItems(i).Key
         End If
      Next
      .SERVIN = sServIn
      .TRATIN = sTratIn
      .AREAIN = sAreaIn

      'If .IDPROMO = 0 Then .DTCADASTRO =       .Sys.Xdb.SysDate()
      If .isDirt Or Not bExiste Then Queries.Add .QrySave
   End With
End Sub
Private Sub PopulaClasseProd(ByRef Queries As Collection)
   Dim TbItens As Object
   Dim i As Integer
   '******
   '* Produtos
   If GetTag(mvarMe.GrdProd, "ISDIRT", 0) = 1 Then
      Set TbItens = CriarObjeto("BANCO_3R.TB_CPROMO_PROD")
      Set TbItens.xDb = mvarSys.xDb
      TbItens.IDLOJA = mvarIDLOJA
      TbItens.IDPROMO = mvarID
      
      With mvarMe.GrdProd
         Queries.Add TbItens.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDPROMO:=mvarID)
         For i = 1 To .RowCount - 1
            If xVal(.CellValue(i, "Total")) <> 0 Then
               TbItens.IDLOJA = mvarIDLOJA
               TbItens.IDPROMO = mvarID
               TbItens.IDPROD = Val(.CellValue(i, "IDPROD"))
               TbItens.NMPROD = Mid(Trim(.CellText(i, "IDPROD")), 1, 50)
               TbItens.QTDPROD = xVal(.CellValue(i, "QTDPROD"))
               TbItens.UNIDCONTROLE = .CellValue(i, "UNIDCONTROLE")
               TbItens.VLUNIT = xVal(.CellValue(i, "VLUNIT"))
               If Val(.CellValue(i, "IDPROD")) <> 0 Then
                  Queries.Add TbItens.QrySave
               End If
            End If
         Next
      End With
   End If
End Sub
Private Function Salvar() As Boolean
   Dim bOk     As Boolean
   Dim Queries As Collection
   Dim bSalvou As Boolean
   
   If ValidaCampos() Then
      Set Queries = New Collection
      Call PopulaClasse(Queries)
      
      bOk = True
      If TbCad.isDirt Then
         bOk = TbCad.Salvar()
         If bOk Then
            mvarID = TbCad.IDPROMO
            mvarMe.TxtID.Text = mvarID
         End If
         bSalvou = bOk
      End If
      If bOk Then
         Set Queries = New Collection
         Call PopulaClasseProd(Queries)
         If Queries.Count > 0 Then
            bOk = mvarSys.xDb.Executa(Queries)
            If bOk Then Call SetTag(mvarMe.GrdProd, "ISDIRT", 0)
            bSalvou = bOk
         End If
      End If
      If bSalvou Then Call ExibirResultado(mvarSys, bOk, IIf(bOk, 1, 2))
   End If
   Salvar = bOk
End Function
Private Function VerificarDependencias()
   Dim bResult As Boolean
   Dim sMsg    As String
   Dim Tb01    As Object
   Set Tb01 = CriarObjeto("BANCO_3R.TB_CPACOTE")
   With Tb01
      Set .xDb = mvarSys.xDb
      If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:=IdName & "=" & SqlNum(mvarMe.TxtID)) Then
         sMsg = "Esta promoção possui venda associada e não pode ser excluída."
         Call ExibirAviso(sMsg)
         GoTo Saida
      End If
   End With
   
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
   If xVal(mvarMe.TxtID) = 0 And Trim(mvarMe.TxtDSCPRO) = "" Then
      Call LimpaTela
      mvarMe.TxtDSCPRO.SetFocus
   Else
      If Salvar Then
         Call LimpaTela
         Call InstanciaClasse
         mvarMe.TxtDSCPRO.SetFocus
      Else
         mvarMe.TxtDSCPRO.SetFocus
      End If
   End If
End Sub
Private Sub mvarMe_CmdSairClick()
   Call Salvar
   Unload mvarMe
End Sub
Private Sub mvarMe_CmdSalvarClick()
   If Salvar Then
      Call ExibirResultado(mvarSys, True)
      If Not mvarShortBar Is Nothing Then
         mvarShortBar.TLCadastro.F_Refresh
      End If
   Else
      Call ExibirResultado(mvarSys, False)
   End If
End Sub
Private Sub mvarMe_CmdTodosClick(Index As Integer)
   Dim i As Integer
   Dim MyLst As Object
   
   Select Case Index
      Case 0: Set MyLst = mvarMe.LstTPSERVICO
      Case 1: Set MyLst = mvarMe.LstTPTRATAMENTO
      Case 2: Set MyLst = mvarMe.LstAREA
   End Select
   For i = 1 To MyLst.ListItems.Count
      MyLst.ListItems(i).Checked = True
   Next
End Sub
Private Sub mvarMe_GrdProdAfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
   Dim MyProd  As Object 'TB_SPRODUTO
   
   With mvarMe.GrdProd
      Call .SetCurCell(lRow, lCol + IIf(lCol < .ColCount, 1, -1))
      Call .SetCurCell(lRow, lCol)
   
      If .ColKey(lCol) = "IDPROD" Then
         Set MyProd = CriarObjeto("BANCO_3R.TB_SPRODUTO")
         Set MyProd.xDb = mvarSys.xDb
         If MyProd.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDPROD:=.CellValue(lRow, "IDPROD")) Then
            .CellValue(lRow, "ESERVICO") = MyProd.ESERVICO
            .CellValue(lRow, "UNIDCONTROLE") = MyProd.UNIDCONTROLE
            .CellValue(lRow, "VLUNIT") = ValBr(MyProd.VLVENDA)
            .CellValue(lRow, "Total") = ValBr(xVal(.CellValue(lRow, "QTDPROD")) * xVal(.CellValue(lRow, "VLUNIT")))
            Call CalculaPromo
         End If
         Set MyProd = Nothing
         
      ElseIf .ColKey(lCol) = "QTDPROD" Then
         .CellValue(lRow, "Total") = ValBr(xVal(.CellValue(lRow, "QTDPROD")) * xVal(.CellValue(lRow, "VLUNIT")))
         Call CalculaPromo
      End If
   End With
End Sub
Private Sub mvarMe_GrdProdBeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
   If lConvErr <> 0 Then
      ExibirAviso "Valor inválido."
      eResult = igEditResProceed
   End If
End Sub
Private Sub mvarMe_GrdProdColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
   bDoDefault = False
End Sub
Private Sub mvarMe_GrdProdLostFocus()
   Dim sTag As String
   Dim i    As Integer
   Dim j    As Integer
   
   With mvarMe.GrdProd
      '.Tag = SetTag(.Tag, "ISDIRT", 0)
      If .RowCount <= 1 Then .Tag = SetTag(.Tag, "ISDIRT", 1)
      For i = 1 To .RowCount - 1
         If Not IsMissing(.RowTag(i)) Then sTag = .RowTag(i)
         For j = 1 To .ColCount
            If GetTag(sTag, .ColKey(j), 0) <> .CellValue(i, j) Then .Tag = SetTag(.Tag, "ISDIRT", 1)
            If GetTag(.Tag, "ISDIRT", 0) = 1 Then Exit For
         Next
         If GetTag(.Tag, "ISDIRT", 0) = 1 Then Exit For
      Next
   End With
End Sub

Private Sub mvarMe_GrdProdMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)

   With mvarMe.GrdProd
      If lRow = .RowCount Then
         If .RowCount > .Combos("IDPROD").ListCount Then
            Call ExibirAviso("Não existem mais produtos de venda. (" + CStr(.Combos("IDPROD").ListCount) + ")")
            Exit Sub
         End If
         
         If lRow > 1 Then
            If Val(.CellValue(lRow - 1, "IDPROD")) = 0 Then Exit Sub
            If xVal(.CellValue(lRow - 1, "QTDPROD")) = 0 Then Exit Sub
         End If
         Call IncluiGrdLinhaProd
         
         If Val(.CellValue(lRow, "IDPROD")) = 0 Then
            Call .SetCurCell(lRow, "IDPROD")
         Else
            Call .SetCurCell(lRow, "QTDPROD")
         End If
      End If
   End With
End Sub
Private Sub mvarMe_GrdProdRequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
   With mvarMe.GrdProd
      bCancel = (GetTag(.ColTag(lCol), "EDIT", True) = False)
      If .CellFmtString(lRow, lCol) <> "" Then sText = .CellValue(lRow, lCol)
      eTextEditOpt = GetTag(.ColTag(lCol), "EDITOPT", 0)
   End With
End Sub
Private Sub mvarMe_GrdProdValidate(Cancel As Boolean)
   mvarMe.GrdProd.CommitEdit
End Sub
Private Sub mvarMe_KeyUp(KeyCode As Integer, Shift As Integer)
   Dim lRow As Integer
   If KeyCode = vbKeyDelete Then
      With mvarMe
         If .ActiveControl Is .GrdProd Then
            With .ActiveControl
               lRow = .CurRow
               If lRow <> .RowCount And lRow > 0 Then
                  .RemoveRow (.CurRow)
                  .Tag = SetTag(.Tag, "ISDIRT", 1)
                  .SetFocus
                  If lRow < .RowCount Then
                     .SetCurCell lRow, 1
                  ElseIf lRow > 1 Then
                     .SetCurCell lRow - 1, 1
                  End If
               End If
               For lRow = 1 To .RowCount - 1
                  .CellValue(lRow, 1) = lRow
               Next
            End With
            Call CalculaPromo
         End If
      End With
   End If
End Sub
Private Sub mvarMe_Load()
   DoEvents
   Call MontaTela
   Call PopulaTela
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_TxtIDLostFocus()
   mvarID = xVal(mvarMe.TxtID.Text)
   Call PopulaTela
End Sub
Private Sub mvarMe_TxtVALORChange()
   If mvarMe.TxtVLTOTAL.Tag = "1" Then Exit Sub
   mvarMe.TxtVLTOTAL.Tag = "1"
   mvarMe.TxtVLDESC.Text = ValBr(xVal(mvarMe.TxtVLTOTAL.Text) - xVal(mvarMe.TxtVALOR.Text))
   mvarMe.TxtVLTOTAL.Tag = "0"
End Sub
Private Sub mvarMe_TxtVLDESCChange()
   If mvarMe.TxtVLTOTAL.Tag = "1" Then Exit Sub
   mvarMe.TxtVLTOTAL.Tag = "1"
   mvarMe.TxtVALOR.Text = ValBr(xVal(mvarMe.TxtVLTOTAL.Text) - xVal(mvarMe.TxtVLDESC.Text))
   If xVal(mvarMe.TxtVLTOTAL.Text) > 0 Then
      mvarMe.LblDesc.Caption = ValBr(100 * (xVal(mvarMe.TxtVLDESC.Text) / xVal(mvarMe.TxtVLTOTAL.Text))) & "%"
   Else
      mvarMe.LblDesc.Caption = "0%"
   End If
   mvarMe.TxtVLTOTAL.Tag = "0"
End Sub
Private Sub mvarMe_TxtVLTOTALChange()
   If mvarMe.TxtVLTOTAL.Tag = "1" Then Exit Sub
   mvarMe.TxtVLTOTAL.Tag = "1"
   mvarMe.TxtVALOR.Text = ValBr(xVal(mvarMe.TxtVLTOTAL.Text) - xVal(mvarMe.TxtVLDESC.Text))
   If xVal(mvarMe.TxtVLTOTAL.Text) > 0 Then
      mvarMe.LblDesc.Caption = ValBr(100 * (xVal(mvarMe.TxtVLDESC.Text) / xVal(mvarMe.TxtVLTOTAL.Text))) & "%"
   Else
      mvarMe.LblDesc.Caption = "0%"
   End If
   mvarMe.TxtVLTOTAL.Tag = "0"
End Sub
Private Sub IncluiGrdLinhaProd()
   Dim lRow As Long
   Dim i As Integer
   Dim j As Integer
   Dim k As Integer
   Dim bAchou As Boolean
   Dim nColDefault As Integer
      
   nColDefault = 2
   With mvarMe.GrdProd
      '.Header.Font.Size = 9
      .Header.Font.Name = mvarMe.Font.Name
      .Font.Size = .Header.Font.Size
      .Font.Name = .Header.Font.Name
      
      lRow = .RowCount
      If lRow > 0 Then .RemoveRow lRow
      If lRow > 1 Then .CellForeColor(lRow - 1, "IDPROD") = vbBlack
      
      'If Not (mvarIDATENDIMENTO = 0 And lRow = 0) Then
      If lRow > 0 Then
         .AddRow
         lRow = .RowCount
         .RowHeight(.RowCount) = 20
   
         For i = 1 To .ColCount
            If .CellType(lRow, .ColKey(i)) = igCellCombo Then
               If .Combos(.ColKey(i)).ListCount = 1 Then
                  .CellValue(lRow, .ColKey(i)) = .Combos(.ColKey(i)).ItemValue(0)
               ElseIf .Combos(.ColKey(i)).ListCount > 1 Then
                  If GetTag(.ColTag(.ColKey(i)), "DEFAULT", "") <> "" Then
                     If IsNumeric(GetTag(.ColTag(.ColKey(i)), "DEFAULT", "")) Then
                        .CellValue(lRow, .ColKey(i)) = Val(GetTag(.ColTag(.ColKey(i)), "DEFAULT", "0"))
                     Else
                        .CellValue(lRow, .ColKey(i)) = Val(GetTag(.ColTag(.ColKey(i)), "DEFAULT", ""))
                     End If
                     If lRow > 1 Then
                        For j = 0 To .Combos(.ColKey(i)).ListCount - 1
                           bAchou = False
                           If Not bAchou Then
                              .CellValue(lRow, .ColKey(i)) = .Combos(.ColKey(i)).ItemValue(j)
                           End If
                           For k = 1 To (lRow - 1)
                              bAchou = (.CellValue(lRow, .ColKey(i)) = .CellValue(k, .ColKey(i)))
                              If bAchou Then Exit For
                           Next
                        Next
                     End If
                     
                  End If
               End If
            Else
               If GetTag(.ColTag(.ColKey(i)), "DEFAULT", "") = "IDENTITY" Then
                  .CellValue(lRow, .ColKey(i)) = lRow
               Else
                  .CellValue(lRow, i) = GetTag(.ColTag(.ColKey(i)), "DEFAULT", "")
               End If
            End If

            If GetTag(.ColTag(i), "EDIT", True) = False Then
               .CellForeColor(lRow, i) = vbGrayText
            End If
         Next
         Call mvarMe_GrdProdAfterCommitEdit(lRow, mvarMe.GrdProd.ColIndex("IDPROD"))
         
         If Val(.CellValue(lRow, "IDPROD")) <> 0 And Val(.CellValue(lRow, "QTDPROD")) <> 0 Then
            Call mvarMe_GrdProdAfterCommitEdit(lRow, .ColIndex("IDPROD"))
         End If
      End If
      
      .AddRow
      For i = 1 To .ColCount
         .CellType(.RowCount, i) = igCellText
      Next
      For i = 1 To .ColCount
         If .ColVisible(i) Then
            .CellValue(.RowCount, i) = "Clique para incluir nova linha."
            .CellForeColor(.RowCount, i) = vbGrayText
            .CellTextFlags(.RowCount, i) = igTextNoClip Or igTextCenter
            Exit For
         End If
      Next
      
      
      .RowMode = (.RowCount = 1)
      If lRow <> 0 Then
         .SetCurCell .RowCount, 1
         .SetCurCell lRow, nColDefault
      End If
      If .Visible Then
         .SetFocus
      End If
      .Redraw = True
   End With
End Sub
Private Sub MontaGrdProd()
   Dim TbProd  As Object 'TB_SPRODUTO
   Dim SQL     As String
   Dim sAux    As String
   Dim MyRs    As Object
   Dim i       As Integer
   
   Set TbProd = CriarObjeto("BANCO_3R.TB_SPRODUTO")
   Set TbProd.xDb = mvarSys.xDb
   
   With mvarMe.GrdProd
      .ShowControlsInAllCells = True
      .Header.DragCols = False
      .SilentValidation = True
      '.Header.ImageList = mvarMe.IlsIcons
      '.ImageList = mvarMe.IlsIcons
      With .Combos
         For i = .Count To 1 Step -1
            .Remove i
         Next

         With .Add("IDPROD")
            SQL = ""
            SQL = SQL & "Select Distinct P.IDPROD, P.NMPROD" & vbNewLine
            SQL = SQL & " From SPRODUTO P" & vbNewLine
            SQL = SQL & " Left Join CPROMO_PROD S On S.IDLOJA=P.IDLOJA And S.IDPROD=P.IDPROD And S.IDPROMO=" & mvarID & vbNewLine
            SQL = SQL & " Where P.IDLOJA=" & mvarIDLOJA & vbNewLine
            SQL = SQL & " And P.ESERVICO=1" & vbNewLine
            SQL = SQL & " And (P.SITPROD='A' OR NOT S.IDPROD IS NULL)" & vbNewLine
            SQL = SQL & " Order By P.NMPROD"
            If mvarSys.xDb.AbreTabela(SQL, MyRs) Then
               While Not MyRs.EOF
                  sAux = MyRs("NMPROD")
                  If Len(sAux) < 10 Then
                     sAux = sAux & Space(10 - Len(sAux))
                  End If
                  .AddItem sAux, xVal(MyRs("IDPROD"))
                  MyRs.MoveNext
               Wend
            End If
            Set MyRs = Nothing
            .AutoAdjustWidth
         End With
      End With
      
      .Clear True
      With .AddCol(sKey:="IDITEM", sHeader:="#", lWidth:=20, bVisible:=True)
         .eType = igCellText
         .sCtrlKey = "IDITEM"
      End With
      With .AddCol(sKey:="IDPROD", sHeader:="Produto", lWidth:=185, bVisible:=True)
         .eType = igCellCombo
         .eTypeFlags = igComboBtnFlat
         .sCtrlKey = "IDPROD"
      End With
      With .AddCol(sKey:="ESERVICO", sHeader:="Serviço", lWidth:=65, bVisible:=False)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextRight
         .sCtrlKey = "ESERVICO"
      End With
      
      With .AddCol(sKey:="QTDPROD", sHeader:="Qtd.", lWidth:=40, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextCenter
         .sCtrlKey = "QTDPROD"
      End With
      With .AddCol(sKey:="UNIDCONTROLE", sHeader:="Unid.", lWidth:=40, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextCenter
         .sCtrlKey = "UNIDCONTROLE"
      End With
      With .AddCol(sKey:="VLUNIT", sHeader:="Valor Unit.", lWidth:=50, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextRight
         '.sFmtString = "R$ #,###.00"
         .sCtrlKey = "VLUNIT"
      End With
      With .AddCol(sKey:="Total", sHeader:="Total", lWidth:=60, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextRight
         .sCtrlKey = "Total"
      End With
            
      For i = 1 To .ColCount
         .ColTag(i) = ""
      Next
            
      .ColTag("IDITEM") = SetTag(.ColTag("IDITEM"), "EDIT", False)
      .ColTag("UNIDCONTROLE") = SetTag(.ColTag("UNIDCONTROLE"), "EDIT", False)
      .ColTag("VLUNIT") = SetTag(.ColTag("VLUNIT"), "EDIT", False)
      .ColTag("Total") = SetTag(.ColTag("Total"), "EDIT", False)
            
      .ColTag("QTDPROD") = SetTag(.ColTag("QTDPROD"), "EDITOPT", igTextEditNumberOnly)
      
      .ColTag("IDITEM") = SetTag(.ColTag("IDITEM"), "DEFAULT", "IDENTITY")
      .ColTag("IDPROD") = SetTag(.ColTag("IDPROD"), "DEFAULT", 1)
      .ColTag("UNIDCONTROLE") = SetTag(.ColTag("UNIDCONTROLE"), "DEFAULT", "")
      .ColTag("QTDPROD") = SetTag(.ColTag("QTDPROD"), "DEFAULT", 0)
      .ColTag("VLUNIT") = SetTag(.ColTag("VLUNIT"), "DEFAULT", ValBr(0))
      .ColTag("Total") = SetTag(.ColTag("Total"), "DEFAULT", ValBr(0))
      
      IncluiGrdLinhaProd

   End With
End Sub
Private Sub PopulaGrdProd()
   Dim TbItem As Object 'TB_CITENSVENDA
   Dim sTag As String
   Dim i    As Integer
   Dim SQL  As String

   DoEvents
   Set TbItem = CriarObjeto("BANCO_3R.TB_CPROMO_PROD")
   If TbItem Is Nothing Then Exit Sub
   Set TbItem.xDb = mvarSys.xDb
      
   
   If TbItem.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDPROMO:=mvarID) Then
      With mvarMe.GrdProd
         .Clear
         Call IncluiGrdLinhaProd
         
         While Not TbItem.Rs.EOF
            TbItem.Popula
                  
            Call IncluiGrdLinhaProd
         
            i = .RowCount - 1
            If Not IsMissing(.RowTag(i)) Then sTag = .RowTag(i)
                        
            Call SetTag(sTag, "IDPROD", TbItem.IDPROD)
            Call SetTag(sTag, "QTDPROD", TbItem.QTDPROD)
            Call SetTag(sTag, "UNIDCONTROLE", TbItem.UNIDCONTROLE)
            Call SetTag(sTag, "VLUNIT", ValBr(TbItem.VLUNIT))
            Call SetTag(sTag, "Total", ValBr(TbItem.QTDPROD * TbItem.VLUNIT))
            
            .RowTag(i) = sTag
            
            .CellValue(i, "IDITEM") = i
            .CellValue(i, "IDPROD") = TbItem.IDPROD
            .CellValue(i, "QTDPROD") = TbItem.QTDPROD
            .CellValue(i, "UNIDCONTROLE") = TbItem.UNIDCONTROLE
            .CellValue(i, "VLUNIT") = ValBr(TbItem.VLUNIT)
            .CellValue(i, "Total") = ValBr(TbItem.QTDPROD * TbItem.VLUNIT)

            TbItem.Rs.MoveNext
         Wend
         .RowVisible(1) = True
         '.CellSelected(1, "FOTOTIPO") = True
         .Tag = SetTag(.Tag, "ISDIRT", 0)
      End With
   End If
   Set TbItem = Nothing
End Sub
Private Sub CalculaPromo()
   Dim i As Integer
   Dim nValor As Currency
   
   nValor = 0
   For i = 1 To mvarMe.GrdProd.RowCount
      nValor = nValor + xVal(mvarMe.GrdProd.CellValue(i, "Total"))
   Next
   mvarMe.TxtVLTOTAL.Text = ValBr(nValor)
   If xVal(mvarMe.TxtVLTOTAL.Text) > 0 Then
      mvarMe.LblDesc.Caption = ValBr(100 * (xVal(mvarMe.TxtVLDESC.Text) / xVal(mvarMe.TxtVLTOTAL.Text))) & "%"
   Else
      mvarMe.LblDesc.Caption = "0%"
   End If
End Sub
