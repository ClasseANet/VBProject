VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_VincularVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe    As FrmVincularVenda
Attribute mvarMe.VB_VarHelpID = -1

Dim mvarTbAtend            As Object
Dim CollVendas             As Collection
Dim sQVenda                As String

Private mvarSys            As Object
Private mvarIDLOJA         As Integer
Private mvarIDATENDIMENTO  As Long
Private mvarIDCLIENTE      As Long
Private mvarNOME           As String
Public Property Get NOME() As Long
   NOME = mvarNOME
End Property
Public Property Let IDCLIENTE(ByVal vData As Long)
    mvarIDCLIENTE = vData
    Call DefCliente
End Property
Public Property Get IDCLIENTE() As Long
   IDCLIENTE = mvarIDCLIENTE
End Property
Public Property Let IDATENDIMENTO(ByVal vData As Long)
   If mvarIDATENDIMENTO <> vData Then
      mvarIDATENDIMENTO = vData
      Call DefAtend
      Call DefCliente
   End If
End Property
Public Property Get IDATENDIMENTO() As Long
   IDATENDIMENTO = mvarIDATENDIMENTO
End Property
Public Property Let IDLOJA(ByVal vData As Integer)
    If mvarIDLOJA <> vData Then
      mvarIDLOJA = vData
      Call DefAtend
      Call DefCliente
   End If
End Property
Public Property Get IDLOJA() As Integer
   IDLOJA = mvarIDLOJA
End Property
Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
   If mvarIDLOJA = 0 Then mvarIDLOJA = mvarSys.Propriedades("IDLOJA")
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Private Sub Class_Initialize()
   Set mvarMe = New FrmVincularVenda
End Sub
Private Sub Class_Terminate()
   Set mvarMe = Nothing
End Sub

Private Sub mvarMe_Activate()
   'Call MontaTela
   'Call PopulaTela
   Call VerificaAjuste
End Sub
Private Sub VerificaAjuste()
   Dim Sql     As String
   Dim MyRS    As Object
   Sql = " Select S.*, A.IDCLIENTE" & vbNewLine
   Sql = Sql & " From OSESSAO  S" & vbNewLine
   Sql = Sql & " Join OATENDIMENTO A On S.IDLOJA=A.IDLOJA And S.IDATENDIMENTO=A.IDATENDIMENTO" & vbNewLine
   Sql = Sql & " Where S.IDLOJA = " & mvarIDLOJA & vbNewLine
   Sql = Sql & " And (S.IDVENDA Is Null Or S.IDVENDA=0 Or S.IDITEM Is Null Or S.IDITEM=0)" & vbNewLine
   Sql = Sql & " And S.IDTPSERVICO > 1" & vbNewLine
   Sql = Sql & " And S.IDATENDIMENTO <> " & mvarIDATENDIMENTO & vbNewLine
   Sql = Sql & " And DateDiff(dd, A.DTATEND, GetDate())>=1" & vbNewLine
   Sql = Sql & " Order By S.IDLOJA, S.IDATENDIMENTO, S.IDSESSAO"
   If mvarSys.xDb.AbreTabela(Sql, MyRS) Then
      mvarMe.CmdReconstruir.Enabled = True
      mvarMe.CmdReconstruir.Visible = True
   End If
End Sub
Private Sub mvarMe_ChkClienteClick()
   If mvarMe.ChkCliente.Visible Then
      Call PopulaTela(mvarMe.ChkCliente.Value = xtpChecked)
   End If
End Sub
Private Sub mvarMe_ChkVendasClick()
   If mvarMe.ChkCliente.Visible Then
      Call PopulaTela(mvarMe.ChkCliente.Value = xtpChecked)
   End If
End Sub
Private Sub mvarMe_CmdCancelClick()
   Unload mvarMe
End Sub
Private Sub mvarMe_CmdOkClick()
   If SalvarVinculo Then
      Unload mvarMe
   Else
      Call ExibirStop("Erro ao salvar vinculo.", "Vinculo Venda<->Atendimento")
   End If
End Sub
Private Sub mvarMe_CmdReconstruirClick()
   Call AjustaRelVendas(True, False)
End Sub
Private Sub mvarMe_GrdItensAfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
   Dim sKey As String
   Dim bExiste As Boolean

   With mvarMe.GrdItens
      If lCol < .ColCount Then
         Call .SetCurCell(lRow, lCol + 1)
      Else
         Call .SetCurCell(lRow, lCol - 1)
      End If
      Call .SetCurCell(lRow, lCol)
      
      
      sKey = .CellText(lRow, lCol)
      bExiste = ExisteLstItem(mvarMe.LstVendas, sKey)
      If bExiste Then
         If xVal(mvarMe.LstVendas.ListItems(sKey).ListSubItems(7).Text) < xVal(GetTag(sQVenda, sKey, 0)) Then
            Call SetTag(sQVenda, sKey, xVal(GetTag(sQVenda, sKey, 0)) - 1)
            Call ExibirAviso("Venda com saldo esgotado!")
            .CellValue(lRow, lCol) = ""
            Call .SetCurCell(lRow, lCol)
         End If
     End If
   End With
End Sub

Private Sub mvarMe_GrdItensBeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
   Dim sKey As String
   If lCol = 6 Then
      sKey = mvarMe.GrdItens.CellText(lRow, lCol)
      If sKey <> vNewValue Then
         If Not InArray(vNewValue, Array("", " ", "0", 0, "000000/0")) Then
'            If mvarMe.LstVendas.ListItems(vNewValue).SubItems(7) = 0 Then
'               Call ExibirAviso("Venda com saldo esgotado!")
'               eResult = igEditResCancel
'            End If
            Call SetTag(sQVenda, vNewValue, 1 + xVal(GetTag(sQVenda, CStr(vNewValue), 0)))
'            mvarMe.LstVendas.ListItems(vNewValue).SubItems(7) = mvarMe.LstVendas.ListItems(vNewValue).SubItems(7) - 1
         End If
         If Not InArray(sKey, Array("", " ", "0", 0, "000000/0")) Then
            Call SetTag(sQVenda, sKey, xVal(GetTag(sQVenda, sKey, 0)) - 1)
'            mvarMe.LstVendas.ListItems(sKey).SubItems(7) = mvarMe.LstVendas.ListItems(sKey).SubItems(7) + 1
         End If
         
         
      End If
   Else
      eResult = igEditResCancel
   End If
End Sub

Private Sub mvarMe_KeyUp(KeyCode As Integer, Shift As Integer)
   If mvarMe.ActiveControl Is mvarMe.GrdItens Then
      If KeyCode = vbKeyDelete Then
         mvarMe.GrdItens.CellValue(mvarMe.GrdItens.CurRow, "VENDA") = 0
      End If
   End If
End Sub

Private Sub mvarMe_LblVendasDblClick()
   Call AjustaRelVendas
End Sub

Private Sub mvarMe_Load()
   Call MontaTela
   Call PopulaTela
End Sub
Private Sub MontaTela()
   Dim i As Integer
   With mvarMe
      .ChkCliente.Caption = "Vendas do Cliente"
      .ChkVendas.Caption = "Serviços do Atendimento Nº" & StrZero(mvarIDATENDIMENTO, 6)
      .ChkCliente.Value = xtpChecked
      .ChkVendas.Value = xtpChecked
      With .LstVendas
         .LabelEdit = xtpListViewLabelManual
         .AllowColumnReorder = False
         .Checkboxes = True
         .GridLines = False
         .View = xtpListViewReport
         .FullRowSelect = True
         .Sorted = True
         
         .ColumnHeaders.Clear
         .ColumnHeaders.Add 1, "IDVENDA", "Nº Venda", 1000, xtpListViewColumnLeft
         .ColumnHeaders.Add 2, "IDITEM", "#", 400, xtpListViewColumnLeft
         .ColumnHeaders.Add 3, "DTVENDA", "Data", 1100, xtpListViewColumnCenter
         .ColumnHeaders.Add 4, "NOME", "Cliente", 2300, xtpListViewColumnLeft
         .ColumnHeaders.Add 5, "VALOR", "Valor", 1000, xtpListViewColumnRight
         .ColumnHeaders.Add 6, "IDPROD", "IDPROD", 0, xtpListViewColumnLeft
         .ColumnHeaders.Add 7, "PRODUTO", "Produto", 2300, xtpListViewColumnLeft
         .ColumnHeaders.Add 8, "SALDO", "Saldo", 700, xtpListViewColumnRight
         
         .ColumnHeaders(5).Width = 0
      End With
      With .GrdItens
         .Clear True
         With .Combos
            If .Count > 0 Then
               For i = .Count To 1 Step -1
                 .Remove i
               Next
            End If
            With .Add("VENDA")
               .AddItem " ", 0
               '.AddItem TbArea.IDAREA, xval(TbArea.IDAREA)
               '.AutoAdjustWidth
               .Width = 80
            End With
         End With
            
         
         With .AddCol(sKey:="IDATENDIMENTO", sHeader:="Nº Atendimento", lWidth:=90, bVisible:=True)
            .eType = igCellText
            .eTextFlags = igTextCenter
            .sCtrlKey = "IDATENDIMENTO"
         End With
         With .AddCol(sKey:="IDSESSAO", sHeader:="Item", lWidth:=40, bVisible:=True)
            .eType = igCellText
            .eTextFlags = igTextCenter
            .sCtrlKey = "IDSESSAO"
         End With
         With .AddCol(sKey:="DSCSERVICO", sHeader:="Serviço", lWidth:=120, bVisible:=True)
            .eType = igCellText
            .eTextFlags = igTextLeft
            .sCtrlKey = "DSCSERVICO"
         End With
         With .AddCol(sKey:="DSCTRATAMENTO", sHeader:="Tratamento", lWidth:=120, bVisible:=True)
            .eType = igCellText
            .eTextFlags = igTextLeft
            .sCtrlKey = "DSCTRATAMENTO"
         End With
         With .AddCol(sKey:="DSCAREA", sHeader:="Área", lWidth:=80, bVisible:=True)
            .eType = igCellText
            .eTextFlags = igTextLeft
            .sCtrlKey = "DSCAREA"
         End With
         With .AddCol(sKey:="VENDA", sHeader:="Venda", lWidth:=100, bVisible:=True)
            .eType = igCellCombo
            .eTextFlags = igTextLeft
            .sCtrlKey = "VENDA"
         End With
        
        With .AddCol(sKey:="IDTPSERVICO", sHeader:="IDTPSERVICO", lWidth:=60, bVisible:=False)
            .eType = igCellText
            .eTextFlags = igTextLeft
            .sCtrlKey = "IDTPSERVICO"
         End With
        With .AddCol(sKey:="IDTPTRATAMENTO", sHeader:="IDTPTRATAMENTO", lWidth:=60, bVisible:=False)
            .eType = igCellText
            .eTextFlags = igTextLeft
            .sCtrlKey = "IDTPTRATAMENTO"
         End With
        With .AddCol(sKey:="IDAREA", sHeader:="IDAREA", lWidth:=60, bVisible:=False)
            .eType = igCellText
            .eTextFlags = igTextLeft
            .sCtrlKey = "IDAREA"
         End With
        With .AddCol(sKey:="IDPROD", sHeader:="IDPROD", lWidth:=60, bVisible:=False)
            .eType = igCellText
            .eTextFlags = igTextLeft
            .sCtrlKey = "IDPROD"
         End With
        With .AddCol(sKey:="IDVENDA", sHeader:="IDVENDA", lWidth:=60, bVisible:=False)
            .eType = igCellText
            .eTextFlags = igTextLeft
            .sCtrlKey = "IDVENDA"
         End With
        With .AddCol(sKey:="IDITEM", sHeader:="IDITEM", lWidth:=60, bVisible:=False)
            .eType = igCellText
            .eTextFlags = igTextLeft
            .sCtrlKey = "IDITEM"
         End With
      End With
   End With
End Sub
Private Sub PopulaTela(Optional bCliente = True)
   Dim Sql     As String
   Dim Sql2    As String
   Dim MyRS    As Object
   Dim MyRsV    As Object
   Dim oItem   As ListViewItem
   Dim n       As Variant
   Dim sKey    As String
   Dim sText   As String
   Dim i       As Integer
   Dim bLimpa  As Boolean
   
   mvarMe.MousePointer = vbHourglass
   
   mvarMe.LstVendas.ListItems.Clear
'   Sql = GetSqlSaldoVenda0(bCliente)
'   Sql = GetSqlSaldoVenda1(bCliente)
   Sql = GetSqlSaldoVenda(bCliente)
   If mvarSys.xDb.AbreTabela(Sql, MyRS) Then
      While Not MyRS.EOF
         sKey = MyRS("IDVENDA") & "/" & MyRS("IDITEM")
         
         With mvarMe.LstVendas.ListItems
            If ExisteLstItem(mvarMe.LstVendas, sKey) Then
               Set oItem = .Add(, , MyRS("IDVENDA"))
            Else
               Set oItem = .Add(, sKey, MyRS("IDVENDA"))
            End If
            oItem.ListSubItems.Add 1, "IDITEM", MyRS("IDITEM")
            oItem.ListSubItems.Add 2, "DTVENDA", Format(MyRS("DTVENDA"), "dd/mm/yy")
            oItem.ListSubItems.Add 3, "NOME", MyRS("NOME")
            oItem.ListSubItems.Add 4, "VALOR", "R$ " & MyRS("VLVENDA")
            oItem.ListSubItems.Add 5, "IDPROD", MyRS("IDPROD")
            oItem.ListSubItems.Add 6, "PRODUTO", MyRS("PRODUTO")
            oItem.ListSubItems.Add 7, "SALDO", MyRS("SALDO")
            oItem.ListSubItems(6).Bold = True
            oItem.ListSubItems(7).Bold = True
         End With
         MyRS.MoveNext
      Wend
   End If
   mvarMe.LstVendas.ColumnHeaders(6).Width = 0

   If mvarIDATENDIMENTO <> 0 Then
      Sql = ""
      Sql = Sql & "Select Distinct Right('000000'+Rtrim(Cast(I.IDATENDIMENTO As Char)),6) [IDATENDIMENTO]" & vbNewLine
      Sql = Sql & ", I.IDSESSAO, S.DSCSERVICO, T.DSCTRATAMENTO, isNull(A.DSCAREA, '--') [DSCAREA]" & vbNewLine
      Sql = Sql & ", Right('000000'+Rtrim(Cast(I.IDVENDA As Char)),6)+ '/' + RTrim(Cast(I.IDITEM As Char)) [VENDA]" & vbNewLine
      Sql = Sql & ", I.IDTPSERVICO, I.IDTPTRATAMENTO, I.IDAREA"
      Sql = Sql & ", RTrim(cast(P.IDPROD As char)) + '|' [IDPROD]"
      Sql = Sql & ", isNull(I.IDVENDA,0) [IDVENDA], isNull(I.IDITEM,0) [IDITEM]"
      Sql = Sql & " From OSESSAO I " & vbNewLine
      Sql = Sql & " Join OTPSERVICO S     On I.IDLOJA=S.IDLOJA And I.IDTPSERVICO=S.IDTPSERVICO" & vbNewLine
      Sql = Sql & " Join OTPTRATAMENTO T  On I.IDLOJA=T.IDLOJA And I.IDTPTRATAMENTO=T.IDTPTRATAMENTO" & vbNewLine
      Sql = Sql & " Left Join OTRATAMENTO_PROD TP On TP.IDLOJA=I.IDLOJA And TP.IDTPSERVICO=I.IDTPSERVICO And (TP.IDTPTRATAMENTO=I.IDTPTRATAMENTO or TP.IDTPTRATAMENTO=0) And (TP.IDAREA=I.IDAREA or TP.IDAREA=0)" & vbNewLine
      Sql = Sql & " Left Join SPRODUTO P On TP.IDLOJA=P.IDLOJA And TP.IDPROD=P.IDPROD AND P.VLVENDA>0" & vbNewLine
      Sql = Sql & " Left Join OAREA A          On I.IDLOJA=A.IDLOJA And I.IDAREA=A.IDAREA" & vbNewLine
      Sql = Sql & " Where I.IDLOJA = " & mvarIDLOJA & vbNewLine
      Sql = Sql & " And I.IDATENDIMENTO= " & mvarIDATENDIMENTO
      If mvarSys.xDb.AbreTabela(Sql, MyRS) Then
         Set CollVendas = New Collection
         MyRS.MoveFirst
         sText = ""
         bLimpa = False
         While Not MyRS.EOF
            sKey = MyRS("VENDA") & ""
            bLimpa = bLimpa Or (MyRS("IDVENDA") <> 0 And MyRS("IDITEM") = 0) Or (MyRS("IDVENDA") = 0 And MyRS("IDITEM") <> 0)
            '* Contagem de Vendas já associadas
            If Not InArray(sKey, Array("", " ", "0", 0, "000000/0")) Then
               If MyRS("IDSESSAO") <> sText Then
                  If Not ExisteItem(CollVendas, sKey) Then
                     If (MyRS("IDVENDA") <> 0 And MyRS("IDITEM") <> 0) Then
                        CollVendas.Add sKey, sKey
                        
                        With mvarMe.LstVendas.ListItems
                           If Not ExisteLstItem(mvarMe.LstVendas, sKey) Then
                              Sql = GetSqlSaldoVenda(bCliente, pIDVENDA:=MyRS("IDVENDA"))
                              If mvarSys.xDb.AbreTabela(Sql, MyRsV) Then
                                 Set oItem = .Add(, sKey, MyRsV("IDVENDA"))
                              
                                 oItem.ListSubItems.Add 1, "IDITEM", MyRsV("IDITEM")
                                 oItem.ListSubItems.Add 2, "DTVENDA", Format(MyRsV("DTVENDA"), "dd/mm/yy")
                                 oItem.ListSubItems.Add 3, "NOME", MyRsV("NOME")
                                 oItem.ListSubItems.Add 4, "VALOR", "R$ " & MyRsV("VLVENDA")
                                 oItem.ListSubItems.Add 5, "IDPROD", MyRsV("IDPROD")
                                 oItem.ListSubItems.Add 6, "PRODUTO", MyRsV("PRODUTO")
                                 oItem.ListSubItems.Add 7, "SALDO", MyRsV("SALDO")
                                 oItem.ListSubItems(6).Bold = True
                                 oItem.ListSubItems(7).Bold = True
                                 oItem.Checked = True
                              End If
                           End If
                        End With
                        
                     End If
                  End If
                  'Call SetTag(sQVenda, sKey, 1 + xVal(GetTag(sQVenda, sKey, 0)))
               End If
            End If
            sText = MyRS("IDSESSAO")
            MyRS.MoveNext
         Wend
         If bLimpa Then
            Sql2 = "Update OSESSAO "
            Sql2 = Sql2 & " Set IDVENDA=0"
            Sql2 = Sql2 & " , IDITEM=0"
            Sql2 = Sql2 & " , AlterStamp=1"
            Sql2 = Sql2 & " , TimeStamp=GetDate()"
            Sql2 = Sql2 & " Where IDLOJA=" & mvarIDLOJA
            Sql2 = Sql2 & " And IDATENDIMENTO=" & mvarIDATENDIMENTO
            Sql2 = Sql2 & " And (isNull(IDVENDA)=0 Or isNull(IDITEM,0)=0)"
            Call mvarSys.xDb.Executa(Sql2)
            Call mvarSys.xDb.AbreTabela(Sql, MyRS)
         End If
         MyRS.MoveFirst
         With mvarMe.GrdItens
            '* Popular Vendas associadas ao combo do grid
            .Combos("VENDA").Clear
            .Combos("VENDA").AddItem " ", 0
            For Each n In CollVendas
               .Combos("VENDA").AddItem n, n
            Next
            
            '* Remover itens repetidos e montar lista de posíveis produtos
            .Redraw = False
            .Clear
            .FillFromRS MyRS
            For i = .RowCount To 2 Step -1
               If .CellText(i, "IDSESSAO") = .CellText(i - 1, "IDSESSAO") Then
                  .CellValue(i - 1, "IDPROD") = .CellText(i - 1, "IDPROD") & .CellText(i, "IDPROD")
                  .RemoveRow i
               End If
            Next
            .Redraw = True
         End With
      End If
   End If
   
   mvarMe.MousePointer = vbDefault
End Sub
Private Function SalvarVinculo() As Boolean
   Dim i       As Integer
   Dim Sql     As String
   Dim sAux    As String
   Dim cQuery  As Collection
   Dim nVenda  As Double
   Dim nItem   As Integer
   Dim bOk     As Boolean
   Dim cVendas As Collection
   
   Set cQuery = New Collection
   Sql = "Delete "
   Sql = Sql & " From OATENDIMENTO_VENDA "
   Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
   Sql = Sql & " And IDATENDIMENTO=" & mvarIDATENDIMENTO
   cQuery.Add Sql
   
   Set cVendas = New Collection
   For i = 1 To mvarMe.GrdItens.RowCount
      nVenda = 0
      nItem = 0
      sAux = mvarMe.GrdItens.CellValue(i, "VENDA")
      If Not InArray(sAux, Array("", " ", "0", 0, "000000/0")) Then
         nVenda = xVal(Mid(sAux, 1, InStr(sAux, "/") - 1))
         nItem = xVal(Mid(sAux, InStr(sAux, "/") + 1))
      End If
      If Not InArray(sAux, Array("", " ", "0", 0, "000000/0")) And nVenda <> 0 And nItem <> 0 Then
         If Not ExisteItem(cVendas, CStr(nVenda)) Then
            cVendas.Add nVenda, CStr(nVenda)
            
            Sql = "Insert Into OATENDIMENTO_VENDA (IDLOJA, IDVENDA, IDATENDIMENTO, ALTERSTAMP" & vbNewLine
            Sql = Sql & ", TIMESTAMP) " & vbNewLine
            Sql = Sql & " Values " & vbNewLine
            Sql = Sql & "(" & vbNewLine
            Sql = Sql & mvarIDLOJA & vbNewLine
            Sql = Sql & ", " & nVenda & vbNewLine
            Sql = Sql & ", " & mvarIDATENDIMENTO & vbNewLine
            Sql = Sql & ", " & "1" & vbNewLine
            Sql = Sql & ", " & "GetDate()" & vbNewLine
            Sql = Sql & ")" & vbNewLine
            cQuery.Add Sql
         End If
      End If
      
      Sql = "Update OSESSAO "
      Sql = Sql & " Set IDITEM = " & nItem
      Sql = Sql & " , IDVENDA= " & IIf(nItem = 0, 0, nVenda)
      Sql = Sql & " , ALTERSTAMP=1 "
      Sql = Sql & " , TIMESTAMP=GetDate() "
      Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
      Sql = Sql & " AND IDATENDIMENTO=" & mvarIDATENDIMENTO
      Sql = Sql & " AND IDSESSAO=" & mvarMe.GrdItens.CellValue(i, "IDSESSAO")
      cQuery.Add Sql
   Next
   bOk = mvarSys.xDb.Executa(cQuery)
   SalvarVinculo = bOk
End Function
Public Sub Show(Optional pMode As FormShowConstants = vbModal)
   DoEvents
   If mvarIDLOJA = 0 Then
      Call ExibirAviso("Código da Loja Inválido! (""" & mvarIDLOJA & """)")
      Exit Sub
   End If
   mvarMe.Show pMode
End Sub
Private Function GetSqlSaldoVenda0(Optional bCliente, Optional bAteDTVENDA As Boolean) As String
   Dim Sql        As String
   Dim sMsg       As String
   Dim sIDVENDA   As String
   Dim i          As Integer
   Dim bPorItem   As Boolean
   
   bPorItem = False
   sIDVENDA = "0"
   
'   For i = 1 To mvarMe.GrdVenda.RowCount
'      If xVal(mvarMe.GrdVenda.CellValue(i, "VENDA")) <> 0 Then
'         sIDVENDA = sIDVENDA & ", " & xVal(mvarMe.GrdVenda.CellValue(i, "VENDA"))
'      End If
'   Next
   
   sMsg = "Apenas o cliente corrente? " & vbNewLine & vbNewLine '& "''" & mvarMe.TxtNOME & "''"
   If IsMissing(bCliente) Then bCliente = (vbYes = ExibirPergunta(sMsg, "Saldo de Vendas"))
      
   Sql = "Select Distinct Right('000000'+Rtrim(Cast(V.IDVENDA As Char)),6) [IDVENDA]" & vbNewLine
   Sql = Sql & ", Cast(Convert(Char(10),V.DTVENDA, 103) As SmallDatetime) [DTVENDA]" & vbNewLine
   Sql = Sql & ", C.TEL1, C.NOME" & vbNewLine
   Sql = Sql & ", 'R$ ' + Cast(V.VLVENDA as Varchar) As [VENDA], 'R$ '+Cast(V.VLPGTO as Varchar) [PGTO]" & vbNewLine
   Sql = Sql & ", V.VLVENDA, V.VLPGTO, V.IDMOEDA" & vbNewLine
   Sql = Sql & ", Cast(( IsNull((SELECT SUM(I.QTDVENDA)" & vbNewLine
   Sql = Sql & "     FROM CITENSVENDA I" & vbNewLine
   Sql = Sql & "     JOIN SPRODUTO P ON I.IDLOJA=P.IDLOJA And I.IDPROD=P.IDPROD And P.ESERVICO=1 And P.EVENDA=1" & vbNewLine
   Sql = Sql & "     Where I.IDLOJA = V.IDLOJA" & vbNewLine
   If bPorItem Then
      Sql = Sql & "     AND I.IDVENDA=V.IDVENDA" & vbNewLine
      Sql = Sql & "     AND I.IDLOJA=IV.IDLOJA" & vbNewLine
      Sql = Sql & "     AND I.IDPROD=IV.IDPROD),0)" & vbNewLine
   Else
      Sql = Sql & "     AND I.IDVENDA=V.IDVENDA),0)" & vbNewLine
   End If
   Sql = Sql & "   -" & vbNewLine
   Sql = Sql & "     IsNull((SELECT COUNT(S.IDSESSAO)" & vbNewLine
   Sql = Sql & "     From OSESSAO S" & vbNewLine
   Sql = Sql & "     Join OATENDIMENTO_VENDA AV On AV.IDLOJA=S.IDLOJA" & vbNewLine
   Sql = Sql & "                     And AV.IDATENDIMENTO=S.IDATENDIMENTO" & vbNewLine
   Sql = Sql & "                     And AV.IDVENDA=S.IDVENDA" & vbNewLine
   Sql = Sql & "                     And S.IDTPSERVICO<>1" & vbNewLine
   Sql = Sql & "     Where S.IDLOJA = V.IDLOJA" & vbNewLine
   Sql = Sql & "     AND S.IDVENDA=V.IDVENDA),0)" & vbNewLine
   Sql = Sql & "   ) as Integer) [SALDO]" & vbNewLine
   If bPorItem Then
      Sql = Sql & ", IV.NMPROD [PRODUTO]" & vbNewLine
   Else
      Sql = Sql & ", IV.DSCDESC [PRODUTO]" & vbNewLine
   End If
   Sql = Sql & ", PR.DTFIM" & vbNewLine
   
   Sql = Sql & " From CVENDA V" & vbNewLine
   Sql = Sql & " Join CITENSVENDA IV on V.IDLOJA=IV.IDLOJA and V.IDVENDA=IV.IDVENDA" & vbNewLine
   Sql = Sql & " Join OCLIENTE C On C.IDLOJA=V.IDLOJA And V.IDCLIENTE=C.IDCLIENTE" & vbNewLine
   Sql = Sql & " Join OTRATAMENTO_PROD TP on TP.IDLOJA=IV.IDLOJA and TP.IDPROD=IV.IDPROD" & vbNewLine
'   Sql = Sql & " Join OSESSAO SS on SS.IDLOJA=TP.IDLOJA And (TP.IDTPTRATAMENTO=SS.IDTPTRATAMENTO or TP.IDTPTRATAMENTO=0)" & vbNewLine
   
   '* Validade do Pacote
   Sql = Sql & " Left Join CPACOTE PC on IV.IDLOJA=PC.IDLOJA And IV.IDVENDA=PC.IDVENDA And IV.IDPACOTE=PC.IDPACOTE" & vbNewLine
   Sql = Sql & " Left Join CPROMOCAO PR on PC.IDLOJA=PR.IDLOJA And PC.IDPROMO=PR.IDPROMO" & vbNewLine
      
   Sql = Sql & " Where V.IDLOJA= " & mvarIDLOJA & vbNewLine
   Sql = Sql & " And Not V.IDVENDA In (" & sIDVENDA & ")" & vbNewLine
''   If bAteDTVENDA Then Sql = Sql & " And V.DTVENDA<=" & SqlDate(Format(mvarMe.CmbDTATEND.Value, "dd/mm/yyyy 23:59:59")) & vbNewLine
   If bCliente Then Sql = Sql & " And V.IDCLIENTE =" & mvarIDCLIENTE & vbNewLine
   
   '* Para vincular não pode restringir IDATENDIMENTO
'      If mvarIDATENDIMENTO <> 0 Then
'         Sql = Sql & " And SS.IDATENDIMENTO=" & mvarIDATENDIMENTO & vbNewLine
'      End If
   
   '******
   '* Validade do Pacote
   Sql = Sql & " And isNull(PR.DTFIM, Cast('01/01/9999' As DateTime))>=convert(datetime,convert(nvarchar(10), getdate() ,103))" & vbNewLine
   
   Sql = Sql & " And (Cast(( IsNull((SELECT SUM(I.QTDVENDA)" & vbNewLine
   Sql = Sql & "      FROM CITENSVENDA I" & vbNewLine
   Sql = Sql & "      JOIN SPRODUTO P ON I.IDLOJA=P.IDLOJA And I.IDPROD=P.IDPROD And P.ESERVICO=1 And P.EVENDA=1" & vbNewLine
   Sql = Sql & "      Where i.IDLOJA = V.IDLOJA" & vbNewLine
   Sql = Sql & "      AND I.IDVENDA=V.IDVENDA),0)" & vbNewLine
   Sql = Sql & "    -" & vbNewLine
   Sql = Sql & "      IsNull((SELECT COUNT(S.IDSESSAO)" & vbNewLine
   Sql = Sql & "      From OSESSAO S" & vbNewLine
   Sql = Sql & "      Join OATENDIMENTO_VENDA AV On AV.IDLOJA=S.IDLOJA" & vbNewLine
   Sql = Sql & "                      And AV.IDATENDIMENTO=S.IDATENDIMENTO" & vbNewLine
   Sql = Sql & "                      And AV.IDVENDA=S.IDVENDA" & vbNewLine
   Sql = Sql & "                      And S.IDTPSERVICO<>1" & vbNewLine
   Sql = Sql & "      Where S.IDLOJA = V.IDLOJA" & vbNewLine
   Sql = Sql & "      AND S.IDVENDA=V.IDVENDA),0)" & vbNewLine
   Sql = Sql & "    ) as Integer))>0" & vbNewLine
   Sql = Sql & " Order By [DTVENDA] Desc, [IDVENDA] Desc"
   'Sql = Sql & " Order By DTVENDA Desc, V.IDVENDA Desc"
      
   GetSqlSaldoVenda0 = Sql
End Function
Private Function GetSqlSaldoVenda1(Optional bCliente = True, Optional bAteDTVENDA As Boolean) As String
   Dim Sql        As String
   Dim sMsg       As String
   Dim sIDVENDA   As String
   Dim i          As Integer
   Dim bPorItem   As Boolean
   
   bPorItem = False
   sIDVENDA = "0"
   'For i = 1 To mvarMe.GrdVenda.RowCount
   '   If xVal(mvarMe.GrdVenda.CellValue(i, "VENDA")) <> 0 Then
   '      sIDVENDA = sIDVENDA & ", " & xVal(mvarMe.GrdVenda.CellValue(i, "VENDA"))
   '   End If
   'Next
   
   'sMsg = "Apenas o cliente corrente? " & vbNewLine & vbNewLine & "''" & mvarNOME & "''"
   'If Not bCliente Then bCliente = (vbYes = ExibirPergunta(sMsg, "Saldo de Vendas"))

      
   Sql = ""
   Sql = Sql & " SELECT Right('000000'+Rtrim(Cast(I.IDVENDA As Char)),6) [VENDA]" & vbNewLine
   Sql = Sql & ", Cast(Convert(Char(10),V.DTVENDA, 103) As SmallDatetime) [DTVENDA]" & vbNewLine
   Sql = Sql & ", C.NOME" & vbNewLine
   Sql = Sql & ", I.NMPROD [PRODUTO]"
   Sql = Sql & ", SUM(DISTINCT I.QTDVENDA)-COUNT(S.IDSESSAO) [Saldo]" & vbNewLine
   Sql = Sql & ", V.VLVENDA, V.VLPGTO, V.IDMOEDA, I.IDPROD" & vbNewLine
   

   
   Sql = Sql & " FROM CITENSVENDA I" & vbNewLine
   Sql = Sql & " JOIN CVENDA V ON I.IDLOJA=V.IDLOJA And I.IDVENDA=V.IDVENDA" & vbNewLine
   Sql = Sql & " Join OCLIENTE C On C.IDLOJA=V.IDLOJA And V.IDCLIENTE=C.IDCLIENTE" & vbNewLine
   Sql = Sql & " JOIN SPRODUTO P ON I.IDLOJA=P.IDLOJA And I.IDPROD=P.IDPROD And P.ESERVICO=1 And P.EVENDA=1" & vbNewLine
   
   Sql = Sql & " Left Join OATENDIMENTO_VENDA AV On AV.IDLOJA=V.IDLOJA And AV.IDVENDA=V.IDVENDA"
   Sql = Sql & " Left Join OATENDIMENTO A On A.IDLOJA=AV.IDLOJA And A.IDATENDIMENTO=AV.IDATENDIMENTO"
   Sql = Sql & " Left Join OTRATAMENTO_PROD TP ON TP.IDLOJA=P.IDLOJA And TP.IDPROD=I.IDPROD"
   Sql = Sql & " Left Join OSESSAO S On   A.IDLOJA=S.IDLOJA And A.IDATENDIMENTO=S.IDATENDIMENTO AND S.IDTPSERVICO>1 AND S.IDVENDA=I.IDVENDA"
   Sql = Sql & " And S.IDTPSERVICO=TP.IDTPSERVICO"
   Sql = Sql & " And (S.IDTPTRATAMENTO=TP.IDTPTRATAMENTO OR TP.IDTPTRATAMENTO=0)"
   Sql = Sql & " And (S.IDAREA=TP.IDAREA OR TP.IDAREA=0)"
   
   
   Sql = Sql & " Where V.IDLOJA= " & mvarIDLOJA & vbNewLine
   Sql = Sql & " And Not V.IDVENDA In (" & sIDVENDA & ")" & vbNewLine
   If bAteDTVENDA Then Sql = Sql & " And V.DTVENDA<=" & SqlDate(Format(mvarTbAtend.dtatend, "dd/mm/yyyy 23:59:59")) & vbNewLine
   If bCliente Then Sql = Sql & " And V.IDCLIENTE =" & mvarIDCLIENTE & vbNewLine
   
   Sql = Sql & " GROUP BY I.IDLOJA, I.IDVENDA, V.DTVENDA, C.NOME, I.IDITEM, I.NMPROD" & vbNewLine
   Sql = Sql & ", V.VLVENDA, V.VLPGTO, V.IDMOEDA, I.IDPROD" & vbNewLine
   Sql = Sql & " Having SUM(DISTINCT i.QTDVENDA)-COUNT(S.IDSESSAO)>0" & vbNewLine
   Sql = Sql & " ORDER BY I.IDVENDA,I.IDITEM"
   
   GetSqlSaldoVenda1 = Sql
End Function
Public Function GetSqlSaldoVenda(Optional bCliente = True, Optional bAteDTVENDA As Boolean, Optional pIDVENDA As Long) As String
   Dim Sql        As String
   Dim sMsg       As String
   Dim sIDVENDA   As String
   Dim i          As Integer
   Dim bPorItem   As Boolean
   
   bPorItem = False
   sIDVENDA = "0"
   
   Sql = ""
   Sql = Sql & " SELECT Right('000000'+Rtrim(Cast(I.IDVENDA As Char)),6) [IDVENDA]" & vbNewLine
   Sql = Sql & " , I.IDITEM" & vbNewLine
   Sql = Sql & " , Cast(Convert(Char(10),V.DTVENDA, 103) As SmallDatetime) [DTVENDA]" & vbNewLine
   Sql = Sql & " , C.NOME" & vbNewLine
   Sql = Sql & " , I.NMPROD [PRODUTO]" & vbNewLine
   Sql = Sql & " , I.QTDVENDA-(Select Count(*) From OSESSAO S Where S.IDLOJA=I.IDLOJA And S.IDVENDA=I.IDVENDA And S.IDITEM=I.IDITEM) [Saldo]" & vbNewLine
   Sql = Sql & " , V.VLITENS, V.VLVENDA, V.VLPGTO, V.IDMOEDA, I.IDPROD" & vbNewLine
   Sql = Sql & " FROM CITENSVENDA I" & vbNewLine
   Sql = Sql & " JOIN CVENDA V ON I.IDLOJA=V.IDLOJA And I.IDVENDA=V.IDVENDA" & vbNewLine
   Sql = Sql & " Join OCLIENTE C On C.IDLOJA=V.IDLOJA And V.IDCLIENTE=C.IDCLIENTE" & vbNewLine
   Sql = Sql & " JOIN SPRODUTO P ON I.IDLOJA=P.IDLOJA And I.IDPROD=P.IDPROD And P.ESERVICO=1 And P.EVENDA=1" & vbNewLine
   
   If mvarMe.ChkVendas.Value = xtpChecked Then
      Sql = Sql & " Join OTRATAMENTO_PROD TP On TP.IDLOJA=P.IDLOJA AND TP.IDPROD=P.IDPROD" & vbNewLine
      Sql = Sql & " Join OSESSAO S On S.IDLOJA=TP.IDLOJA" & vbNewLine
      Sql = Sql & " And (S.IDTPTRATAMENTO=TP.IDTPTRATAMENTO )" & vbNewLine
      Sql = Sql & " And (S.IDAREA=TP.IDAREA Or TP.IDAREA=0)" & vbNewLine
      Sql = Sql & " And S.IDATENDIMENTO=" & mvarIDATENDIMENTO & vbNewLine
   End If
 
   Sql = Sql & " Where V.IDLOJA= " & mvarIDLOJA & vbNewLine
   If pIDVENDA <> 0 Then
      Sql = Sql & " And V.IDVENDA In (" & pIDVENDA & ")" & vbNewLine
   End If
   Sql = Sql & " And Not V.IDVENDA In (" & sIDVENDA & ")" & vbNewLine
   If bAteDTVENDA Then Sql = Sql & " And V.DTVENDA<=" & SqlDate(Format(mvarTbAtend.dtatend, "dd/mm/yyyy 23:59:59")) & vbNewLine
   If bCliente Then Sql = Sql & " And V.IDCLIENTE =" & mvarIDCLIENTE & vbNewLine
   Sql = Sql & " GROUP BY I.IDLOJA, I.IDVENDA, V.DTVENDA, C.NOME, I.IDITEM, I.NMPROD, I.QTDVENDA" & vbNewLine
   Sql = Sql & ", V.VLITENS, V.VLVENDA, V.VLPGTO, V.IDMOEDA, I.IDPROD" & vbNewLine
   If pIDVENDA = 0 Then
      Sql = Sql & " Having I.QTDVENDA>(Select Count(*) From OSESSAO S Where S.IDLOJA=I.IDLOJA And S.IDVENDA=I.IDVENDA And S.IDITEM=I.IDITEM)" & vbNewLine
   End If
   Sql = Sql & "  ORDER BY DTVENDA" & vbNewLine
   
   GetSqlSaldoVenda = Sql
End Function
Private Sub DefCliente()
   Dim Sql As String
   Sql = ""
   If mvarIDLOJA > 0 Then
      If mvarIDCLIENTE <> 0 Then
         Sql = Sql & "Select C.IDCLIENTE, C.NOME"
         Sql = Sql & " From OCLIENTE C"
         Sql = Sql & " Where A.IDLOJA=" & mvarIDLOJA
         Sql = Sql & " And A.IDCLIENTE=" & mvarIDCLIENTE
      
      ElseIf mvarIDCLIENTE = 0 And mvarIDATENDIMENTO <> 0 Then
         Sql = Sql & "Select C.IDCLIENTE, C.NOME"
         Sql = Sql & " From OCLIENTE C"
         Sql = Sql & " Join OATENDIMENTO A On A.IDLOJA=C.IDLOJA And A.IDCLIENTE=C.IDCLIENTE"
         Sql = Sql & " Where A.IDLOJA=" & mvarIDLOJA
         Sql = Sql & " And A.IDATENDIMENTO=" & mvarIDATENDIMENTO
      End If
   End If
   If Sql <> "" Then
      If mvarSys.xDb.AbreTabela(Sql) Then
         mvarIDCLIENTE = mvarSys.xDb.RsAux("IDCLIENTE")
         mvarNOME = mvarSys.xDb.RsAux("NOME")
      End If
   End If
End Sub
Private Sub DefAtend()
   If mvarIDATENDIMENTO <> 0 And mvarIDLOJA <> 0 Then
      Set mvarTbAtend = CriarObjeto("BANCO_3R.TB_OATENDIMENTO")
      Set mvarTbAtend.xDb = mvarSys.xDb
      Call mvarTbAtend.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDATENDIMENTO:=mvarIDATENDIMENTO)
   End If
End Sub

Private Sub mvarMe_LstVendasItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
   Dim i    As Integer
   Dim sId  As String
   Dim sItem As String
   Dim sProduto As String
   Dim sIDPROD As String
   Dim bExiste As Boolean
   Dim sKey    As String
   
   sId = Item.Text
   sItem = Item.SubItems(1)
   sProduto = Item.ListSubItems("PRODUTO")
   sIDPROD = Item.ListSubItems("IDPROD")
   sKey = sId & "/" & sItem
   
   If Item.Checked Then
      '* Adicionar Item no Combo
      If Not ExisteItem(CollVendas, sKey) Then
         bExiste = False
         For i = 0 To mvarMe.GrdItens.Combos("VENDA").ListCount - 1
            If mvarMe.GrdItens.Combos("VENDA").ItemText(i) = sKey Then
               bExiste = True
               i = mvarMe.GrdItens.Combos("VENDA").ListCount
            End If
         Next
         If Not bExiste Then
            mvarMe.GrdItens.Combos("VENDA").AddItem sKey, sKey
         End If
      End If
      
      For i = 1 To mvarMe.GrdItens.RowCount
         If InStr(mvarMe.GrdItens.CellValue(i, "IDPROD"), sIDPROD & "|") <> 0 Then
            If xVal(Item.ListSubItems(7).Text) > xVal(GetTag(sQVenda, sKey, 0)) Then
               If IsNull(mvarMe.GrdItens.CellValue(i, "VENDA")) Then
                  sItem = ""
               Else
                  sItem = mvarMe.GrdItens.CellValue(i, "VENDA")
               End If
               If Not ExisteItem(CollVendas, sItem) Then
                  If Not InArray(sItem, Array("", " ", "0", 0, "000000/0")) Then
                     sId = mvarMe.GrdItens.CellValue(i, "VENDA")
                     Call SetTag(sQVenda, sId, xVal(GetTag(sQVenda, sId, 0)) - 1)
                  End If
                  Call SetTag(sQVenda, sKey, 1 + xVal(GetTag(sQVenda, sKey, 0)))
                  mvarMe.GrdItens.CellValue(i, "VENDA") = sKey
               End If
            End If
         End If
      Next
      'mvarMe.GrdItens.Redraw = True
      
   Else
      If ExisteItem(CollVendas, sKey) Then
         Item.Checked = True
      Else
         For i = 0 To mvarMe.GrdItens.Combos("VENDA").ListCount - 1
            If mvarMe.GrdItens.Combos("VENDA").ItemText(i) = sKey Then
               mvarMe.GrdItens.Combos("VENDA").RemoveItem i
               i = mvarMe.GrdItens.Combos("VENDA").ListCount
            End If
         Next
         For i = 1 To mvarMe.GrdItens.RowCount
            If mvarMe.GrdItens.CellValue(i, "VENDA") = sKey Then
               Call SetTag(sQVenda, sKey, xVal(GetTag(sQVenda, sKey, 0)) - 1)
               mvarMe.GrdItens.CellValue(i, "VENDA") = ""
            End If
         Next
      End If
   End If
End Sub
Private Function AjustaRelVendas(Optional ComIntegridade As Boolean = True, Optional ApagaRel) As Long
   Dim Sql     As String
   Dim MyRS    As Object
   Dim RsItem  As Object
   Dim nItem   As Integer
   Dim nVenda  As Integer
   Dim bOk     As Boolean
   Dim QtdErro As Integer
   Dim cQuery  As Collection
   Dim sCap    As String
   Dim SqlAtend As String
   
   QtdErro = 0
   
   Sql = " Select S.*, A.IDCLIENTE" & vbNewLine
   Sql = Sql & " From OSESSAO  S" & vbNewLine
   Sql = Sql & " Join OATENDIMENTO A On S.IDLOJA=A.IDLOJA And S.IDATENDIMENTO=A.IDATENDIMENTO" & vbNewLine
   Sql = Sql & " Where S.IDLOJA = " & mvarIDLOJA & vbNewLine
   Sql = Sql & " And (S.IDVENDA Is Null Or S.IDVENDA=0 Or S.IDITEM Is Null Or S.IDITEM=0)" & vbNewLine
   Sql = Sql & " And S.IDTPSERVICO > 1" & vbNewLine
   Sql = Sql & " And S.IDATENDIMENTO <> " & mvarIDATENDIMENTO & vbNewLine
   Sql = Sql & " And DateDiff(dd, A.DTATEND, GetDate())>=1" & vbNewLine
   Sql = Sql & " Order By S.IDLOJA, S.IDATENDIMENTO, S.IDSESSAO"
   If mvarSys.xDb.AbreTabela(Sql, MyRS) Then
      If ComIntegridade Then
         bOk = (vbYes = ExibirPergunta("Existem " & MyRS.RecordCount & " atendimento(s) pendente(s)." & vbNewLine & vbNewLine & "Deseja corrigir agora?"))
      Else
         bOk = True
      End If
      If bOk Then
         sCap = mvarMe.Caption
                                          
         SqlAtend = "Select Distinct A.IDATENDIMENTO "
         SqlAtend = SqlAtend & " From OSESSAO  S"
         SqlAtend = SqlAtend & " Join OATENDIMENTO A On S.IDLOJA=A.IDLOJA And S.IDATENDIMENTO=A.IDATENDIMENTO"
         SqlAtend = SqlAtend & " Where S.IDLOJA =" & MyRS("IDLOJA")
         SqlAtend = SqlAtend & " And (IsNull(S.IDVENDA,0)=0"
         'SqlAtend = SqlAtend & " Or isNull(S.IDITEM,0)="
         SqlAtend = SqlAtend & " )" & vbNewLine
         SqlAtend = SqlAtend & " And S.IDTPSERVICO > 1"

         If mvarSys.xDb.ExisteReg(SqlAtend) Then
            If ComIntegridade Then
               If IsMissing(ApagaRel) Then
                  bOk = (vbYes = ExibirPergunta("Apagar os relacionamentos?"))
               Else
                  If ApagaRel = True Then
                     bOk = (vbYes = ExibirPergunta("Apagar os relacionamentos?"))
                  Else
                     bOk = False
                  End If
               End If
            Else
               bOk = True
            End If
            If bOk Then
               Sql = "Delete "
               Sql = Sql & " From OATENDIMENTO_VENDA "
               Sql = Sql & " Where IDLOJA=" & MyRS("IDLOJA")
               Sql = Sql & " And IDATENDIMENTO in (" & SqlAtend & ")"
               If mvarSys.xDb.Executa(Sql) Then
                  Sql = Sql
               End If
            End If
         End If
         
         Sql = " Select S.*, A.IDCLIENTE" & vbNewLine
         Sql = Sql & " From OSESSAO  S" & vbNewLine
         Sql = Sql & " Join OATENDIMENTO A On S.IDLOJA=A.IDLOJA And S.IDATENDIMENTO=A.IDATENDIMENTO" & vbNewLine
         Sql = Sql & " Where S.IDLOJA = " & mvarIDLOJA & vbNewLine
         Sql = Sql & " And (isNull(S.IDVENDA,0)=0 Or isNull(S.IDITEM,0)=0)" & vbNewLine
         Sql = Sql & " And S.IDTPSERVICO > 1" & vbNewLine
         'Sql = Sql & " And S.IDATENDIMENTO <> " & mvarIDATENDIMENTO & vbNewLine
         Sql = Sql & " Order By S.IDLOJA, S.IDATENDIMENTO, S.IDSESSAO"
         While Not MyRS.EOF
            nItem = 0
            nVenda = 0
            mvarMe.Caption = MyRS.AbsolutePosition & " / " & MyRS.RecordCount
            
            Sql = "Select I.IDLOJA, I.IDVENDA, I.IDITEM, I.QTDVENDA" & vbNewLine
            Sql = Sql & " From CITENSVENDA I" & vbNewLine
            Sql = Sql & " Join SPRODUTO P On P.IDLOJA=I.IDLOJA And P.IDPROD=I.IDPROD And P.ESERVICO=1" & vbNewLine
            Sql = Sql & " Where I.IDLOJA =" & MyRS("IDLOJA") & vbNewLine
            Sql = Sql & " And I.IDVENDA=" & MyRS("IDVENDA") & vbNewLine
            Sql = Sql & " Group By I.IDPROD, I.IDLOJA, I.IDVENDA, I.IDITEM, I.QTDVENDA" & vbNewLine
            Sql = Sql & " Having I.QTDVENDA>(Select Count(*) From OSESSAO S Where S.IDLOJA=I.IDLOJA And S.IDVENDA=I.IDVENDA And S.IDITEM=I.IDITEM)"
            If mvarSys.xDb.AbreTabela(Sql, RsItem) Then
               nItem = RsItem("IDITEM")
               
               Sql = "Update OSESSAO "
               Sql = Sql & " Set IDITEM = " & nItem
               Sql = Sql & " Where IDLOJA=" & MyRS("IDLOJA")
               Sql = Sql & " AND IDATENDIMENTO=" & MyRS("IDATENDIMENTO")
               Sql = Sql & " AND IDSESSAO=" & MyRS("IDSESSAO")
               bOk = mvarSys.xDb.Executa(Sql)
               If Not bOk Then
                  QtdErro = QtdErro + 1
               End If
            Else
               Set cQuery = New Collection
               '* Venda não existe ou não há mais saldo "
               Sql = "Select I.IDVENDA,I.IDITEM, I.IDPROD, I.QTDVENDA, IsNull(A.IDATENDIMENTO,0) [IDATENDIMENTO]" & vbNewLine
               Sql = Sql & " , (Select Count(*) From OSESSAO S Where S.IDLOJA=I.IDLOJA And S.IDVENDA=I.IDVENDA And S.IDITEM=I.IDITEM) [SALDO]" & vbNewLine
               Sql = Sql & " From CITENSVENDA I" & vbNewLine
               Sql = Sql & " Join CVENDA V On I.IDLOJA=V.IDLOJA And I.IDVENDA=V.IDVENDA " & vbNewLine
               Sql = Sql & " Join SPRODUTO S On I.IDLOJA=S.IDLOJA And I.IDPROD=S.IDPROD And S.ESERVICO=1" & vbNewLine
               Sql = Sql & " Left Join OATENDIMENTO_VENDA A  On V.IDLOJA=A.IDLOJA And V.IDVENDA=A.IDVENDA" & vbNewLine
               Sql = Sql & " And A.IDATENDIMENTO=" & MyRS("IDATENDIMENTO") & vbNewLine
               Sql = Sql & " Where V.IDLOJA=" & MyRS("IDLOJA")
               If ComIntegridade Then
                  Sql = Sql & " And V.IDCLIENTE=" & MyRS("IDCLIENTE")
               End If
               Sql = Sql & " Group By I.IDLOJA, A.IDATENDIMENTO, V.DTVENDA, I.IDVENDA,I.IDITEM, I.IDPROD, I.QTDVENDA,  A.IDLOJA" & vbNewLine
               Sql = Sql & " Having I.QTDVENDA>(Select Count(*) From OSESSAO S Where S.IDLOJA=I.IDLOJA And S.IDVENDA=I.IDVENDA And S.IDITEM=I.IDITEM)" & vbNewLine
               Sql = Sql & " Order By IDATENDIMENTO, V.DTVENDA, I.IDVENDA, I.IDITEM" & vbNewLine
               If mvarSys.xDb.AbreTabela(Sql, RsItem) Then
                  nItem = 0
                  While Not RsItem.EOF
                     Sql = Sql & " SELECT * FROM OTRATAMENTO_PROD"
                     Sql = Sql & " Where IDPROD =" & RsItem("IDPROD")
                     Sql = Sql & " And (IDTPSERVICO=" & SqlNum(MyRS("IDTPSERVICO" & "")) & " or IDTPSERVICO=0)"
                     Sql = Sql & " And (IDTPTRATAMENTO=" & SqlNum(MyRS("IDTPTRATAMENTO") & "") & " or IDTPTRATAMENTO=0)"
                     Sql = Sql & " And (IDAREA=" & SqlNum(MyRS("IDAREA") & "") & " or IDAREA=0)"
                     If mvarSys.xDb.ExisteReg(Sql) Then
                        nItem = RsItem("IDITEM")
                        nVenda = RsItem("IDVENDA")
                        RsItem.MoveLast
                     End If
                     If nItem = 0 Then
                        nItem = RsItem("IDITEM")
                        nVenda = RsItem("IDVENDA")
                     End If
                     RsItem.MoveNext
                  Wend
               End If
   
               '* Verificar relacionamento
               If nItem = 0 Then
                  QtdErro = QtdErro + 1
               Else
                  Sql = "Select * "
                  Sql = Sql & " From OATENDIMENTO_VENDA "
                  Sql = Sql & " Where IDLOJA=" & MyRS("IDLOJA")
                  Sql = Sql & " And IDATENDIMENTO=" & MyRS("IDATENDIMENTO")
                  Sql = Sql & " And IDVENDA=" & nVenda
                  If Not mvarSys.xDb.ExisteReg(Sql) Then
                     Sql = "Insert Into OATENDIMENTO_VENDA (IDLOJA, IDVENDA, IDATENDIMENTO, ALTERSTAMP" & vbNewLine
                     Sql = Sql & ", TIMESTAMP) " & vbNewLine
                     Sql = Sql & " Values " & vbNewLine
                     Sql = Sql & "(" & vbNewLine
                     Sql = Sql & MyRS("IDLOJA") & vbNewLine
                     Sql = Sql & ", " & nVenda & vbNewLine
                     Sql = Sql & ", " & MyRS("IDATENDIMENTO") & vbNewLine
                     Sql = Sql & ", " & "1" & vbNewLine
                     Sql = Sql & ", " & "GetDate()" & vbNewLine
                     Sql = Sql & ")" & vbNewLine
                     cQuery.Add Sql
                  End If
                  Sql = "Update OSESSAO "
                  Sql = Sql & " Set IDITEM = " & nItem
                  Sql = Sql & " , IDVENDA= " & nVenda
                  Sql = Sql & " Where IDLOJA=" & MyRS("IDLOJA")
                  Sql = Sql & " AND IDATENDIMENTO=" & MyRS("IDATENDIMENTO")
                  Sql = Sql & " AND IDSESSAO=" & MyRS("IDSESSAO")
                  cQuery.Add Sql
                  
                  bOk = mvarSys.xDb.Executa(cQuery)
                  If Not bOk Then
                     QtdErro = QtdErro + 1
                  End If
               End If
            End If
            MyRS.MoveNext
         Wend
         If mvarSys.xDb.ExisteReg(SqlAtend) Then
            If vbYes = ExibirPergunta("Ainda Existem atendimentos pendentes." & vbNewLine & vbNewLine & "Deseja repetir sem integridade?") Then
               QtdErro = AjustaRelVendas(False)
            End If
         End If
         Call PopulaTela
         mvarMe.Caption = sCap
         AjustaRelVendas = QtdErro
         If MyRS.RecordCount > 0 And ComIntegridade Then
            If QtdErro = 0 Then
               ExibirInformacao MyRS.RecordCount & " atendimento(s) atualizado(s) com sucesso!!"
            Else
               ExibirInformacao "Total: " & MyRS.RecordCount & vbNewLine & "Erros: " & QtdErro
            End If
         End If
      End If
   End If
End Function
