VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tl_VincularAtend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private WithEvents mvarMe     As FrmVincularAtend
Attribute mvarMe.VB_VarHelpID = -1
Dim m_oFont       As StdFont

Private mvarSys      As Object
Private mvarIDLOJA   As Integer
Private mvarIDVENDA  As Integer
Private mvarIDCLIENTE As Integer
Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
   mvarIDLOJA = mvarSys.Propriedades("IDLOJA")
End Property
Public Property Get Sys() As Object
   Set Sys = mvarSys
End Property
Public Property Let IDLOJA(ByVal vData As Integer)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Integer
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDVENDA(ByVal vData As Integer)
   mvarIDVENDA = vData
End Property
Public Property Get IDVENDA() As Integer
   IDATEND = mvarIDVENDA
End Property
Public Property Let IDCLIENTE(ByVal vData As Integer)
   mvarIDCLIENTE = vData
End Property
Public Property Get IDCLIENTE() As Integer
   IDCLIENTE = mvarIDCLIENTE
End Property
Public Sub Show()
   mvarMe.Show vbModal
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmVincularAtend
End Sub
Private Sub Class_Terminate()
   On Error Resume Next
   Unload mvarMe
   Set mvarMe = Nothing
End Sub
Private Sub PopulaTela()
   Dim Sql     As String
   Dim MyRs    As Object

   DoEvents
   
   Sql = ""
   Sql = Sql & " Select I.IDITEM, I.NMPROD, I.QTDVENDA, I.UNIDCONTROLE" & vbNewLine
   Sql = Sql & " , I.QTDVENDA-(Select Count(*) " & vbNewLine
   Sql = Sql & "                 From OSESSAO S" & vbNewLine
   Sql = Sql & "                 Where S.IDLOJA=I.IDLOJA" & vbNewLine
   Sql = Sql & "                 And S.IDVENDA=I.IDVENDA" & vbNewLine
   Sql = Sql & "                 And S.IDITEM=I.IDITEM) [Saldo]" & vbNewLine
   Sql = Sql & " FROM CITENSVENDA I" & vbNewLine
   Sql = Sql & " JOIN CVENDA V ON I.IDLOJA=V.IDLOJA And I.IDVENDA=V.IDVENDA" & vbNewLine
   Sql = Sql & " JOIN SPRODUTO P ON I.IDLOJA=P.IDLOJA And I.IDPROD=P.IDPROD And P.ESERVICO=1 And P.EVENDA=1" & vbNewLine
   Sql = Sql & " Where I.IDLOJA=" & mvarIDLOJA & vbNewLine
   Sql = Sql & " And I.IDVENDA=" & mvarIDVENDA & vbNewLine
   
   'Set TbItem = CriarObjeto("BANCO_3R.TB_CITENSVENDA")
   'Set TbItem.xDb = mvarSys.xDb
   'If TbItem.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=mvarIDVENDA, Ch_ORDERBY:="IDITEM") Then
   
   If mvarSys.xDb.AbreTabela(Sql, MyRs) Then
      mvarMe.GrdVenda.FillFromRS MyRs
   End If
   If mvarMe.GrdVenda.CurRow = 0 Then
      If mvarMe.GrdVenda.RowCount > 0 Then
         mvarMe.GrdVenda.CurRow = 1
      End If
   End If

End Sub
Private Sub PopulaGrdAtend()
   Dim Sql     As String
   Dim MyRs    As Object
   Dim TbItem  As Object 'TB_CITENSVENDA
   Dim i As Integer

   DoEvents
   
   If mvarMe.GrdVenda.CurRow <> 0 Then
      If xVal(GetTag(mvarMe.GrdVenda, "CURROW")) = mvarMe.GrdVenda.CurRow Then Exit Sub
      Call SetTag(mvarMe.GrdVenda, "CURROW", mvarMe.GrdVenda.CurRow)
      
      Sql = ""
      Sql = Sql & "Set dateformat 'dmy';" & vbNewLine
      Sql = Sql & "Select Row_Number() Over(Order By (Case D.IDATENDIMENTO When 0 Then Cast(Convert(varchar(8), I.TIMESTAMP,3) As SmallDateTime) ELSE Cast(Convert(varchar(8),D.DTATEND,3) As SmallDateTime) end), D.IDATENDIMENTO) As LIN, D.IDATENDIMENTO" & vbNewLine
      Sql = Sql & ", DTATEND=Case D.IDATENDIMENTO When 0 Then Cast(Convert(varchar(8), I.TIMESTAMP,3) As SmallDateTime) ELSE Cast(Convert(varchar(8),D.DTATEND,3) As SmallDateTime) end" & vbNewLine
      Sql = Sql & ", I.IDSESSAO" & vbNewLine
      'Sql = Sql & ", T.DSCTRATAMENTO" & vbNewLine
      'Sql = Sql & ", DSCTRATAMENTO=Case D.IDATENDIMENTO When 0 Then(Select T2.DSCTRATAMENTO From OTPTRATAMENTO T2 WHERE I.IDLOJA=T2.IDLOJA And T2.IDTPTRATAMENTO=I.IDTPTRATAMENTO) ELSE T.DSCTRATAMENTO end" & vbNewLine
      Sql = Sql & ", DSCTRATAMENTO=Case D.IDATENDIMENTO When 0 Then '--' ELSE T.DSCTRATAMENTO end" & vbNewLine
      
      Sql = Sql & ", NOME=Case D.IDATENDIMENTO When 0 Then(Select C2.NOME From OCLIENTE C2 WHERE D.IDLOJA=C2.IDLOJA And C2.IDCLIENTE=" & mvarIDCLIENTE & ") ELSE C.NOME end" & vbNewLine
      Sql = Sql & ", ISNULL(A.DSCAREA,'--') AS DSCAREA" & vbNewLine
      Sql = Sql & ", ISNULL(A.DSCAREA,'--') AS DSCAREA" & vbNewLine
      Sql = Sql & " From OSESSAO I " & vbNewLine
      Sql = Sql & " Join OATENDIMENTO D On I.IDLOJA=D.IDLOJA And I.IDATENDIMENTO=D.IDATENDIMENTO" & vbNewLine
      Sql = Sql & " Left Join OCLIENTE C On D.IDLOJA=C.IDLOJA And D.IDCLIENTE=C.IDCLIENTE" & vbNewLine
      Sql = Sql & " Left Join OTPTRATAMENTO T  On I.IDLOJA=T.IDLOJA And I.IDTPTRATAMENTO=T.IDTPTRATAMENTO" & vbNewLine
      Sql = Sql & " Left Join OAREA A On I.IDLOJA=A.IDLOJA And I.IDAREA=A.IDAREA" & vbNewLine
      Sql = Sql & " Where I.IDLOJA = " & mvarIDLOJA & vbNewLine
      Sql = Sql & " And I.IDVENDA= " & mvarIDVENDA & vbNewLine
      Sql = Sql & " And I.IDITEM= " & mvarMe.GrdVenda.CellValue(mvarMe.GrdVenda.CurRow, "IDITEM")
      If mvarSys.xDb.AbreTabela(Sql, MyRs) Then
         mvarMe.GrdAtend.FillFromRS MyRs
         
         mvarMe.GrdAtend.ColWidth("IDATENDIMENTO") = 50
         mvarMe.GrdAtend.ColWidth("DTATEND") = 65
         mvarMe.GrdAtend.ColWidth("IDSESSAO") = 35
         mvarMe.GrdAtend.ColWidth("DSCTRATAMENTO") = 140
         mvarMe.GrdAtend.ColWidth("DSCAREA") = 40
         mvarMe.GrdAtend.ColWidth("NOME") = 60
                                       
      End If
   End If
   Call IncluiGrdLinhaAtendimento
   mvarMe.GrdAtend.Header.AutoHeight
   mvarMe.CmdOk.Enabled = False
End Sub
Private Sub IncluiGrdLinhaAtendimento(Optional pBuscaAtendimento As Boolean = False)
   Dim Sql     As String
   Dim SqlIn   As String
   Dim MYADD   As Object 'DsAdd
   Dim n       As Object
   Dim i       As Integer
   Dim MySel   As Collection
   Dim bRedraw As Boolean
   Dim lRow    As Long
   Dim nColDefault As Integer
   Dim bCliente As Boolean
   Dim sMsg    As String
   
   nColDefault = 2
   With mvarMe.GrdAtend
       bRedraw = .Redraw
      '.Header.Font.Size = 9
      .Header.Font.Name = mvarMe.Font.Name
      .Font.Size = .Header.Font.Size
      .Font.Name = .Header.Font.Name
      
      lRow = .RowCount
      If lRow > 0 Then
         If .CellValue(lRow, "LIN") = 0 Then
            .RemoveRow lRow
         End If
      End If
'      If lRow > 1 Then
'         .CellForeColor(lRow - 1, "IDATENDIMENTO") = vbBlack
'         '.CellFont(lRow - 1, "IDATENDIMENTO").Underline = True
'      End If
      
      'If Not (mvarIDATENDIMENTO = 0 And lRow = 0) Then
'      If lRow > 0 Then
         .AddRow
         lRow = .RowCount
         .RowHeight(.RowCount) = .DefaultRowHeight
   
         For i = 1 To .ColCount
            If .CellType(lRow, .ColKey(i)) = igCellCombo Then
               If .Combos(.ColKey(i)).ListCount = 1 Then
                  .CellValue(lRow, .ColKey(i)) = .Combos(.ColKey(i)).ItemValue(0)
               ElseIf .Combos(.ColKey(i)).ListCount > 1 Then
                  .CellValue(lRow, .ColKey(i)) = GetTag(.ColTag(.ColKey(i)), "DEFAULT", "")
               End If
            Else
               If GetTag(.ColTag(.ColKey(i)), "DEFAULT", "") = "IDENTITY" Then
                  .CellValue(lRow, .ColKey(i)) = lRow
               Else
                  .CellValue(lRow, i) = GetTag(.ColTag(.ColKey(i)), "DEFAULT", "")
               End If
            End If
         Next
'      End If
      
      If .RowCount >= 1 And pBuscaAtendimento Then
         sMsg = "Apenas o cliente corrente? "
         bCliente = (vbYes = ExibirPergunta(sMsg, "Atendimentos / Faturas"))
      
         Sql = "Select Distinct 'Atend.' [TIPO], Right(REPLICATE('0'," + CStr(mvarSys.Propriedades("LENIDVENDA")) + ")+ Cast(A.IDATENDIMENTO as varchar)," + CStr(mvarSys.Propriedades("LENIDVENDA")) + ") [IDATENDIMENTO]" & vbNewLine
         Sql = Sql & " , Convert(varchar, A.DTATEND, 103) [DATA], C.NOME, '--' [VALOR], Count(S.IDSESSAO) [Itens]" & vbNewLine
         Sql = Sql & " From OATENDIMENTO A" & vbNewLine
         Sql = Sql & " Left Join OSESSAO S               On A.IDLOJA=S.IDLOJA And A.IDATENDIMENTO=S.IDATENDIMENTO And S.IDTPSERVICO IN (Select Distinct IDTPSERVICO From OTRATAMENTO_PROD)" & vbNewLine
         Sql = Sql & " Left Join OATENDIMENTO_PRODUTO P  On P.IDLOJA=A.IDLOJA And P.IDATENDIMENTO=A.IDATENDIMENTO" & vbNewLine
         Sql = Sql & " Left Join OCLIENTE C              On A.IDLOJA=C.IDLOJA And A.IDCLIENTE=C.IDCLIENTE" & vbNewLine
         Sql = Sql & " Left Join OATENDIMENTO_VENDA V    On A.IDLOJA=V.IDLOJA And A.IDATENDIMENTO=V.IDATENDIMENTO" & vbNewLine
         Sql = Sql & " Where A.IDLOJA=" & mvarIDLOJA & vbNewLine
         Sql = Sql & " And V.IDVENDA Is Null" & vbNewLine
         Sql = Sql & " And (Not S.IDTPSERVICO Is Null Or Not P.IDPROD Is Null)" & vbNewLine
         If bCliente Then Sql = Sql & " And A.IDCLIENTE=" & mvarIDCLIENTE & vbNewLine
         If mvarMe.GrdAtend.RowCount > 1 Then
            SqlIn = ""
            For i = 1 To mvarMe.GrdAtend.RowCount - 1
               SqlIn = SqlIn & IIf(SqlIn = "", "", ", ")
               SqlIn = SqlIn & Val(mvarMe.GrdAtend.CellValue(i, "IDATENDIMENTO"))
            Next
            If Trim(SqlIn) <> "" Then
               Sql = Sql & " And A.IDATENDIMENTO Not In ( " & SqlIn & " )" & vbNewLine
            End If
         End If
         Sql = Sql & " Group By A.IDATENDIMENTO, A.DTATEND,  C.NOME" & vbNewLine
         Sql = Sql & " Union All" & vbNewLine
         Sql = Sql & " Select Distinct 'Fat.' [TIPO],  Right(REPLICATE('0',6)+ Cast(F.IDFATURA as varchar),6) [IDATENDIMENTO]" & vbNewLine
         Sql = Sql & " , Convert(varchar, F.DTEMISSAO, 103) [DATA]" & vbNewLine
         Sql = Sql & " , C2.NOME, 'R$ ' + replace(cast(F.VALOR as varchar), '.', ',') [VALOR], 0 [Itens]" & vbNewLine
         Sql = Sql & " From FFATURA F" & vbNewLine
         Sql = Sql & " Left Join OCLIENTE C2 On F.IDLOJA=C2.IDLOJA And F.IDCLIENTE=C2.IDCLIENTE" & vbNewLine
         Sql = Sql & " Where F.IDLOJA=" & mvarIDLOJA & vbNewLine
         Sql = Sql & " And (F.IDVENDA Is Null Or F.IDVENDA = 0)" & vbNewLine
         If bCliente Then Sql = Sql & " And F.IDCLIENTE=" & mvarIDCLIENTE
         If mvarMe.GrdAtend.RowCount > 1 Then
            SqlIn = ""
            For i = 1 To mvarMe.GrdAtend.RowCount - 1
               SqlIn = SqlIn & IIf(SqlIn = "", "", ", ")
               SqlIn = SqlIn & Val(mvarMe.GrdAtend.CellValue(i, "IDATENDIMENTO"))
            Next
            If Trim(SqlIn) <> "" Then
               Sql = Sql & " And F.IDFATURA Not In ( " & SqlIn & " )" & vbNewLine
            End If
         
         End If
         Sql = Sql & " Order By IDATENDIMENTO"
         
         'Set MyAdd = New DsAdd
         Set MYADD = CriarObjeto("XActive.DsAdd")
         With MYADD
            .xDb = mvarSys.xDb
            .Caption = "Atendimentos e Faturas Sem Venda."
            .Chaves.Add "TIPO", "TIPO"
            .Chaves.Add "IDATENDIMENTO", "IDATENDIMENTO"
            .Chaves.Add "DATA", "DATA"
            .Chaves.Add "NOME", "NOME"
            .Chaves.Add "VALOR", "VALOR"
            .Chaves.Add "ITENS", "ITENS"
            With .TITULO
               '& 0=eAlin.Esquerda, 1=eAlin.Direita, 2=eAlin.Centro
               Call .Add("TIPO", "Doc.", 8, 0)
               Call .Add("IDATENDIMENTO", "Nº", 7, 0)
               Call .Add("DATA", "Data", 10, 2)
               Call .Add("NOME", "Nome", 25, 0)
               Call .Add("VALOR", "Valor", 10, 1)
               Call .Add("ITENS", "Itens", 5, 2)
            End With
            .CheckBox = True
            .QUERY = Sql
      
            .Show
            Set MySel = .Selecao
         End With
         
         .RemoveRow lRow
         If Not MySel Is Nothing Then
            If .CellValue(.RowCount, 1) = "Clique para incluir nova linha." Then
               .RemoveRow .RowCount
            End If
            For Each n In MySel
               .AddRow
               On Error Resume Next
               .CellValue(.RowCount, "TIPO") = n("TIPO")
               .CellValue(.RowCount, "IDATENDIMENTO") = n("IDATENDIMENTO")
               .CellValue(.RowCount, "DATA") = n("DATA")
               .CellValue(.RowCount, "IDCLIENTE") = n("NOME")
               .CellValue(.RowCount, "VALOR") = xVal(n("VALOR"))
               For i = 1 To .ColCount
                  If GetTag(.ColTag(i), "EDIT", True) = False Then
                     '.CellForeColor(.RowCount, i) = vbGrayText
                  End If
               Next
               
            Next
'            If MySel.Count > 0 Then
'               Call PopulaGRDPRODUTO
'               Call CalculaVenda
'               Call PopulaValores
'            End If
         End If
         If .CellValue(.RowCount, 1) <> "Clique para incluir nova linha." Then
            .AddRow
         End If
      End If
      
      
     
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
      
      On Error Resume Next
      If lRow <> 0 Then
         .SetCurCell .RowCount, 1
         .SetCurCell lRow, nColDefault
      End If
      If .Visible Then
         .SetFocus
      End If
      .Redraw = bRedraw
   End With
End Sub
Private Sub mvarMe_Active()
   Call MontaTela
   Call PopulaTela
   
   If GetTag(mvarMe, "1VEZ", 1) = 1 Then
      '* Definir foco
      With mvarMe.GrdVenda
         If .Visible And .Enabled Then
            .SetFocus
         End If
      End With
      Call SetTag(mvarMe, "1VEZ", 0)
   End If
End Sub
Private Sub mvarMe_CmdCancelClick()
   Unload mvarMe
End Sub
Private Sub mvarMe_CmdOkClick()
   Call SalvarItens
   Unload mvarMe
End Sub
Private Sub SalvarItens()
   Dim Sql As String
   Dim i As Integer
   Dim Queries As Collection
   Dim nIDITEM As Integer
   
   Set Queries = New Collection
   nIDITEM = mvarMe.GrdVenda.CellValue(mvarMe.GrdVenda.CurRow, "IDITEM")
   
   Sql = "Update OSESSAO " & vbNewLine
   Sql = Sql & " Set IDVENDA= Null" & vbNewLine
   Sql = Sql & " , IDITEM= Null" & vbNewLine
   Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA) & vbNewLine
   Sql = Sql & " And IDVENDA=" & SqlNum(mvarIDVENDA) & vbNewLine
   Sql = Sql & " And IDITEM=" & SqlNum(nIDITEM) & vbNewLine
   Queries.Add Sql
   For i = 1 To mvarMe.GrdAtend.RowCount
      Sql = "Update OSESSAO " & vbNewLine
      Sql = Sql & " Set IDVENDA= " & mvarIDVENDA & vbNewLine
      Sql = Sql & " , IDITEM=" & SqlNum(nIDITEM) & vbNewLine
      Sql = Sql & ", ALTERSTAMP=1" & vbNewLine
      Sql = Sql & ", TIMESTAMP = GetDate()" & vbNewLine
      Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA) & vbNewLine
      Sql = Sql & " And IDATENDIMENTO=" & SqlNum(mvarMe.GrdAtend.CellValue(i, "IDATENDIMENTO")) & vbNewLine
      Sql = Sql & " And IDSESSAO=" & SqlNum(mvarMe.GrdAtend.CellValue(i, "IDSESSAO"))
      Queries.Add Sql
   Next
   Call mvarSys.xDb.Executa(Queries)
End Sub

Private Sub mvarMe_GrdAtendMouseEnter(ByVal lRow As Long, ByVal lCol As Long)
   If mvarMe.GrdAtend.ColKey(lCol) = "IDATENDIMENTO" Then
      mvarMe.GrdAtend.CellFont(lRow, lCol) = m_oFont
      'mvarMe.MousePointer = vbCrosshair
      'mvarMe.MouseIcon = vbCrosshair
      'If ExisteArquivo("C:\Windows\Cursors\aero_link.cur") Then
      '   mvarMe.MouseIcon = LoadPicture("C:\Windows\Cursors\aero_link.cur")
      'End If
   End If
End Sub

Private Sub mvarMe_GrdAtendMouseLeave(ByVal lRow As Long, ByVal lCol As Long)
   On Error Resume Next
   If mvarMe.GrdAtend.ColKey(lCol) = "IDATENDIMENTO" Then
      mvarMe.GrdAtend.CellDefaultFont lRow, lCol
      mvarMe.MousePointer = vbDefault
   End If
End Sub

Private Sub mvarMe_GrdAtendMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
   With mvarMe.GrdAtend
      If lRow = .RowCount Then
         If bVendaFechada Then
            If vbYes = ExibirPergunta("Venda fechada." & vbNewLine & "Deseja reabrir venda?") Then
               'Call mvarMe_CmdChaveClick
            End If
            If bVendaFechada Then
               Exit Sub
            End If
         End If
      
         'If lRow > 1 Then
         '   If Val(.CellValue(lRow - 1, "IDATENDIMENTO")) = 0 Then Exit Sub
         'End If
         Call IncluiGrdLinhaAtendimento(True)
         
         If lRow > 0 Then
            If xVal(.CellValue(lRow, "IDATENDIMENTO")) = 0 Then
               Call .SetCurCell(lRow, "IDATENDIMENTO")
            End If
         End If
      Else
         If mvarMe.GrdAtend.ColKey(lCol) = "IDATENDIMENTO" Then
            If lRow > 0 Then
               Call mvarMe_GrdAtendMouseLeave(lRow, lCol)
               Call ExibirAtendimento(xVal(mvarMe.GrdAtend.CellValue(lRow, "IDATENDIMENTO")))
            End If
            
         ElseIf mvarMe.GrdAtend.ColKey(lCol) = "ITENS" Then
            If m_lBtnRow > 0 And m_lBtnCol > 0 Then ' if the user has clicked a cell
               mvarMe.GrdAtend.CellValue(m_lBtnRow, m_lBtnCol) = Empty ' redraw the pushed button
               'For i = 1 To lCol - 1
               '   If mvarMe.GrdAtend.ColVisible(i) Then
               '      m_lBtnRight = m_lBtnRight + mvarMe.GrdAtend.ColWidth(i)
               '   End If
               'Next
              
               ' checking whether we need to raise ButtonClick:
               If (x >= m_lBtnRight - BTN_SIZE) And (x < m_lBtnRight - 1) And (y >= m_lBtnTop + 1) And (y < m_lBtnTop + BTN_SIZE) Then
                  ' --- here you can place code for the ButtonClick event:
                  ' ! Don't use modal dialog because VB can't process the iGrid1_CustomDrawCell event
                  ' if you will move the modal dialog over your form
                  Call ExibirItensAtend(lRow)
               End If
            End If
         End If
      End If
   End With
End Sub
Private Sub ExibirAtendimento(pIDATEND As Integer)
   Dim TlAtend As Object
   
   If pIDATEND = 0 Then Exit Sub
   
   Screen.MousePointer = vbHourglass
   Set TlAtend = CriarObjeto("Calendario3R.TL_Atendimento")
   With TlAtend
      Set .Sys = mvarSys
      'Set .CalControl = mvarMe.CalendarControl
      'Set .CalEvent = ModEvent
      .IDLOJA = mvarIDLOJA
      '.IDEVENTO = ModEvent.Id
      '.IDSALA = ModEvent.ScheduleID
      .IDATENDIMENTO = pIDATEND
      .Show vbModal
   End With
   Set TlAtend = Nothing
   Screen.MousePointer = vbDefault
End Sub
Private Sub ExibirItensAtend(lRow As Long)
   Dim MyItens As TL_ITENSATEND
   
   If mvarMe.GrdAtend.RowCount <= 1 Then Exit Sub
   If mvarMe.GrdAtend.CellValue(lRow, "TIPO") <> "Atend." Then Exit Sub
   
   Set MyItens = New TL_ITENSATEND
   With MyItens
      Set .Sys = mvarSys
      .IDLOJA = mvarIDLOJA
      .IDATEND = xVal(mvarMe.GrdAtend.CellValue(lRow, "IDATENDIMENTO"))
      .DeFaultIDVENDA = xVal(mvarIDVENDA)
      
      .Show
      
      mvarMe.GrdAtend.Redraw = False
      Call PopulaGrdAtend
      On Error Resume Next
      'Call mvarMe.GrdAtendimento.EnsureVisibleCell(lRow, "TIPO")
      Call mvarMe.GrdAtend.SetCurCell(lRow, "TIPO")
      mvarMe.GrdAtend.Redraw = True
   End With
Exit Sub
   
'   Set MyItens = Nothing
'
'   Dim MYADD   As Object
'   Dim Sql     As String
'   Dim MySel   As Collection
'   Dim MyPre   As Collection
'   Dim MyRs    As Object
'   Dim i       As Integer
'   Dim nIDATEND As Integer
'   Dim n       As Object
'
'   If lRow <= 0 Then Exit Sub
'
'   nIDATEND = xval(mvarMe.GrdAtendimento.CellValue(lRow, "IDATENDIMENTO"))
'   Sql = "Select IDLOJA, IDATENDIMENTO, IDSESSAO, IDTPSERVICO, IDTPTRATAMENTO, IDAREA, IDVENDA"
'   Sql = "Select IDSESSAO, IDTPSERVICO, IDTPTRATAMENTO, IDAREA, IDVENDA"
'   Sql = Sql & " From OSESSAO"
'   Sql = Sql & " Where IDLOJA = " & mvarIDLOJA
'   Sql = Sql & " And IDATENDIMENTO= " & nIDATEND
'
'   If mvarsys.xdb.AbreTabela(Sql, MyRs) Then
'      Set MyPre = New Collection
'      While Not MyRs.EOF
'         If Trim(MyRs("IDVENDA") & "") <> "" Then
'            If Not ExisteItem(MySel, "k" & (MyRs("IDSESSAO") & "")) Then
'               MyPre.Add MyRs("IDSESSAO") & "", "k" & MyRs("IDSESSAO") & ""
'            End If
'         End If
'         MyRs.MoveNext
'      Wend
'      Set MYADD = CriarObjeto("XActive.DsAdd")
'      With MYADD
'         .xdb = mvarsys.xdb
'         .Caption = "Itens do Atendimento."
'         '.Chaves.Add "IDLOJA", "IDLOJA"
'         '.Chaves.Add "IDATENDIMENTO", "IDATENDIMENTO"
'         .Chaves.Add "IDSESSAO", "IDSESSAO"
'         '.Chaves.Add "IDTPSERVICO", "IDTPSERVICO"
'         '.Chaves.Add "IDTPTRATAMENTO", "IDTPTRATAMENTO"
'         '.Chaves.Add "IDAREA", "IDAREA"
'
'         With .TITULO
'            Call .Add("IDSESSAO", "Item", 8, 0) 'eAlin.Esquerda)
'            Call .Add("IDTPSERVICO", "Serviço", 10, 2) 'eAlin.Centro)
'            Call .Add("IDTPTRATAMENTO", "Tratamento", 10, 2) 'eAlin.Centro)
'            Call .Add("IDAREA", "Área", 20, 0) 'eAlin.Esquerda)
'         End With
'         .CheckBox = True
'         .QUERY = Sql
'
'         Set .PreSelecao = MyPre
'
'         .Show
'         Set MySel = .Selecao
'
'
'         Sql = "Update OSESSAO "
'         Sql = Sql & " Set IDVENDA= Null"
'         Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
'         Sql = Sql & " And IDATENDIMENTO=" & SqlNum(nIDATEND)
'         Sql = Sql & " And IDVENDA=" & SqlNum(mvarIDVENDA)
'         Call mvarsys.xdb.Executa(Sql)
'         If MySel.Count > 0 Then
'            For Each n In MySel
'               Sql = "Update OSESSAO "
'               Sql = Sql & " Set IDVENDA= " & mvarIDVENDA
'               Sql = Sql & " Where IDLOJA=" & SqlNum(mvarIDLOJA)
'               Sql = Sql & " And IDATENDIMENTO=" & SqlNum(nIDATEND)
'               Sql = Sql & " And IDSESSAO=" & SqlNum(n("IDSESSAO"))
'               Call mvarsys.xdb.Executa(Sql)
'            Next
'         End If
'      End With
'   End If
End Sub

Private Sub mvarMe_GrdAtendRequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
   With mvarMe.GrdAtend
      bCancel = (.ColKey(lCol) <> "BTN")
      'If .CellFmtString(lRow, lCol) <> "" Then sText = .CellValue(lRow, lCol)
      'eTextEditOpt = GetTag(.ColTag(lCol), "EDITOPT", 0)
   End With
End Sub

Private Sub mvarMe_GrdVendaCurCellChange(ByVal lRow As Long, ByVal lCol As Long)
   Call PopulaGrdAtend
End Sub

Private Sub mvarMe_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF5 Then
      Call SetTag(mvarMe.GrdVenda, "CURROW", 0)
      Call PopulaGrdAtend
   End If
End Sub

Private Sub mvarMe_Load()
   'mvarMe.CmdOk.Enabled = False
   'Call MotaTela
End Sub
Private Sub MontaTela()
   With mvarMe.GrdVenda
      .Editable = False
      .ImmediateColumnResizing = False
      .MultiSelect = True
      .ShowControlsInAllCells = True
      .Header.DragCols = False
      .SilentValidation = True
      .BorderStyle = igBorderThin
      .AutoWidthColWithHeader = True
      .RowMode = True
      '.FocusRect = False
      
      With .AddCol(sKey:="IDITEM", sHeader:="#", lWidth:=30, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextCenter
         .sCtrlKey = "IDITEM"
      End With
      With .AddCol(sKey:="NMPROD", sHeader:="Produto", lWidth:=160, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextLeft Or igTextNoClip
         .sCtrlKey = "NMPROD"
      End With
      With .AddCol(sKey:="QTDVENDA", sHeader:="Qtd.", lWidth:=40, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextRight Or igTextNoClip
         .sCtrlKey = "QTDVENDA"
      End With
      With .AddCol(sKey:="UNIDCONTROLE", sHeader:="Unid.", lWidth:=40, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextCenter Or igTextNoClip
         .sCtrlKey = "UNIDCONTROLE"
      End With
      With .AddCol(sKey:="Saldo", sHeader:="Saldo", lWidth:=40, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextRight Or igTextNoClip
         .sCtrlKey = "Saldo"
      End With
      .ColHeaderTextFlags("IDITEM") = igTextCenter
      .ColHeaderTextFlags("QTDVENDA") = igTextRight
      .ColHeaderTextFlags("Saldo") = igTextRight
   End With
   
   With mvarMe.GrdAtend
      '.Editable = False
      .ImmediateColumnResizing = True
      .MultiSelect = True
      .ShowControlsInAllCells = True
      .Header.DragCols = False
      .SilentValidation = True
      .BorderStyle = igBorderThin
      .AutoWidthColWithHeader = True
      .RowMode = True

      With .AddCol(sKey:="LIN", sHeader:="#", lWidth:=30, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextLeft
         .sCtrlKey = "LIN"
      End With
      With .AddCol(sKey:="IDATENDIMENTO", sHeader:="Atend.", lWidth:=50, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextLeft
         .sCtrlKey = "IDATENDIMENTO"
         .oForeColor = vbBlue
      End With
      With .AddCol(sKey:="DTATEND", sHeader:="Data", lWidth:=65, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextCenter
         .sCtrlKey = "DTATEND"
      End With
      With .AddCol(sKey:="IDSESSAO", sHeader:="Item", lWidth:=35, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextLeft
         .sCtrlKey = "IDSESSAO"
      End With
      With .AddCol(sKey:="DSCTRATAMENTO", sHeader:="Tratamento", lWidth:=100, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextLeft Or igTextNoClip
         .sCtrlKey = "DSCTRATAMENTO"
      End With
      With .AddCol(sKey:="DSCAREA", sHeader:="Área", lWidth:=40, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextLeft Or igTextNoClip
         .sCtrlKey = "DSCAREA"
      End With
      With .AddCol(sKey:="NOME", sHeader:="Cliente", lWidth:=40, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextBoxHScrollBar
         .eTextFlags = igTextLeft Or igTextNoClip
         .sCtrlKey = "NOME"
      End With
      Set m_oFont = New StdFont
      m_oFont.Name = .Font.Name
      m_oFont.Size = .Font.Size
      m_oFont.Underline = True
   End With
   
   
End Sub

