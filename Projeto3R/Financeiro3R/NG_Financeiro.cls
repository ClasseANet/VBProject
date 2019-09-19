VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NG_Financeiro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mvarSys      As Object
Private mvarIDLOJA   As Integer
Private mvarIDCONTA  As Integer

Private mvarBarItem  As Object
Private mvarTLPane   As Object
Private mvarPane     As Object
Private mvarTLCal    As Object
Public Property Set TLCal(ByVal vData As Object)
   Set mvarTLCal = vData
   Set mvarTLCal.Sys = mvarSys
End Property
Public Property Get TLCal() As Object
   Set TLCal = mvarTLCal
End Property
Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
   Set gSys = vData
   
   If Val(mvarIDCONTA) = 0 Then
      Dim Tb_Conta As Object
      If DefineContaLAN(Tb_Conta, 1, mvarSys) Then
         mvarSys.Propriedades("IDCONTA") = Tb_Conta.IDCONTA
      End If
      Set Tb_Conta = Nothing
   End If
   mvarIDLOJA = mvarSys.Propriedades("IDLOJA")
   mvarIDCONTA = mvarSys.Propriedades("IDCONTA")
   
   Set mvarTLPane.Sys = mvarSys

   'If Not mvarBarItem Is Nothing Then Set mvarBarItem.Sys = mvarSys
   'If Not mvarPane Is Nothing Then Set mvarPane.Sys = mvarSys
   If Not mvarTLCal Is Nothing Then Set mvarTLCal.Sys = mvarSys
   
End Property
Public Property Get Sys() As Object
   Set Sys = mvarSys
End Property
Public Property Set BarItem(ByVal vData As Object)
   Set mvarBarItem = vData
End Property
Public Property Set Pane(ByVal vData As Object)
   Set mvarPane = vData
End Property
Public Property Get Pane() As Object
   Set Pane = mvarPane
End Property
Public Sub Show(Optional Reload As Boolean = False)
   mvarTLPane.Show Reload
End Sub
Public Sub Hide()
   'mvarTLPane.Hide
End Sub
Private Sub Class_Initialize()
   Set mvarTLPane = New TL_PaneCC
   Set mvarPane = mvarTLPane.Pane
End Sub
Public Sub ExibirFatura(ByRef pIDFATURA As Long, Optional pIDVENDAORIGEM, Optional pIDATENDIMENTO, Optional pIDCLIENTE As Long, Optional ByRef pVLFATURA)
   Dim TlFatura As Object
   Dim TbAtend As Object
   Dim NgFinanc As Object
   Dim AtendSalvo As Boolean
   Dim nMouse As Integer
  
   On Error GoTo Saida
            
   nMouse = Screen.MousePointer
   Screen.MousePointer = vbHourglass
   
   Set TlFatura = CriarObjeto("Financ3R.TL_Fatura")
   With TlFatura
      Set .Sys = mvarSys
      .IDLOJA = mvarIDLOJA
      .IDFATURA = pIDFATURA
      If Not IsMissing(pIDCLIENTE) Then .IDCLIENTE = pIDCLIENTE
      If Not IsMissing(pIDATENDIMENTO) Then .IDATENDIMENTO = pIDATENDIMENTO
      If Not IsMissing(pVLFATURA) Then .VLFATURA = pVLFATURA
      If Not IsMissing(pIDVENDAORIGEM) Then .IDVENDAORIGEM = pIDVENDAORIGEM
      
      .Show vbModal
      pIDFATURA = .IDFATURA
      If Not IsMissing(pVLFATURA) And pIDFATURA <> 0 Then
         pVLFATURA = .VLFATURA
      End If
   End With
   Set TlFatura = Nothing

Saida:
   Screen.MousePointer = nMouse
End Sub
Public Sub ExibirVenda(pCod As Long, Optional pCalControl, Optional pIDATENDIMENTO As Integer, Optional pIDCLIENTE As Integer)
   Dim TbVenda As Object
   Dim TbAtend As Object
   Dim TlVenda As TL_Venda
   
   Set TbVenda = CriarObjeto("BANCO_3R.TB_CVENDA")
   Set TbVenda.xDb = mvarSys.xDb
   If TbVenda.PESQUISAR(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=pCod) Then
      Set TlVenda = New TL_Venda
      With TlVenda
         Set TbAtend = CriarObjeto("BANCO_3R.TB_OATENDIMENTO_VENDA")
         Set TbAtend.xDb = mvarSys.xDb
         If TbAtend.PESQUISAR(Ch_IDLOJA:=mvarIDLOJA, Ch_IDVENDA:=pCod) Then
            .IDATENDIMENTO = TbAtend.IDATENDIMENTO
         End If
         
         Set .Sys = mvarSys
         If Not IsMissing(pCalControl) Then
            Set .CalControl = pCalControl
         End If
         .IDLOJA = TbVenda.IDLOJA
         .IDCLIENTE = TbVenda.IDCLIENTE
         .IDVENDA = TbVenda.IDVENDA
         
         .Show vbModal
      End With
      Set TlVenda = Nothing
      Set TbAtend = Nothing
      Set TbVenda = Nothing
   Else
      If pCod = 0 Then
         Set TlVenda = New TL_Venda
         With TlVenda
            Set .Sys = mvarSys
            If Not IsMissing(pCalControl) Then Set .CalControl = pCalControl
            
            .IDLOJA = mvarIDLOJA
            .IDATENDIMENTO = pIDATENDIMENTO
            .IDCLIENTE = pIDCLIENTE
            .IDVENDA = pCod
            
            .Show vbModal
         End With
      End If
   End If
End Sub
Public Function ArquivoRPS() As String
   Dim Sql As String
   Dim sCod As String
   Dim sArq As String
   
   Sql = "Select IsNull(Max(IDLOTE),0)+1  As LOTE" & vbNewLine
   Sql = Sql & " From FRECIBO " & vbNewLine
   Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
   
   If mvarSys.xDb.AbreTabela(Sql) Then
      sCod = StrZero(mvarSys.xDb.RsAux("LOTE") & "", 3)
      'sCod = Format(mvarsys.xDb.SysDate, "yy") & sCod
               
      sArq = Trim(mvarSys.Propriedades("RZABREV"))
      sArq = sArq & IIf(sArq <> "", "_", "")
      sArq = sArq & "RPS" & sCod & ".txt"
   End If
   ArquivoRPS = sArq
End Function
Public Function BaixarSaldo(pIDVENDA As String, pSALDO As Long, Optional pSALDOAtual As Long) As Long
   Dim TbAtend As Object
   Dim TbAtendVenda As Object
   Dim TbOSessao As Object
   Dim RsAux As Object
   Dim Sql As String
   Dim bOk As Boolean
   Dim i As Integer
   Dim QtdItens As Integer
   Dim Queries As Collection
   
   Set Queries = New Collection
   If pSALDOAtual = 0 Then
      Sql = GetSqlVendaSaldo(mvarIDLOJA, CInt(pIDVENDA))
      If mvarSys.xDb.AbreTabela(Sql, RsAux) Then
         pSALDOAtual = RsAux("Sld")
      End If
   End If
  
   Set TbAtend = CriarObjeto("BANCO_3R.TB_OATENDIMENTO")
   Set TbAtend.xDb = mvarSys.xDb
   If Not TbAtend.PESQUISAR(Ch_IDLOJA:=mvarIDLOJA, Ch_IDATENDIMENTO:=0) Then
      
      'TbAtend.IDLOJA = mvarIDLOJA
      'TbAtend.IDATENDIMENTO = 0
      'TbAtend.DTATEND = "01/01/2000"
      Call mvarSys.xDb.AbreTabela("Select Min(IDCLIENTE) [IDCLIENTE] From OCLIENTE", RsAux)
      'TbAtend.IDCLIENTE = mvarSys.xDb.RsAux(0)
      
      Sql = "Insert Into OATENDIMENTO "
      Sql = Sql & "(IDLOJA, IDATENDIMENTO, DTATEND, IDCLIENTE, IDITEM, ALTERSTAMP, TIMESTAMP)"
      Sql = Sql & " Select " & mvarIDLOJA
      Sql = Sql & " ,0"
      Sql = Sql & ",'2000-01-01'"
      Sql = Sql & ", " & RsAux("IDCLIENTE")
      Sql = Sql & ", 1"
      Sql = Sql & ", 1"
      Sql = Sql & ", GetDate()"
      Queries.Add Sql
      'bOk = mvarsys.xDb.Executa(Sql)
   End If
   
   QtdItens = 0
   Sql = "Select *"
   Sql = Sql & " From OSESSAO"
   Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
   Sql = Sql & " And IDATENDIMENTO=0"
   Sql = Sql & " And (IDVENDA=0 Or IDVENDA is Null)"
   If mvarSys.xDb.AbreTabela(Sql, RsAux) Then
      While Not RsAux.EOF
         Sql = "Update OSESSAO Set" & vbNewLine
         Sql = Sql & "  IDVENDA=" & pIDVENDA & vbNewLine
         Sql = Sql & ", IDITEM=1" & vbNewLine
         Sql = Sql & ", ALTERSTAMP=1" & vbNewLine
         Sql = Sql & ", TIMESTAMP=GetDate()" & vbNewLine
         Sql = Sql & " Where IDLOJA=" & mvarIDLOJA & vbNewLine
         Sql = Sql & " And IDATENDIMENTO=0" & vbNewLine
         Sql = Sql & " And IDSESSAO=" & RsAux("IDSESSAO")
         Queries.Add Sql
         'If mvarsys.xDb.Executa(Sql) Then
            QtdItens = QtdItens + 1
         'End If
         RsAux.MoveNext
      Wend
   End If
   
   QtdItens = pSALDOAtual - pSALDO - QtdItens
   If QtdItens > 0 Then
      Sql = "Select IsNull(Max(IDSESSAO), 0) [IDSESSAO]"
      Sql = Sql & " From OSESSAO"
      Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
      Sql = Sql & " And IDATENDIMENTO=0"
      If mvarSys.xDb.AbreTabela(Sql, RsAux) Then
         For i = 1 To QtdItens
            Sql = "Insert Into OSESSAO "
            Sql = Sql & "(IDLOJA, IDATENDIMENTO, IDSESSAO, IDVENDA, IDITEM, IDTPSERVICO"
            Sql = Sql & ", ALTERSTAMP, TIMESTAMP)"
            Sql = Sql & " Select " & mvarIDLOJA
            Sql = Sql & " ,0"
            Sql = Sql & ", " & RsAux("IDSESSAO") + i
            Sql = Sql & ", " & pIDVENDA
            Sql = Sql & ", 1"
            Sql = Sql & ", 2"
            Sql = Sql & ", 1"
            Sql = Sql & ", GetDate()"
            'bOk = mvarsys.xDb.Executa(Sql)
            Queries.Add Sql
         Next
         Set TbAtendVenda = CriarObjeto("BANCO_3R.TB_OATENDIMENTO_VENDA")
         With TbAtendVenda
            Set .xDb = mvarSys.xDb
            .IDLOJA = mvarIDLOJA
            .IDATENDIMENTO = 0
            .IDVENDA = pIDVENDA
            '.IDITEM = 1
            '.Salvar
            Queries.Add .QrySave
         End With
      End If
   ElseIf QtdItens = 0 And (pSALDO = pSALDOAtual) Then
      Set TbOSessao = CriarObjeto("BANCO_3R.TB_OSESSAO")
      With TbOSessao
         Set .xDb = mvarSys.xDb
         .IDLOJA = mvarIDLOJA
         .IDATENDIMENTO = 0
         .IDSESSAO = 1
         If .PESQUISAR(mvarIDLOJA, 0, 1) Then
            .IDITEM = 1
            .Salvar
            'Queries.Add .QrySave
         End If
      End With
   End If
   If Queries.Count > 0 Then
      If mvarSys.xDb.Executa(Queries) Then
         BaixarSaldo = pSALDO
      End If
   End If
End Function
Public Function GetSqlVendaSaldo(pIDLOJA As Integer, Optional pIDVENDA As Integer, Optional pIDCLIENTE As Integer)
   Dim Sql As String
   
   Sql = "Select Distinct Right(REPLICATE('0',6)+Rtrim(Cast(V.IDVENDA As Char)),6) [Venda]" & vbNewLine
   Sql = Sql & ", Cast(Convert(Char(10),V.DTVENDA, 103) As SmallDatetime) [Data]" & vbNewLine
   Sql = Sql & ", C.NOME [Nome], C.IDCLIENTE" & vbNewLine
   Sql = Sql & ", Right(REPLICATE('0',6)+Rtrim(Cast(R.IDRECIBO As Char)),6) [Recibo]" & vbNewLine
   Sql = Sql & ", Right(REPLICATE('0',6)+Rtrim(Cast(N.NUMNOTA As Char)),6) [N.F.]" & vbNewLine
   Sql = Sql & ", V.VLVENDA+V.VLACRESC [Valor]" & vbNewLine
   Sql = Sql & ", V.VLDESC [Desc.]" & vbNewLine
   Sql = Sql & ", V.VLPGTO-V.VLTROCO [Pgto.]" & vbNewLine
   Sql = Sql & " , Cast(( IsNull((SELECT SUM(I.QTDVENDA)" & vbNewLine
   Sql = Sql & "     FROM CITENSVENDA I" & vbNewLine
   Sql = Sql & "     JOIN SPRODUTO P ON I.IDLOJA=P.IDLOJA And I.IDPROD=P.IDPROD And P.ESERVICO=1 And P.EVENDA=1" & vbNewLine
   Sql = Sql & "     Where i.IDLOJA = V.IDLOJA" & vbNewLine
   Sql = Sql & "     AND I.IDVENDA=V.IDVENDA),0)" & vbNewLine
   Sql = Sql & "   -" & vbNewLine
   Sql = Sql & "     IsNull((SELECT COUNT(S.IDSESSAO)" & vbNewLine
   Sql = Sql & "     From OSESSAO S" & vbNewLine
   Sql = Sql & "     Join OATENDIMENTO_VENDA AV On AV.IDLOJA=S.IDLOJA " & vbNewLine
   Sql = Sql & "                     And AV.IDATENDIMENTO=S.IDATENDIMENTO" & vbNewLine
   Sql = Sql & "                     And AV.IDVENDA=S.IDVENDA" & vbNewLine
   Sql = Sql & "                     And S.IDTPSERVICO<>1" & vbNewLine
   Sql = Sql & "     Where S.IDLOJA = V.IDLOJA" & vbNewLine
   Sql = Sql & "     AND S.IDVENDA=V.IDVENDA),0)" & vbNewLine
   Sql = Sql & "   ) as Integer) [Sld]" & vbNewLine
   Sql = Sql & " From CVENDA V" & vbNewLine
   Sql = Sql & " Join OCLIENTE C On C.IDLOJA=V.IDLOJA And V.IDCLIENTE=C.IDCLIENTE" & vbNewLine
   Sql = Sql & " Left Join FRECIBO R On V.IDLOJA=R.IDLOJA And V.IDVENDA=R.IDVENDA" & vbNewLine
   Sql = Sql & " Left Join FNOTAFISCAL N On V.IDLOJA=N.IDLOJA And V.IDVENDA=N.IDVENDA And N.FLGCANCELADA=0" & vbNewLine
   Sql = Sql & " Where V.IDLOJA=" & SqlNum(mvarIDLOJA) & vbNewLine
   If xVal(pIDVENDA) <> 0 Then
      Sql = Sql & " And V.IDVENDA=" & pIDVENDA & vbNewLine
   End If
   If xVal(pIDCLIENTE) <> 0 Then
      Sql = Sql & " And V.IDCLIENTE=" & pIDCLIENTE & vbNewLine
   End If
   Sql = Sql & " Union" & vbNewLine
   Sql = Sql & " Select Distinct Right(REPLICATE('0',6)+ Cast(V.IDVENDA as varchar),6) [Venda]" & vbNewLine
   Sql = Sql & ", Cast(Convert(Char(10), V.DTVENDA, 103) As SmallDatetime) [Data]" & vbNewLine
   Sql = Sql & ", C.NOME [Nome], C.IDCLIENTE" & vbNewLine
   Sql = Sql & ", Right(REPLICATE('0',6)+Rtrim(Cast(R.IDRECIBO As Char)),6) [Recibo]" & vbNewLine
   Sql = Sql & ", Right(REPLICATE('0',6)+Rtrim(Cast(N.NUMNOTA As Char)),6) [N.F.]" & vbNewLine
   Sql = Sql & ", V.VLVENDA [Valor]" & vbNewLine
   Sql = Sql & ", V.VLDESC [Desc.]" & vbNewLine
   Sql = Sql & ", V.VLPGTO-V.VLTROCO [Pgto.]" & vbNewLine
   Sql = Sql & " , Cast(( IsNull((SELECT SUM(I.QTDVENDA)" & vbNewLine
   Sql = Sql & "     FROM CITENSVENDA I" & vbNewLine
   Sql = Sql & "     JOIN SPRODUTO P ON I.IDLOJA=P.IDLOJA And I.IDPROD=P.IDPROD And P.ESERVICO=1 And P.EVENDA=1" & vbNewLine
   Sql = Sql & "     Where i.IDLOJA = V.IDLOJA" & vbNewLine
   Sql = Sql & "     AND I.IDVENDA=V.IDVENDA),0)" & vbNewLine
   Sql = Sql & "   -" & vbNewLine
   Sql = Sql & "     IsNull((SELECT COUNT(S.IDSESSAO)" & vbNewLine
   Sql = Sql & "     From OSESSAO S" & vbNewLine
   Sql = Sql & "     Join OATENDIMENTO_VENDA AV On AV.IDLOJA=S.IDLOJA " & vbNewLine
   Sql = Sql & "                     And AV.IDATENDIMENTO=S.IDATENDIMENTO" & vbNewLine
   Sql = Sql & "                     And AV.IDVENDA=S.IDVENDA" & vbNewLine
   Sql = Sql & "                     And S.IDTPSERVICO<>1" & vbNewLine
   Sql = Sql & "     Where S.IDLOJA = V.IDLOJA" & vbNewLine
   Sql = Sql & "     AND S.IDVENDA=V.IDVENDA),0)" & vbNewLine
   Sql = Sql & "   ) as Integer) [Sld]" & vbNewLine
   
   Sql = Sql & " From CVENDA V" & vbNewLine
   Sql = Sql & " Join CITENSVENDA I on V.IDLOJA=I.IDLOJA and V.IDVENDA=I.IDVENDA" & vbNewLine
   Sql = Sql & " Join OATENDIMENTO_VENDA AV on AV.IDLOJA=V.IDLOJA and AV.IDVENDA=V.IDVENDA" & vbNewLine
   Sql = Sql & " Join OATENDIMENTO A on AV.IDLOJA=A.IDLOJA and AV.IDATENDIMENTO=A.IDATENDIMENTO" & vbNewLine
   Sql = Sql & " Join OCLIENTE C on A.IDLOJA=C.IDLOJA and A.IDCLIENTE=C.IDCLIENTE And A.IDATENDIMENTO<>0" & vbNewLine
   Sql = Sql & " Left Join FRECIBO R On V.IDLOJA=R.IDLOJA And V.IDVENDA=R.IDVENDA" & vbNewLine
   Sql = Sql & " Left Join FNOTAFISCAL N On V.IDLOJA=N.IDLOJA And V.IDVENDA=N.IDVENDA And N.FLGCANCELADA=0" & vbNewLine
   Sql = Sql & " Where V.IDLOJA =" & pIDLOJA & vbNewLine
   If xVal(pIDVENDA) <> 0 Then
      Sql = Sql & " And V.IDVENDA=" & pIDVENDA & vbNewLine
   End If
   If xVal(pIDCLIENTE) <> 0 Then
      Sql = Sql & " And V.IDCLIENTE=" & pIDCLIENTE & vbNewLine
   End If
   Sql = Sql & " Order By Data, Venda"
   
   GetSqlVendaSaldo = Sql
End Function
Public Function ExportarRPS(Optional pDTINI As String, Optional pDTFIM As String, Optional bExibe As Boolean = True, Optional pDTFechamento As String) As String
   Dim Sql As String
   Dim MyRs As Recordset
   Dim ObjXML As MSXML2.DOMDocument
   Dim sCod     As String
   Dim sDrive   As String
   Dim sArquivo As String
   Dim sMsg     As String
   Dim sAux     As String
   Dim Queries  As Collection
   Dim nVltotal As Double
      
   Screen.MousePointer = vbHourglass
            
   If IsMissing(pDTFIM) Then pDTFIM = mvarSys.xDb.SysDate
   If Not IsDate(pDTFIM) Then pDTFIM = mvarSys.xDb.SysDate
   If IsMissing(pDTINI) Then pDTINI = ""
   If Not IsDate(pDTINI) Then pDTINI = ""
   Dim QtdDia As Integer
   QtdDia = 15
   If pDTINI = "" Then
      If Day(pDTFIM) > QtdDia Then
         pDTINI = "16/" & Format(pDTFIM, "mm/yyyy")
      Else
         pDTINI = "01/" & Format(pDTFIM, "mm/yyyy")
      End If
   End If
   
   
   
   '*************
   '* Limpar Recibos com valores zerados
   Sql = "Select R.IDLOJA, R.IDRECIBO, R.SERIE"
   Sql = Sql & " From FRECIBO R"
   Sql = Sql & " Where R.IDLOJA=" & mvarIDLOJA
   Sql = Sql & " And VLTOTAL=0"
   Sql = Sql & " And (IDLOTE=0 Or IDLOTE is Null)"
   If mvarSys.xDb.AbreTabela(Sql, MyRs) Then
      While Not MyRs.EOF
         Set Queries = New Collection
         Sql = "Delete From FRECIBO"
         Sql = Sql & " Where IDLOJA=" & SqlNum(MyRs("IDLOJA") & "")
         Sql = Sql & " And IDRECIBO=" & SqlNum(MyRs("IDRECIBO") & "")
         Sql = Sql & " And SERIE=" & SqlStr(MyRs("SERIE") & "")
         Queries.Add Sql
         Sql = "Update FRECIBO"
         Sql = Sql & " Set IDRECIBO=IDRECIBO-1"
         Sql = Sql & " Where IDLOJA=" & SqlNum(MyRs("IDLOJA") & "")
         Sql = Sql & " And SERIE=" & SqlStr(MyRs("SERIE") & "")
         Sql = Sql & " And IDRECIBO>" & SqlNum(MyRs("IDRECIBO") & "")
         Queries.Add Sql
         
         Call mvarSys.xDb.Executa(Queries)
         
         MyRs.MoveNext
         Set Queries = Nothing
      Wend
   End If
   
   Sql = "Select R.IDLOJA, R.IDRECIBO, R.SERIE, R.IDVENDA, Cast(Convert(Char(10),R.DTEMISSAO, 103) As Datetime) [DTEMISSAO]" & vbNewLine
   Sql = Sql & " , R.VLTOTAL, R.TIPO, R.STATUS, R.NATUREZAOP, R.ASSINATURA" & vbNewLine
   Sql = Sql & " , R.REGESPECIALTRIB, R.SIMPLES, R.IDRECIBO0, R.SERIE0, R.TRIBUTACAO" & vbNewLine
   Sql = Sql & " , R.INCENTIVOCULT, R.CODSERVFEDERAL, R.CODSERVMUNIC" & vbNewLine
   Sql = Sql & " , C.NOME, C.REGISTRO, C.ENDERECO, C.BAIRRO, C.CIDADE" & vbNewLine
   Sql = Sql & " , C.ESTADO, C.CEP, C.TEL1, C.EMAIL, C.CIDADE, C.ESTADO" & vbNewLine
   Sql = Sql & " , 'SERVIÇO DE ESTÉTICA' [DSCSERV]" & vbNewLine
   Sql = Sql & " From FRECIBO R" & vbNewLine
   Sql = Sql & " Left Join CVENDA V On V.IDLOJA=R.IDLOJA And V.IDVENDA=R.IDVENDA AND V.FLGCANCELADA<>1" & vbNewLine
   Sql = Sql & " Left Join OCLIENTE C On V.IDLOJA=C.IDLOJA And V.IDCLIENTE=C.IDCLIENTE" & vbNewLine
   Sql = Sql & " Left Join FNOTAFISCAL F On V.IDLOJA=F.IDLOJA And V.IDVENDA=F.IDVENDA" & vbNewLine
   Sql = Sql & " Where R.IDLOJA=" & mvarIDLOJA & vbNewLine
   Sql = Sql & " And R.TIPO=0" & vbNewLine
   Sql = Sql & " And (R.IDLOTE = 0 Or R.IDLOTE is Null)" & vbNewLine
   Sql = Sql & " And (F.NUMNOTA= 0 Or F.NUMNOTA is Null)" & vbNewLine
   Sql = Sql & " And R.DTEMISSAO >=" & SqlDate(Format(pDTINI, "dd/mm/yyyy") & " 00:00:00") & vbNewLine
   Sql = Sql & " And R.DTEMISSAO <=" & SqlDate(Format(pDTFIM, "dd/mm/yyyy") & " 23:59:59") & vbNewLine
   Sql = Sql & " Order By R.DTEMISSAO, R.IDRECIBO"
   
   If mvarSys.xDb.AbreTabela(Sql, MyRs) Then
      sMsg = ""
      sMsg = sMsg & "Data Ini  : " & Format(pDTINI, "dd/mm/yy") & vbNewLine
      sMsg = sMsg & "Data Fim : " & Format(pDTFIM, "dd/mm/yy") & vbNewLine
      sMsg = sMsg & "Qtd.        : " & MyRs.RecordCount & vbNewLine
                      
      sDrive = ResolvePathName(Environ("TEMP") & "\" & mvarSys.CODSIS)
      Call CriarDiretorio(sDrive)
      
      sArquivo = ArquivoRPS
      
      Call ExcluirArquivo(sDrive & Mid(sArquivo, 1, Len(sArquivo) - 3) & "xml")
      ''Set ObjXML = CreateObject("MSXML2.DOMDocument")
      'Set ObjXML = New DOMDocument
      'MyRs.Save ObjXML, adPersistXML
      'ObjXML.Save sDrive & sArquivo
      'Set ObjXML = Nothing
            
      Call ExcluirArquivo(sDrive & sArquivo)
      'Close #1
      Open sDrive & sArquivo For Output As #1
      
      Sql = "Select Min(R.DTEMISSAO) [DTINI], Max(R.DTEMISSAO) [DTFIM]"
      Sql = Sql & " From FRECIBO R"
      Sql = Sql & " Left Join CVENDA V On V.IDLOJA=R.IDLOJA And V.IDVENDA=R.IDVENDA AND V.FLGCANCELADA<>1"
      Sql = Sql & " Left Join FNOTAFISCAL F On V.IDLOJA=F.IDLOJA And V.IDVENDA=F.IDVENDA"
      Sql = Sql & " Where R.IDLOJA=" & mvarIDLOJA
      Sql = Sql & " And R.TIPO=0"
      Sql = Sql & " And R.STATUS=1"
      Sql = Sql & " And (R.IDLOTE = 0 Or R.IDLOTE is Null)"
      Sql = Sql & " And (F.NUMNOTA= 0 Or F.NUMNOTA is Null)"
      Sql = Sql & " And R.DTEMISSAO >=" & SqlDate(Format(pDTINI, "dd/mm/yyyy") & " 00:00:00") & vbNewLine
      Sql = Sql & " And R.DTEMISSAO <=" & SqlDate(Format(pDTFIM, "dd/mm/yyyy") & " 23:59:59") & vbNewLine
         
      If mvarSys.xDb.AbreTabela(Sql) Then
         If Not MontarCabecalho(mvarSys.xDb.RsAux("DTINI"), mvarSys.xDb.RsAux("DTFIM")) Then GoTo Fim
         
         MyRs.MoveFirst
         nVltotal = 0
         While Not MyRs.EOF
            If Not MontarDetalhe(MyRs) Then GoTo Fim
            nVltotal = nVltotal + MyRs("VLTOTAL")
            MyRs.MoveNext
         Wend

'         Sql = "Select Sum(R.VLTOTAL) [TOTAL]"
'         Sql = Sql & " From FRECIBO R"
'         Sql = Sql & " Left Join CVENDA V On V.IDLOJA=R.IDLOJA And V.IDVENDA=R.IDVENDA AND V.FLGCANCELADA<>1"
'         Sql = Sql & " Left Join FNOTAFISCAL F On V.IDLOJA=F.IDLOJA And V.IDVENDA=F.IDVENDA"
'         Sql = Sql & " Where R.IDLOJA=" & mvarIDLOJA
'         Sql = Sql & " And R.TIPO=0"
'         Sql = Sql & " And R.STATUS=1"
'         Sql = Sql & " And (R.IDLOTE = 0 Or R.IDLOTE is Null)"
'         Sql = Sql & " And (F.NUMNOTA= 0 Or F.NUMNOTA is Null)"
'         Sql = Sql & " And R.DTEMISSAO >=" & SqlDate(Format(pDTINI, "dd/mm/yyyy") & " 00:00:00") & vbNewLine
'         Sql = Sql & " And R.DTEMISSAO <=" & SqlDate(Format(pDTFIM, "dd/mm/yyyy") & " 23:59:59") & vbNewLine
'         nVltotal = mvarSys.xdb.RsAux("TOTAL")
         nVltotal = xVal(nVltotal, 2)
         If mvarSys.xDb.AbreTabela(Sql) Then
            sMsg = sMsg & "Valor       : " & ValBr(nVltotal) & vbNewLine
         End If
         If Not MontarFimArquivo(MyRs.RecordCount, nVltotal, 0, 0, 0) Then GoTo Fim
      End If
      Close #1
      
      
      ExportarRPS = sDrive & sArquivo
      
      Dim sTem As String
      Call CriarDiretorio(mvarSys.PathTmp & "RPS")
      sTem = mvarSys.PathTmp & "RPS\"
      Call ExcluirArquivo(sTem & sArquivo)
      Call CopiarArquivo(sDrive & sArquivo, sTem & sArquivo)
      
      If bExibe Then
         If ExibirPergunta(sMsg & vbNewLine & "Abrir pasta?", "R.P.S.", vbYes) = vbYes Then
            Call Shell("explorer " & sDrive, vbMaximizedFocus)
         End If
      End If
   Else
      If bExibe Then
         Call ExibirInformacao("Não existe recibo para exportação")
      End If
   End If
   Screen.MousePointer = vbDefault
   Exit Function
Fim:
   Close #1
   Screen.MousePointer = vbDefault
'   Call ShowError
End Function
Private Function MontarCabecalho(sDTINI As String, sDTFIM As String) As Boolean
   Dim MyRs As Object
   Dim Sql As String
   Dim sLinha As String
   Dim sAux As String
   Dim bAtualizou As Boolean
   
   bAtualizou = False
   Sql = "Select CNPJ, INSCMUNIC, RZABREV"
   Sql = Sql & " From OLOJA"
   Sql = Sql & " Where IDLOJA =" & mvarIDLOJA
   Call mvarSys.xDb.AbreTabela(Sql, MyRs)
   If Trim(MyRs("CNPJ") & "") = "" Then
      sAux = InputBox("Informe o CNPJ sem pontos e traços.", "RPS")
      If Trim(sAux) = "" Then
         Call ExibirAviso("CNPJ Inválido")
         Exit Function
      Else
         Sql = "Update OLOJA"
         Sql = Sql & " Set CNPJ=" & SqlStr(sAux)
         Sql = Sql & ", ALTERSTAMP=1"
         Sql = Sql & ", TIMESTAMP=GetDate()"
         Sql = Sql & " Where IDLOJA =" & mvarIDLOJA
         Call mvarSys.xDb.Executa(Sql)
         bAtualizou = True
      End If
   End If
   If Trim(MyRs("INSCMUNIC") & "") = "" Then
      sAux = InputBox("Informe a Insc. Municipal sem pontos e traços.", "RPS")
      If Trim(sAux) = "" Then
         Call ExibirAviso("Insc. Municipal Inválida.")
         Exit Function
      Else
         Sql = "Update OLOJA"
         Sql = Sql & " Set INSCMUNIC=" & SqlStr(sAux)
         Sql = Sql & ", ALTERSTAMP=1"
         Sql = Sql & ", TIMESTAMP=GetDate()"
         Sql = Sql & " Where IDLOJA =" & mvarIDLOJA
         Call mvarSys.xDb.Executa(Sql)
         bAtualizou = True
      End If
   End If
   If Trim(MyRs("RZABREV") & "") = "" Then
      sAux = InputBox("Informe a Nome da Empresa, em apenas 1 nome.", "RPS")
      If Trim(sAux) = "" Then
         Call ExibirAviso("Nome Inválido.")
         Exit Function
      Else
         Sql = "Update OLOJA"
         Sql = Sql & " Set RZABREV=" & SqlStr(sAux)
         Sql = Sql & ", ALTERSTAMP=1"
         Sql = Sql & ", TIMESTAMP=GetDate()"
         Sql = Sql & " Where IDLOJA =" & mvarIDLOJA
         Call mvarSys.xDb.Executa(Sql)
         bAtualizou = True
      End If
   End If
   If bAtualizou Then
      Sql = "Select CNPJ, INSCMUNIC, RZABREV"
      Sql = Sql & " From OLOJA"
      Sql = Sql & " Where IDLOJA =" & mvarIDLOJA
      Call mvarSys.xDb.AbreTabela(Sql, MyRs)
   End If
   
   sLinha = "10"   'Tipo 10 = Cabeçalho
   sLinha = sLinha & "003" 'Versão do Lay-Out
   sLinha = sLinha & "2"   'Tipo de Identificação 1-CPF e 2-CNPJ
   sLinha = sLinha & UnFormat(MyRs("CNPJ"))
   sLinha = sLinha & StrZero(UnFormat(MyRs("INSCMUNIC")), 15)
   sLinha = sLinha & Format(sDTINI, "yyyymmdd")
   sLinha = sLinha & Format(sDTFIM, "yyyymmdd")
   
   sLinha = sLinha '& Chr(13) & Chr(10)
   Print #1, sLinha
   MontarCabecalho = True
End Function
Private Function MontarDetalhe(MyRs As Recordset) As Boolean
   Dim lRs As Object
   Dim Sql As String
   Dim sLinha As String
   Dim sEnd As String
   Dim sAux As String
   
   Sql = "Select CIDADE, ESTADO"
   Sql = Sql & " From OLOJA"
   Sql = Sql & " Where IDLOJA =" & mvarIDLOJA
   Call mvarSys.xDb.AbreTabela(Sql, lRs)
   If Trim(lRs("CIDADE") & "") = "" Then
      sAux = InputBox("Informe a Cidade.", "RPS")
      If Trim(sAux) = "" Then
         Call ExibirAviso("Cidade Inválida.")
         Exit Function
      Else
         Sql = "Update OLOJA"
         Sql = Sql & " Set CIDADE=" & SqlStr(sAux)
         Sql = Sql & ", ALTERSTAMP=1"
         Sql = Sql & ", TIMESTAMP=GetDate()"
         Sql = Sql & " Where IDLOJA =" & mvarIDLOJA
         Call mvarSys.xDb.Executa(Sql)
      End If
   End If
   If Trim(lRs("ESTADO") & "") = "" Then
      sAux = InputBox("Informe a U.F.", "RPS")
      If Trim(sAux) = "" Then
         Call ExibirAviso("Cidade Inválida.")
         Exit Function
      Else
         Sql = "Update OLOJA"
         Sql = Sql & " Set ESTADO=" & SqlStr(sAux)
         Sql = Sql & ", ALTERSTAMP=1"
         Sql = Sql & ", TIMESTAMP=GetDate()"
         Sql = Sql & " Where IDLOJA =" & mvarIDLOJA
         Call mvarSys.xDb.Executa(Sql)
      End If
   End If
   
         
   sLinha = "20"                                                     '01-Tipo de Registro
   sLinha = sLinha & "0"                                             '02-Tipo de RPS 0-RPs 1 RPS-M
   sLinha = sLinha & Left(UnFormat(MyRs("SERIE")) & Space(5), 5)     '03-Série
   sLinha = sLinha & StrZero(MyRs("IDRECIBO"), 15)                   '04-RPS
   sLinha = sLinha & Format(MyRs("DTEMISSAO"), "yyyymmdd")           '05-Data de Emissão
   sLinha = sLinha & MyRs("STATUS")                                  '06-Status RPS 0-Normal 1-Cancelado
   If Len(Trim(UnFormat(MyRs("REGISTRO") & ""))) = 14 Then           '07-Tipo de Registo (CPF e CNPJ
      If ValidarCNPJ(MyRs("REGISTRO")) Then
         sLinha = sLinha & "2"
         sLinha = sLinha & StrZero(UnFormat(MyRs("REGISTRO") & ""), 14) '08-Num. Registro
      Else
         sLinha = sLinha & "3"
         sLinha = sLinha & StrZero("0", 14)
      End If
   ElseIf Len(Trim(UnFormat(MyRs("REGISTRO") & ""))) = 11 Then
      
      If ValidarCPF(MyRs("REGISTRO")) Then
         sLinha = sLinha & "1"
         sLinha = sLinha & StrZero(UnFormat(MyRs("REGISTRO") & ""), 14)
      Else
         sLinha = sLinha & "3"
         sLinha = sLinha & StrZero("0", 14)
      End If
   Else
      sLinha = sLinha & "3"
      sLinha = sLinha & StrZero("0", 14)
   End If
   sLinha = sLinha & Space(15)                                       '09-Insc Munic
   sLinha = sLinha & StrZero("0", 15)                                '10-Insc Est
   sLinha = sLinha & Left(MyRs("NOME") & Space(115), 115)            '11-Nome
   sLinha = sLinha & DefineTPLogradouro(Trim(MyRs("ENDERECO") & "")) '12-Tipo do endereço
   sLinha = sLinha & TrataEndereco(Trim(MyRs("ENDERECO") & ""))      '13-Endereço
   sLinha = sLinha & Space(10)                                       '14-Numero
   sLinha = sLinha & Space(60)                                       '15-Complemento
   sLinha = sLinha & Left(MyRs("BAIRRO") & Space(72), 72)            '16-Bairro
   sLinha = sLinha & Left(MyRs("CIDADE") & Space(50), 50)            '17-Cidade
   sLinha = sLinha & Left(MyRs("ESTADO") & Space(2), 2)              '18-UF
   sLinha = sLinha & Left(MyRs("CEP") & Space(8), 8)                 '19-CEP
   sLinha = sLinha & Left(MyRs("TEL1") & Space(11), 11)              '20-Telefone
   sLinha = sLinha & Left(MyRs("EMAIL") & Space(80), 80)             '21-e-Mail
   
   sLinha = sLinha & "01"                                            '22-Tipo de Tributação 01- No Município
   sLinha = sLinha & Left(lRs("CIDADE") & Space(50), 50)             '23-Cidade do Serviço
   sLinha = sLinha & Left(lRs("ESTADO") & Space(2), 2)               '24-UF do Serviço
   sLinha = sLinha & StrZero(MyRs("REGESPECIALTRIB"), 2)             '25-Regime Especial de Tributação - 00-Normal
   sLinha = sLinha & StrZero(MyRs("SIMPLES"), 1)                     '26-Optante Simples
   sLinha = sLinha & StrZero(MyRs("INCENTIVOCULT"), 1)               '27-Incentivo Cultural
   sLinha = sLinha & StrZero(MyRs("CODSERVFEDERAL"), 4)              '28-Cod. Serv. Federal
   sLinha = sLinha & StrZero(MyRs("CODSERVMUNIC"), 20)               '29-Cod. Serv. Municipal
   sLinha = sLinha & StrZero("0", 5)                                 '30-Aliquota
   sLinha = sLinha & StrZero(UnFormat(ValBr(MyRs("VLTOTAL"))), 15)   '31-Valor
   sLinha = sLinha & StrZero("0", 15)                                '32-Deducao
   sLinha = sLinha & StrZero("0", 15)                                '33-Desconto Condicionado, não altera base de cálculo
   sLinha = sLinha & StrZero("0", 15)                                '34-Desconto Incondicionado, altera base de cálculo
   sLinha = sLinha & StrZero("0", 15)                                '35-COFINS
   sLinha = sLinha & StrZero("0", 15)                                '36-CSLL
   sLinha = sLinha & StrZero("0", 15)                                '37-INSS
   sLinha = sLinha & StrZero("0", 15)                                '38-IRPJ
   sLinha = sLinha & StrZero("0", 15)                                '39-PIS/PASEP
   sLinha = sLinha & StrZero("0", 15)                                '40-Outras retenções Federias
   sLinha = sLinha & StrZero("0", 15)                                '41-ISS
   sLinha = sLinha & "0"                                             '42-0-Nota sem ISS retido, 1-ISS Retido
   sLinha = sLinha & Format(MyRs("DTEMISSAO"), "yyyymmdd")           '43-Data de Competencia
   sLinha = sLinha & Space(15)                                       '44-Código da Obra
   sLinha = sLinha & Space(15)                                       '45-Anotação de Resp. Tecnica
   sLinha = sLinha & Left(UnFormat(MyRs("SERIE0")) & Space(5), 5)    '46-Série de RPs Substituto
   sLinha = sLinha & StrZero(MyRs("IDRECIBO0"), 15)                  '47-RPS Substituto
   sLinha = sLinha & Space(30)                                       '48-Reservado
   sLinha = sLinha & MyRs("DSCSERV") & ""                            '49-Discriminação

   sLinha = sLinha '& Chr(13) & Chr(10)                               '49-Fim de Linha
   Print #1, sLinha
   
   MontarDetalhe = True
End Function
Private Function MontarFimArquivo(pNum As Long, pVl As Double, pVlDed As Double, pVlDescC As Double, pVlDescI As Double) As Boolean
   Dim sLinha As String
   
   sLinha = "90"                                              'Tipo 90 = Rodapé
   sLinha = sLinha & StrZero(pNum, 8)                         'Qtd Detalhes
   sLinha = sLinha & StrZero(UnFormat(ValBr(pVl)), 15)        'Soma dos Valores
   sLinha = sLinha & StrZero(UnFormat(ValBr(pVlDed)), 15)     'Soma das Deduções
   sLinha = sLinha & StrZero(UnFormat(ValBr(pVlDescC)), 15)   'Soma dos Descontos Condicionados
   sLinha = sLinha & StrZero(UnFormat(ValBr(pVlDescI)), 15)   'Soma dos Descontos Incondicionados
   
   sLinha = sLinha '& Chr(13) & Chr(10)
   Print #1, sLinha
   
   MontarFimArquivo = True
End Function
Public Sub ImportarNF(pArquivo As String, Optional pMsg As Boolean = True)
   Dim sValue     As String
   Dim ObjXML     As MSXML2.DOMDocument
   Dim ObjRoot    As MSXML2.IXMLDOMElement
   Dim ObjLista   As MSXML2.IXMLDOMElement
   Dim ObjComp    As MSXML2.IXMLDOMElement
   Dim ObjNotas   As MSXML2.IXMLDOMElement
   Dim ObjNota    As MSXML2.IXMLDOMElement
   
   Dim sDrive     As String
   Dim sArquivo   As String
   Dim cQueries   As Collection
   Dim Sql        As String
   Dim nIDLOTE    As Integer
   Dim NewLote    As Integer
   Dim nCancelada As Integer
   Dim QtdNFs     As Integer
      
   Screen.MousePointer = vbHourglass
   
   nIDLOTE = 0
   NewLote = 0
   If mvarSys.xDb.AbreTabela("Select Max(IDLOTE) From FRECIBO Where IDLOJA=" & mvarIDLOJA) Then
      NewLote = 1 + xVal(mvarSys.xDb.RsAux(0) & "")
   End If
   'sArquivo = "RPS11002.xml"
   'sDrive = "C:\Tmp\"
   'Call CriarDiretorio(sDrive)
   'sArquivo = sDrive & sArquivo
   sArquivo = pArquivo
   
   Set cQueries = New Collection
   If ExisteArquivo(sArquivo) Then
      
      'Set ObjXML = CreateObject("MSXML2.DOMDocument")
      Set ObjXML = New DOMDocument
      
      'carrega o XML no documento DOM
      ObjXML.resolveExternals = True
      ObjXML.validateOnParse = True
      ObjXML.async = False
      
      ObjXML.Load sDrive & sArquivo
      'ObjXML.loadXML sDrive & sArquivo
      
      'verifica se a carga do XML foi feita com sucesso
      If ObjXML.parseError.reason <> "" Then
         MsgBox ObjXML.parseError.reason
         Exit Sub
      End If
      
      'obtem o elemento raiz do XML
      QtdNFs = 0
      Set ObjRoot = ObjXML.documentElement      '* ConsultarNfseResposta
      Set ObjLista = ObjRoot.childNodes(0)      '* ListaNfse
      Set ObjComp = ObjLista.childNodes(0)      '* CompNfse
      For Each ObjComp In ObjLista.childNodes
         Set ObjNotas = ObjComp.childNodes(0)   '* Nfse
         Set ObjNota = ObjNotas.childNodes(0)   '* InfNfse
         If Not ObjNota Is Nothing Then
            Dim TbNota     As Object
            Dim TBRecibo   As Object
            Dim TbLote     As Object
            
            Set TbNota = CriarObjeto("BANCO_3R.TB_FNOTAFISCAL")
            Set TbNota.xDb = mvarSys.xDb

            nCancelada = 0
            If ObjComp.childNodes.length = 3 Then
               nCancelada = 1
            End If
            QtdNFs = QtdNFs + 1
            If Not TbNota.PESQUISAR(Ch_IDLOJA:=mvarIDLOJA, Ch_NUMNOTA:=GetValueXmlNode(ObjNota, "Numero"), Ch_SERIE:=GetValueXmlNode(ObjNota, "IdentificacaoRps/Serie")) Then
               Set TBRecibo = CriarObjeto("BANCO_3R.TB_FRECIBO")
               Set TbLote = CriarObjeto("BANCO_3R.TB_FLOTERPS")
               Set TBRecibo.xDb = mvarSys.xDb
               Set TbLote.xDb = mvarSys.xDb
            
               TbNota.IDLOJA = mvarIDLOJA
               TbNota.NUMNOTA = GetValueXmlNode(ObjNota, "Numero")
               'If TbNota.NUMNOTA > "340" Then
               '   TbNota.NUMNOTA = TbNota.NUMNOTA
               'End If
               TbNota.SERIE = GetValueXmlNode(ObjNota, "IdentificacaoRps/Serie")
               TbNota.IDRECIBO = GetValueXmlNode(ObjNota, "IdentificacaoRps/Numero")
               If IsDate(Format(Replace(GetValueXmlNode(ObjNota, "DataEmissao"), "T", " "), "DD/MM/YYYY")) Then
                  TbNota.DTEMISSAO = Format(Replace(GetValueXmlNode(ObjNota, "DataEmissao"), "T", " "), "DD/MM/YYYY")
               Else
                  Call ExibirStop("Erro na Data de Emissão")
                  Exit Sub
               End If
               TbNota.VlTotal = GetValueXmlNode(ObjNota, "Servico/Valores/ValorServicos")
               TbNota.FLGENVIADA = 1
               TbNota.FLGCANCELADA = nCancelada
               If TBRecibo.PESQUISAR(Ch_IDLOJA:=TbNota.IDLOJA, Ch_IDRECIBO:=TbNota.IDRECIBO, Ch_SERIE:=TbNota.SERIE) Then
                  TbNota.IDVENDA = TBRecibo.IDVENDA
                  If TBRecibo.IDLOTE = 0 Then
                     '** Se a RPS não possui lote registrado, então cria um novo lote de recebimento
                     nIDLOTE = NewLote
                     TbLote.IDLOTE = nIDLOTE
                     TbLote.IDLOJA = TBRecibo.IDLOJA
                     TbLote.NUMLOTE = ""
                     If IsDate(TbLote.DTINI) Then
                        TbLote.DTINI = IIf(CDate(TBRecibo.DTEMISSAO) < CDate(TbLote.DTINI), TBRecibo.DTEMISSAO, TbLote.DTINI)
                     Else
                        TbLote.DTINI = TBRecibo.DTEMISSAO
                     End If
                     If IsDate(TbLote.DTFIM) Then
                        TbLote.DTFIM = IIf(CDate(TBRecibo.DTEMISSAO) > CDate(TbLote.DTFIM), TBRecibo.DTEMISSAO, TbLote.DTFIM)
                     Else
                        TbLote.DTFIM = TBRecibo.DTEMISSAO
                     End If
                     TbLote.DTENVIO = TbNota.DTEMISSAO
                     TbLote.SITLOTE = 1
                     cQueries.Add TbLote.QrySave
                  Else
                     nIDLOTE = TBRecibo.IDLOTE
                     If TbLote.PESQUISAR(Ch_IDLOJA:=mvarIDLOJA, Ch_IDLOTE:=nIDLOTE) Then
                        If IsDate(TbLote.DTINI) Then
                           TbLote.DTINI = Format(IIf(CDate(TBRecibo.DTEMISSAO) < CDate(TbLote.DTINI), TBRecibo.DTEMISSAO, TbLote.DTINI), "dd/mm/yyyy")
                        Else
                           TbLote.DTINI = Format(TBRecibo.DTEMISSAO, "dd/mm/yyyy")
                        End If
                        If IsDate(TbLote.DTFIM) Then
                           TbLote.DTFIM = Format(IIf(CDate(TBRecibo.DTEMISSAO) > CDate(TbLote.DTFIM), TBRecibo.DTEMISSAO, TbLote.DTFIM), "dd/mm/yyyy")
                        Else
                           TbLote.DTFIM = Format(TBRecibo.DTEMISSAO, "dd/mm/yyyy")
                        End If
                        TbLote.DTENVIO = TbNota.DTEMISSAO
                        TbLote.SITLOTE = 1
                        If TbLote.isDirt Then
                           Call TbLote.Salvar
                        End If
                     Else
                        TbLote.IDLOJA = TBRecibo.IDLOJA
                        TbLote.NUMLOTE = ""
                        If IsDate(TbLote.DTINI) Then
                           TbLote.DTINI = IIf(CDate(TBRecibo.DTEMISSAO) < CDate(TbLote.DTINI), TBRecibo.DTEMISSAO, TbLote.DTINI)
                        Else
                           TbLote.DTINI = TBRecibo.DTEMISSAO
                        End If
                        If IsDate(TbLote.DTFIM) Then
                           TbLote.DTFIM = IIf(CDate(TBRecibo.DTEMISSAO) > CDate(TbLote.DTFIM), TBRecibo.DTEMISSAO, TbLote.DTFIM)
                        Else
                           TbLote.DTFIM = TBRecibo.DTEMISSAO
                        End If
                        TbLote.DTENVIO = TbNota.DTEMISSAO
                        TbLote.SITLOTE = 1
                        If TbLote.Salvar Then
                           nIDLOTE = TbLote.IDLOTE
                        End If
                     End If
                  End If
                  TBRecibo.IDLOTE = nIDLOTE
                  If TBRecibo.isDirt Then
                     cQueries.Add TBRecibo.QryUpdate
                  End If
               End If
               If TbNota.isDirt Then
                  cQueries.Add TbNota.QrySave
               End If
               
               If False Then
                  sValue = GetValueXmlNode(ObjNota, "Numero")
                  If sValue = "260" Then
                  sValue = GetValueXmlNode(ObjNota, "CodigoVerificacao")
                  End If
                  sValue = GetValueXmlNode(ObjNota, "DataEmissao")
                  sValue = GetValueXmlNode(ObjNota, "IdentificacaoRps/Numero")
                  sValue = GetValueXmlNode(ObjNota, "IdentificacaoRps/Serie")
                  sValue = GetValueXmlNode(ObjNota, "IdentificacaoRps/Tipo")
                  sValue = GetValueXmlNode(ObjNota, "DataEmissaoRps")
                  sValue = GetValueXmlNode(ObjNota, "NaturezaOperacao")
                  sValue = GetValueXmlNode(ObjNota, "OptanteSimplesNacional")
                  sValue = GetValueXmlNode(ObjNota, "IncentivadorCultural")
                  sValue = GetValueXmlNode(ObjNota, "Competencia")
                  sValue = GetValueXmlNode(ObjNota, "Servico/Valores/ValorServicos")
                  sValue = GetValueXmlNode(ObjNota, "Servico/Valores/IssRetido")
                  sValue = GetValueXmlNode(ObjNota, "Servico/Valores/BaseCalculo")
                  sValue = GetValueXmlNode(ObjNota, "Servico/Valores/ValorLiquidoNfse")
                  sValue = GetValueXmlNode(ObjNota, "Servico/ItemListaServico")
                  sValue = GetValueXmlNode(ObjNota, "Servico/CodigoTributacaoMunicipio")
                  sValue = GetValueXmlNode(ObjNota, "Servico/Discriminacao")
                  sValue = GetValueXmlNode(ObjNota, "Servico/CodigoMunicipio")
                  sValue = GetValueXmlNode(ObjNota, "ValorCredito")
                  sValue = GetValueXmlNode(ObjNota, "PrestadorServico/IdentificacaoPrestador/Cnpj")
                  sValue = GetValueXmlNode(ObjNota, "PrestadorServico/IdentificacaoPrestador/InscricaoMunicipal")
                  sValue = GetValueXmlNode(ObjNota, "PrestadorServico/RazaoSocial")
                  sValue = GetValueXmlNode(ObjNota, "PrestadorServico/NomeFantasia")
                  sValue = GetValueXmlNode(ObjNota, "PrestadorServico/Endereco/Endereco")
                  sValue = GetValueXmlNode(ObjNota, "PrestadorServico/Endereco/Numero")
                  sValue = GetValueXmlNode(ObjNota, "PrestadorServico/Endereco/Complemento")
                  sValue = GetValueXmlNode(ObjNota, "PrestadorServico/Endereco/Bairro")
                  sValue = GetValueXmlNode(ObjNota, "PrestadorServico/Endereco/CodigoMunicipio")
                  sValue = GetValueXmlNode(ObjNota, "PrestadorServico/Endereco/Uf")
                  sValue = GetValueXmlNode(ObjNota, "PrestadorServico/Endereco/Cep")
                  sValue = GetValueXmlNode(ObjNota, "PrestadorServico/Contato/Telefone")
                  sValue = GetValueXmlNode(ObjNota, "PrestadorServico/Contato/Email")
                  sValue = GetValueXmlNode(ObjNota, "TomadorServico/IdentificacaoTomador/CpfCnpj/Cpf")
                  sValue = GetValueXmlNode(ObjNota, "TomadorServico/RazaoSocial")
                  sValue = GetValueXmlNode(ObjNota, "TomadorServico/Endereco/Endereco")
                  sValue = GetValueXmlNode(ObjNota, "TomadorServico/Endereco/Numero")
                  sValue = GetValueXmlNode(ObjNota, "TomadorServico/Endereco/Bairro")
                  sValue = GetValueXmlNode(ObjNota, "TomadorServico/Endereco/CodigoMunicipio")
                  sValue = GetValueXmlNode(ObjNota, "TomadorServico/Endereco/Uf")
                  sValue = GetValueXmlNode(ObjNota, "TomadorServico/Endereco/Cep")
                  sValue = GetValueXmlNode(ObjNota, "TomadorServico/Contato/Telefone")
                  sValue = GetValueXmlNode(ObjNota, "TomadorServico/Contato/Email")
                  sValue = GetValueXmlNode(ObjNota, "OrgaoGerador/CodigoMunicipio")
                  sValue = GetValueXmlNode(ObjNota, "OrgaoGerador/Uf")
               End If
            Else
               If QtdNFs = 1 Then
                  GoTo Fim
               End If
            End If
         End If
         Set TbNota = Nothing
         Set TBRecibo = Nothing
         Set TbLote = Nothing
      Next
   End If
   On Error Resume Next
   If cQueries.Count = 0 Then
      If pMsg Then Call ExibirAviso("Não existe Nota a serem importadas.")
   Else
      If mvarSys.xDb.Executa(cQueries) Then
         If pMsg Then Call ExibirAviso("Processo realizado com Sucesso!")
      Else
         If pMsg Then Call ExibirAviso("Fim da importação")
      End If
   End If
   GoTo Fim
ShowError:
   Call ShowError
Fim:
   Screen.MousePointer = vbDefault
End Sub
Private Function DefineTPLogradouro(pEnd As String) As String
   Dim sEnd As String
   sEnd = Trim(Replace(pEnd, Chr(13) + Chr(10), " "))
   
   If Mid(UCase(sEnd), 1, 2) = "R " Or Mid(UCase(sEnd), 1, 2) = "R." Or Mid(UCase(sEnd), 1, 2) = "R:" Then
      DefineTPLogradouro = "RUA"                                         '12-Tipo do endereço
   ElseIf Mid(UCase(sEnd), 1, 4) = "RUA " Or Mid(UCase(sEnd), 1, 4) = "RUA." Or Mid(UCase(sEnd), 1, 4) = "RUA:" Then
      DefineTPLogradouro = "RUA"                                         '12-Tipo do endereço
   ElseIf Mid(UCase(sEnd), 1, 3) = "AV " Or Mid(UCase(sEnd), 1, 3) = "AV." Or Mid(UCase(sEnd), 1, 3) = "AV:" Then
      DefineTPLogradouro = "AVE"                                         '12-Tipo do endereço
   ElseIf Mid(UCase(sEnd), 1, 8) = "AVENIDA " Or Mid(UCase(sEnd), 1, 8) = "AVENIDA." Or Mid(UCase(sEnd), 1, 8) = "AVENIDA:" Then
      DefineTPLogradouro = "AVE"                                         '12-Tipo do endereço
   ElseIf Mid(UCase(sEnd), 1, 4) = "EST " Or Mid(UCase(sEnd), 1, 4) = "EST." Or Mid(UCase(sEnd), 1, 4) = "EST:" Then
      DefineTPLogradouro = "EST"                                         '12-Tipo do endereço
   ElseIf Mid(UCase(sEnd), 1, 8) = "ESTRADA " Or Mid(UCase(sEnd), 1, 8) = "ESTRADA." Or Mid(UCase(sEnd), 1, 8) = "ESTRADA:" Then
      DefineTPLogradouro = "EST"                                         '12-Tipo do endereço
   ElseIf Mid(UCase(sEnd), 1, 3) = "TV " Or Mid(UCase(sEnd), 1, 3) = "TV." Or Mid(UCase(sEnd), 1, 3) = "TV:" Then
      DefineTPLogradouro = "TRA"                                         '12-Tipo do endereço
   ElseIf Mid(UCase(sEnd), 1, 5) = "TRAV " Or Mid(UCase(sEnd), 1, 5) = "TRAV." Or Mid(UCase(sEnd), 1, 5) = "TRAV:" Then
      DefineTPLogradouro = "TRA"                                         '12-Tipo do endereço
   ElseIf Mid(UCase(sEnd), 1, 9) = "TRAVESSA " Or Mid(UCase(sEnd), 1, 9) = "TRAVESSA." Or Mid(UCase(sEnd), 1, 9) = "TRAVESSA:" Then
      DefineTPLogradouro = "TRA"                                         '12-Tipo do endereço
   Else
      DefineTPLogradouro = "RUA"                                         '12-Tipo do endereço
   End If

End Function
Private Function TrataEndereco(pEnd As String) As String
   Dim sEnd As String
   
   sEnd = Trim(Replace(pEnd, Chr(13) + Chr(10), " "))
   
   If Mid(UCase(sEnd), 1, 2) = "R " Or Mid(UCase(sEnd), 1, 2) = "R." Or Mid(UCase(sEnd), 1, 2) = "R:" Then
      TrataEndereco = Left(Trim(Mid(sEnd, 3)) & Space(125), 125)    '13-Endereço
      
   ElseIf Mid(UCase(sEnd), 1, 4) = "RUA " Or Mid(UCase(sEnd), 1, 4) = "RUA." Or Mid(UCase(sEnd), 1, 4) = "RUA:" Then
      TrataEndereco = Left(Trim(Mid(sEnd, 5)) & Space(125), 125)    '13-Endereço
      
   ElseIf Mid(UCase(sEnd), 1, 3) = "AV " Or Mid(UCase(sEnd), 1, 3) = "AV." Or Mid(UCase(sEnd), 1, 3) = "AV:" Then
      TrataEndereco = Left(Trim(Mid(sEnd, 4)) & Space(125), 125)    '13-Endereço
      
   ElseIf Mid(UCase(sEnd), 1, 8) = "AVENIDA " Or Mid(UCase(sEnd), 1, 8) = "AVENIDA." Or Mid(UCase(sEnd), 1, 8) = "AVENIDA:" Then
      TrataEndereco = Left(Trim(Mid(sEnd, 9)) & Space(125), 125)    '13-Endereço
      
   ElseIf Mid(UCase(sEnd), 1, 4) = "EST " Or Mid(UCase(sEnd), 1, 4) = "EST." Or Mid(UCase(sEnd), 1, 4) = "EST:" Then
      TrataEndereco = Left(Trim(Mid(sEnd, 5)) & Space(125), 125)    '13-Endereço
      
   ElseIf Mid(UCase(sEnd), 1, 8) = "ESTRADA " Or Mid(UCase(sEnd), 1, 8) = "ESTRADA." Or Mid(UCase(sEnd), 1, 8) = "ESTRADA:" Then
      TrataEndereco = Left(Trim(Mid(sEnd, 9)) & Space(125), 125)    '13-Endereço
      
   ElseIf Mid(UCase(sEnd), 1, 3) = "TV " Or Mid(UCase(sEnd), 1, 3) = "TV." Or Mid(UCase(sEnd), 1, 3) = "TV:" Then
      TrataEndereco = Left(Trim(Mid(sEnd, 4)) & Space(125), 125)    '13-Endereço
      
   ElseIf Mid(UCase(sEnd), 1, 5) = "TRAV " Or Mid(UCase(sEnd), 1, 5) = "TRAV." Or Mid(UCase(sEnd), 1, 5) = "TRAV:" Then
      TrataEndereco = Left(Trim(Mid(sEnd, 6)) & Space(125), 125)    '13-Endereço
      
   ElseIf Mid(UCase(sEnd), 1, 9) = "TRAVESSA " Or Mid(UCase(sEnd), 1, 9) = "TRAVESSA." Or Mid(UCase(sEnd), 1, 9) = "TRAVESSA:" Then
      TrataEndereco = Left(Trim(Mid(sEnd, 10)) & Space(125), 125)   '13-Endereço
      
   Else
      TrataEndereco = Left(sEnd & Space(125), 125)                  '13-Endereço
   End If
End Function

