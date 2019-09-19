VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_CadContato"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe     As frmCadCliente
Attribute mvarMe.VB_VarHelpID = -1

Private mvarSys As Object
Private mvarIDLOJA As Integer
Private mvarIDCLIENTE As Integer

Private mvarIsDirt As Boolean
Private clsCliente As Object
Dim bNmDirt As Boolean
Dim beMailVazio As Boolean
Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
   mvarIDLOJA = mvarSys.Propriedades("IDLOJA")
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Public Sub Show()
   mvarMe.Show vbModal
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New frmCadCliente
End Sub
Private Sub Class_Terminate()
   On Error Resume Next
   Unload mvarMe
   Set mvarMe = Nothing
End Sub
Public Property Get IDLOJA() As Integer
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDLOJA(ByVal vNewValue As Integer)
    mvarIDLOJA = vNewValue
End Property
Public Property Get IDCLIENTE() As Integer
   IDCLIENTE = mvarIDCLIENTE
   Debug.Print mvarIDCLIENTE
End Property
Public Property Let IDCLIENTE(ByVal vNewValue As Integer)
    mvarIDCLIENTE = vNewValue
    Debug.Print mvarIDCLIENTE
End Property
Public Property Get IsDirt() As Boolean
   IsDirt = mvarIsDirt
End Property
Private Sub PopulaTela()
   Dim Sql As String
   
   If clsCliente Is Nothing Then Call f_InstanciaClasseCliente
   With clsCliente
      If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDCLIENTE:=mvarIDCLIENTE) Then
         mvarMe.CmdExcluir.Enabled = True
         
         mvarMe.txtID.Text = .IDCLIENTE
         mvarMe.txtCelular.Text = .TEL1
         mvarMe.txtTel.Text = .TEL2
         mvarMe.txtNome.Text = .NOME
         mvarMe.LblDTCADASTRO.Caption = "Data de Cadastro: " & Format(.DTCADASTRO, "dd/mm/yy")
         Call LocalizarCombo(mvarMe.CmbSexo, .SEXO, True, False)
         mvarMe.TxtRegistro.MaskedText = UnFormat(.REGISTRO)
         mvarMe.txtFax.Text = .FAX
         mvarMe.TxtEmail.Text = .EMAIL
         mvarMe.txtEndereco.Text = .ENDERECO
         mvarMe.CmbBairro.Text = .Bairro
         mvarMe.CmbCidade.Text = .CIDADE
         mvarMe.cmbEstado.Text = .Estado
         mvarMe.cmbPais.Text = .PAIS
         mvarMe.txtCEP.Text = .CEP
         mvarMe.TxtDTNASC.Text = .DTNASC
         Call DefineIdade
         mvarMe.OptATIVO(.Ativo).Value = True
         mvarMe.TxtMotivoInat.Text = .MOTIVOINAT
         mvarMe.txtEmpresa.Text = .EMPRESA
         mvarMe.txtCargo.Text = .CARGO
         Call LocalizarCombo(mvarMe.CmbFototipo, .FOTOTIPO, True, True)
         Call LocalizarCombo(mvarMe.CmbProfissao, .PROFISSAO, True, False)
         Call LocalizarCombo(mvarMe.CmbConhecimento, .IDTPCONHEC, True, True)
         Call LocalizarCombo(mvarMe.CmbFunc, .IDFUNC, True, True)
         mvarMe.CmbDSCTPCONHEC.Text = .DSCTPCONHEC
         mvarMe.ChkFLGMARKETING.Value = IIf(.FLGMARKETING = 1, 0, 1)
         mvarMe.ChkFLGAGENDA.Value = IIf(.FLGAGENDA = 1, 0, 1)
         mvarMe.TxtObs.Text = .OBS
         mvarMe.ChkIsento.Value = .ISENTO
         mvarMe.ChkNFE.Value = .NFE
         
         Call SetTag(mvarMe.LblProx, "IDEVENTO", "")
         Sql = "Select IDEVENTO, StartDateTime"
         Sql = Sql & " From OEVENTOAGENDA "
         Sql = Sql + " Where IDLOJA= " & mvarIDLOJA
         Sql = Sql & " And StartDateTime>=GetDate()"
         Sql = Sql & " And IDCLIENTE=" & .IDCLIENTE
         Sql = Sql & " And FLGCANCELADO=0"
         Sql = Sql & " Order By StartDateTime"
         If mvarSys.xDb.AbreTabela(Sql) Then
            mvarMe.LblProx.Caption = Format(mvarSys.xDb.RsAux("StartDateTime"), "dddd, dd/mm hh:mm") & "h"
            mvarMe.CmdProxSessao.Caption = mvarMe.LblProx.Caption
            Call SetTag(mvarMe.LblProx, "IDEVENTO", mvarSys.xDb.RsAux("IDEVENTO"))
         End If
         
         Call PopulaGrdCupom
         'Call PopulaGrdSessao(1)
         Call PopulaGrdSessao(mvarMe.TabSessao.Selected.Index)
         'Call PopulaGrdVenda
         Call PopulaTrv
         Call ExibeExpHis
      Else
         Call LimpaTela
      End If
   End With
End Sub
Private Sub ExibeExpHis()
   Dim Sql As String
   
   mvarMe.CmdExpHis.Visible = False
   mvarMe.CmdExpHis.Tag = ""
   
   Sql = "Select * "
   Sql = Sql & " From OCLIENTE "
   Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
   Sql = Sql & " And IDCLIENTE<>" & SqlNum(mvarMe.txtID.Text)
   Sql = Sql & " And NOME = " & SqlStr(mvarMe.txtNome.Text)
   If mvarSys.xDb.AbreTabela(Sql) Then
      If mvarSys.xDb.RsAux.recordcount > 0 Then
         mvarMe.CmdExpHis.Visible = True
         mvarMe.CmdExpHis.Tag = mvarSys.xDb.RsAux("IDCLIENTE")
      End If
   End If
End Sub
Private Sub PopulaGrdCupom()
   Dim Sql  As String
   Dim i    As Long
   Dim MyRs As Object
   
   With mvarMe.GrdCupom
      .Clear True
      .RowMode = True
      .Editable = False
      .UseXPStyles = True
      
      Sql = ""
      Sql = Sql + " Select P.DSCPROMO [Promoção], C.CODCUPOM [Cupom], CV.IDVENDA [Venda], P.IDPROMO, C.IDCUPOM"
      Sql = Sql + " From CPROMOCAO P"
      Sql = Sql + " Join CCUPOM C         On  P.IDLOJA= C.IDLOJA AND P.IDPROMO=C.IDPROMO"
      Sql = Sql + " Join CCUPOM_VENDA CV  On CV.IDLOJA= C.IDLOJA AND CV.IDCUPOM=C.IDCUPOM"
      Sql = Sql + " Join CVENDA V         On  V.IDLOJA=CV.IDLOJA And V.IDVENDA=CV.IDVENDA"
      Sql = Sql + " Where V.IDLOJA= " & mvarIDLOJA
      Sql = Sql + " And V.IDCLIENTE=" & SqlNum(mvarIDCLIENTE)
      
      Call mvarSys.xDb.AbreTabela(Sql, MyRs)
      Call .FillFromRS(MyRs)
      .ColVisible("IDPROMO") = False
      .ColVisible("IDCUPOM") = False
      .ColWidth("Promoção") = 120
      .ColWidth("Cupom") = 80
      .ColWidth("Venda") = 80
      '.Width = 5880
   End With
End Sub
Private Sub PopulaGrdSessao(Optional Tipo As Integer = 0)
   Dim Sql  As String
   Dim i    As Long
   Dim MyRs As Object
   Dim bOk  As Boolean
   Dim nValor As Currency
   
   mvarMe.LblVALOR.Caption = "R$ 0,00"
   
   With mvarMe.GrdSessao
      .Clear True
      .RowMode = True
      .Editable = False
      .UseXPStyles = True
      Select Case Tipo
         Case 0
            Sql = "Select Sum(V.VLPGTO)-Sum(V.VLTROCO) [TOTAL]"
            Sql = Sql + " From CVENDA V"
            Sql = Sql + " Where V.IDLOJA= " & mvarIDLOJA
            Sql = Sql + " And V.IDCLIENTE = " & mvarIDCLIENTE
            mvarMe.LblVALOR.Caption = ValBr(0)
            If mvarSys.xDb.AbreTabela(Sql, MyRs) Then
               mvarMe.LblVALOR.Caption = "R$ " & ValBr(MyRs("TOTAL") & "")
            End If
            mvarMe.LblTotal1.Left = mvarMe.LblVALOR.Left - mvarMe.LblTotal1.Width - 60
            
            Sql = GetSqlSessoes(True)
            bOk = mvarSys.xDb.AbreTabela(Sql, MyRs)
            If Not bOk Then
               Sql = GetSqlSessoes(False)
               bOk = mvarSys.xDb.AbreTabela(Sql, MyRs)
            End If
            If bOk Then
               .Redraw = False
               .Visible = False
               Call .FillFromRS(MyRs)
               .ColVisible("IDCLIENTE") = False
               .ColWidth("#") = 20
               .ColWidth("Data") = 80
               .ColHeaderTextFlags("Data") = igTextCenter + igTextVCenter
               .ColSortType("Data") = igSortByValue
               .ColWidth("Item") = 110
               .ColWidth("Área") = 90
               If mvarSys.Propriedades("TPAREA") = "0" Then
                  .ColWidth("Item") = 200
                  .ColVisible("Área") = False
               Else
                  .ColVisible("Área") = True
               End If
               .ColWidth("Pot.") = 32
               .ColWidth("Disp.") = 38
               .ColWidth("Freq.") = 38
               .ColWidth("Tempo") = 45
               .ColWidth("Operadora") = 61
               .ColVisible("DTVALUE") = False
               .ColVisible("IDATENDIMENTO") = False
               .ColVisible("IDVENDA") = False
               '.Width = 5880
               
               For i = 1 To .RowCount
                  .CellForeColor(i, "Freq.") = vbBlack 'IIf(.CellValue(i, "Freq.") > 39, vbRed, vbBlack)
                  .CellBackColor(i, "Freq.") = IIf(.CellValue(i, "Freq.") > 39, vbRed, vbWhite)
               Next
               Call AjustaGrpSessao(GetTag(mvarMe.CmdExtende, "ABERTO", "1"))
               .Visible = True
               .Redraw = True
            End If
        
         Case 1
            Sql = " Select Row_Number() Over (Order By DTVALUE) [#], *"
            Sql = Sql & " From ("
            Sql = Sql & " Select E.IDCLIENTE "
            'Sql = Sql & ", Cast(Convert(Char(10), E.StartDateTime, 103) As SmallDatetime) [Data]"
            Sql = Sql & ", Cast(Convert(VARCHAR, E.StartDateTime, 103)+ ' ' +Convert(VARCHAR(5), E.StartDateTime, 108) as Datetime)  [Data]"
            Sql = Sql & " , S.DSCSERVICO+ ' ' + ' : ' + T.DSCTRATAMENTO "
            Sql = Sql & " + ' : ' + Cast(Count(A.DSCAREA) as Varchar)"
            Sql = Sql & " + ' Área' + Case Count(A.DSCAREA) When 1 Then '' Else 's' End[Item]"
            Sql = Sql & " , (365*Year(E.StartDateTime))+(32*Month(E.StartDateTime))+Day(E.StartDateTime) as [DTVALUE]"
            Sql = Sql & " , E.IDLOJA, E.IDEVENTO, E.FLGCANCELADO"
            Sql = Sql & " From (OEVENTOAGENDA E"
            Sql = Sql & " Join OSERVICOEVT SE   On SE.IDLOJA=E.IDLOJA And SE.IDEVENTO=E.IDEVENTO"
            Sql = Sql & " Join OTPSERVICO S     On SE.IDLOJA=S.IDLOJA AND SE.IDTPSERVICO=S.IDTPSERVICO"
            Sql = Sql & " Join OTPTRATAMENTO T  On SE.IDLOJA=T.IDLOJA AND SE.IDTPTRATAMENTO=T.IDTPTRATAMENTO"
            Sql = Sql & " Join OAREA A          On SE.IDLOJA=A.IDLOJA And SE.IDAREA=A.IDAREA)"
            Sql = Sql & " Where E.IDLOJA = " & mvarIDLOJA
            Sql = Sql & " And E.IDCLIENTE=" & mvarIDCLIENTE
            'Sql = Sql & " And E.FLGCANCELADO=1"
            Sql = Sql & " Group By E.IDLOJA, E.IDCLIENTE, E.StartDateTime, S.DSCSERVICO, T.DSCTRATAMENTO"
            Sql = Sql & ", E.IDEVENTO, E.FLGCANCELADO"
            Sql = Sql & " ) Temp"
         
            If mvarSys.xDb.AbreTabela(Sql, MyRs) Then
               .Redraw = False
               .Visible = False
               '&H00808080&
               Call .FillFromRS(MyRs)
               .ColVisible("IDCLIENTE") = False
               .ColWidth("#") = 20
               .ColWidth("Data") = 120
               '.ColHeaderTextFlags("Data") = igTextCenter
               .ColWidth("Item") = 260
               .ColVisible("DTVALUE") = False
               .ColVisible("IDLOJA") = False
               .ColVisible("IDEVENTO") = False
               .ColVisible("FLGCANCELADO") = False
               .Visible = True
               .Redraw = True
               For i = 1 To .RowCount
                  If .CellValue(i, "FLGCANCELADO") = 1 Then
                     .CellForeColor(i, "#") = vbRed
                     .CellForeColor(i, "Data") = vbRed
                     .CellForeColor(i, "Item") = vbRed
                  ElseIf CDate(.CellValue(i, "DATA")) <= Now Then
                     .CellForeColor(i, "#") = &H808080
                     .CellForeColor(i, "Data") = &H808080
                     .CellForeColor(i, "Item") = &H808080
                  Else
                     .CellForeColor(i, "#") = vbBlack
                     .CellForeColor(i, "Data") = vbBlack
                     .CellForeColor(i, "Item") = vbBlack
                  End If
               Next
               
               mvarMe.TabSessao.Item(Tipo).Enabled = True
            Else
               mvarMe.TabSessao.Item(Tipo).Enabled = False
            End If
         
         Case 2
            .ImageList = mvarMe.ImgFatura
            
            Sql = "Select '' [#], T.SITTAREFA, Right(REPLICATE('0',6)+ Cast(T.IDTAREFA as varchar),6) [Tarefa]"
            Sql = Sql & ", T.FLGDELETE, Cast(Convert(Char(10),T.DTTAREFA, 103) As SmallDatetime)[Data], C.DSCTPTAREFA [Tipo]"
            Sql = Sql & ", T.IDEVENTO, T.IDTAREFA, T.IDLOJA"
            Sql = Sql & " From OTAREFAEVT T "
            Sql = Sql & " Join OTIPOTAREFA C On T.IDLOJA=C.IDLOJA AND T.IDTPTAREFA=C.IDTPTAREFA"
            Sql = Sql & " Where T.IDLOJA= " & mvarIDLOJA
            Sql = Sql & " And T.IDCLIENTE = " & mvarIDCLIENTE
            Sql = Sql & " Order By Data"
            If mvarSys.xDb.AbreTabela(Sql) Then
               .FillFromRS mvarSys.xDb.RsAux
               '.Redraw = True
               
               .ColVisible("SITTAREFA") = False
               .ColVisible("FLGDELETE") = False
               .ColVisible("IDEVENTO") = False
               .ColVisible("IDTAREFA") = False
               .ColVisible("IDLOJA") = False
               For i = 1 To mvarMe.GrdSessao.RowCount
                  If .CellValue(i, "FLGDELETE") = 1 Then
                     .RowVisible(i) = False
                     .CellIcon(i, "#") = 4
                     .CellFont(i, "Tarefa").Strikethrough = True
                     .CellFont(i, "Data").Strikethrough = True
                     .CellFont(i, "Tipo").Strikethrough = True
                     .CellForeColor(i, "Tarefa") = RGB(128, 128, 128)
                     .CellForeColor(i, "Data") = RGB(128, 128, 128)
                     .CellForeColor(i, "Tipo") = RGB(128, 128, 128)
                  
                  Else
                     If .CellValue(i, "SITTAREFA") = 10 Then
                        .CellIcon(i, "#") = 2
                        
                     ElseIf .CellValue(i, "SITTAREFA") = 5 Then
                        .CellIcon(i, "#") = 3
                     End If
                  End If
                  Call .AutoHeightRow(i)
                  .CellTextFlags(i, "Tarefa") = .CellTextFlags(i, "Tarefa") + igTextVCenter
                  .CellTextFlags(i, "Data") = .CellTextFlags(i, "Data") + igTextVCenter
                  .CellTextFlags(i, "Tipo") = .CellTextFlags(i, "Tipo") + igTextVCenter
               Next
               
               For i = 1 To mvarMe.GrdSessao.ColCount
                  mvarMe.GrdSessao.AutoWidthCol (i)
                  mvarMe.GrdSessao.ColWidth(i) = mvarMe.GrdSessao.ColWidth(i) + 10
               Next
               
            End If
         Case 3 'Vendas
            Sql = "Select Distinct Right(REPLICATE('0',6)+Rtrim(Cast(V.IDVENDA As Char)),6) [Venda]" & vbNewLine
            Sql = Sql & ", Cast(Convert(Char(10),V.DTVENDA, 103) As SmallDatetime) [Data]" & vbNewLine
            Sql = Sql & ", C.NOME [Nome]" & vbNewLine
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
            Sql = Sql & " Join OCLIENTE C On C.IDLOJA=V.IDLOJA And V.IDCLIENTE=C.IDCLIENTE" & vbNewLine
            Sql = Sql & " Left Join FRECIBO R On V.IDLOJA=R.IDLOJA And V.IDVENDA=R.IDVENDA" & vbNewLine
            Sql = Sql & " Left Join FNOTAFISCAL N On V.IDLOJA=N.IDLOJA And V.IDVENDA=N.IDVENDA And N.FLGCANCELADA=0" & vbNewLine
            Sql = Sql & " Where V.IDLOJA=" & SqlNum(mvarIDLOJA) & vbNewLine
            Sql = Sql & " And V.IDCLIENTE=" & mvarIDCLIENTE & vbNewLine
            Sql = Sql & " Union" & vbNewLine
            Sql = Sql & " Select Distinct Right(REPLICATE('0',6)+ Cast(V.IDVENDA as varchar),6) [Venda]" & vbNewLine
            Sql = Sql & ", Cast(Convert(Char(10), V.DTVENDA, 103) As SmallDatetime) [Data]" & vbNewLine
            Sql = Sql & ", C.NOME [Nome]" & vbNewLine
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
            Sql = Sql & " Join OCLIENTE C on A.IDLOJA=C.IDLOJA and A.IDCLIENTE=C.IDCLIENTE" & vbNewLine
            Sql = Sql & " Left Join FRECIBO R On V.IDLOJA=R.IDLOJA And V.IDVENDA=R.IDVENDA" & vbNewLine
            Sql = Sql & " Left Join FNOTAFISCAL N On V.IDLOJA=N.IDLOJA And V.IDVENDA=N.IDVENDA And N.FLGCANCELADA=0" & vbNewLine
            Sql = Sql & " Where A.IDLOJA =" & mvarIDLOJA & vbNewLine
            Sql = Sql & " And A.IDCLIENTE=" & mvarIDCLIENTE & vbNewLine
            Sql = Sql & " Order By Data, Venda"
            
            If mvarSys.xDb.AbreTabela(Sql) Then
               mvarMe.GrdSessao.FillFromRS mvarSys.xDb.RsAux
               
               nValor = 0
               'mvarMe.GrdSessao.ColVisible("IDVENDA") = False
               For i = 1 To mvarMe.GrdSessao.RowCount
                  nValor = nValor + xVal(mvarMe.GrdSessao.CellValue(i, "Pgto."))
                  mvarMe.GrdSessao.CellValue(i, "Valor") = "R$ " & ValBr(mvarMe.GrdSessao.CellValue(i, "Valor"))
                  mvarMe.GrdSessao.CellValue(i, "Desc.") = "R$ " & ValBr(mvarMe.GrdSessao.CellValue(i, "Desc."))
                  mvarMe.GrdSessao.CellValue(i, "Pgto.") = "R$ " & ValBr(mvarMe.GrdSessao.CellValue(i, "Pgto."))
                  'mvarMe.GrdSessao.CellValue(i, "Data") = Format(mvarMe.GrdSessao.CellValue(i, "Data"), "dd/mm/yy")
               Next
               mvarMe.LblVALOR.Caption = "R$ " & ValBr(nValor)
               mvarMe.LblTotal1.Left = mvarMe.LblVALOR.Left - mvarMe.LblTotal1.Width - 60
               For i = 1 To mvarMe.GrdSessao.ColCount
                  mvarMe.GrdSessao.AutoWidthCol (i)
                  mvarMe.GrdSessao.ColWidth(i) = mvarMe.GrdSessao.ColWidth(i) + 5
               Next
               mvarMe.GrdSessao.ColVisible("Recibo") = False
               mvarMe.GrdSessao.ColVisible("N.F.") = False
               mvarMe.GrdSessao.ColWidth("Venda") = 45
               mvarMe.GrdSessao.ColWidth("Data") = 65
               mvarMe.GrdSessao.ColWidth("Nome") = 60
               mvarMe.GrdSessao.ColWidth("Sld") = 30
            End If
         Case 4 'Faturas
            Sql = "Select '' [#], F.SITFATURA, F.IDLOJA, Right(REPLICATE('0',6)+ Cast(F.IDFATURA as varchar),6) [Fatura]" & vbNewLine
            Sql = Sql & ", Cast(Convert(Char(10), F.DTEMISSAO, 103) As SmallDatetime) [Data]" & vbNewLine
            Sql = Sql & ", F.IDATENDIMENTO [Atendimento], F.HISTORICO [Historico], F.VALOR [Valor]" & vbNewLine
            Sql = Sql & " From FFATURA F " & vbNewLine
            Sql = Sql & " Where F.IDLOJA= " & mvarIDLOJA & vbNewLine
            Sql = Sql & " And F.IDCLIENTE = " & mvarIDCLIENTE & vbNewLine
            Sql = Sql & " Order By F.DTEMISSAO" & vbNewLine
            If mvarSys.xDb.AbreTabela(Sql) Then
               mvarMe.GrdSessao.FillFromRS mvarSys.xDb.RsAux
               nValor = 0
               'mvarMe.GrdSessao.ColVisible("IDVENDA") = False
               For i = 1 To mvarMe.GrdSessao.RowCount
                  nValor = nValor + xVal(mvarMe.GrdSessao.CellValue(i, "Valor"))
                  mvarMe.GrdSessao.CellValue(i, "Valor") = "R$ " & ValBr(mvarMe.GrdSessao.CellValue(i, "Valor"))
                  If mvarMe.GrdSessao.CellValue(i, "Atendimento") = 0 Then
                     mvarMe.GrdSessao.CellValue(i, "Atendimento") = ""
                  End If
                  If .CellValue(i, "SITFATURA") = 10 Then
                     .CellIcon(i, "#") = 3
                  Else
                     .CellIcon(i, "#") = 1
                  End If
                  Call .AutoHeightRow(i)
               Next
               mvarMe.LblVALOR.Caption = "R$ " & ValBr(nValor)
               mvarMe.LblTotal1.Left = mvarMe.LblVALOR.Left - mvarMe.LblTotal1.Width - 60
               
               mvarMe.GrdSessao.ColVisible("SITFATURA") = False
               mvarMe.GrdSessao.ColVisible("IDLOJA") = False
               For i = 1 To mvarMe.GrdSessao.ColCount
                  mvarMe.GrdSessao.AutoWidthCol (i)
                  mvarMe.GrdSessao.ColWidth(i) = mvarMe.GrdSessao.ColWidth(i) + 5
               Next
               mvarMe.GrdSessao.ColWidth("Historico") = 120
               
            End If
      
      End Select
   End With
End Sub
Private Sub PopulaGrdVenda()
   Dim Sql  As String
   Dim i    As Long
   Dim MyRs As Object
   
   With mvarMe.GrdCupom
      .Clear True
      .RowMode = True
      .Editable = False
      .UseXPStyles = True
      
      Sql = ""
      Sql = Sql + " Select P.DSCPROMO [Promoção], C.CODCUPOM [Cupom], CV.IDVENDA [Venda], P.IDPROMO, C.IDCUPOM"
      Sql = Sql + " From CPROMOCAO P"
      Sql = Sql + " Join CCUPOM C         On  P.IDLOJA= C.IDLOJA AND P.IDPROMO=C.IDPROMO"
      Sql = Sql + " Join CCUPOM_VENDA CV  On CV.IDLOJA= C.IDLOJA AND CV.IDCUPOM=C.IDCUPOM"
      Sql = Sql + " Join CVENDA V         On  V.IDLOJA=CV.IDLOJA And V.IDVENDA=CV.IDVENDA"
      Sql = Sql + " Where V.IDLOJA= " & mvarIDLOJA
      Sql = Sql + " And V.IDCLIENTE=" & SqlNum(mvarIDCLIENTE)
      
      Call mvarSys.xDb.AbreTabela(Sql, MyRs)
      Call .FillFromRS(MyRs)
      .ColVisible("IDPROMO") = False
      .ColVisible("IDCUPOM") = False
      .ColWidth("Promoção") = 120
      .ColWidth("Cupom") = 80
      .ColWidth("Venda") = 80
      '.Width = 5880
   End With
End Sub

Private Function GetSqlSessoes(bFreq As Boolean) As String
   Dim Sql As String
   
   Sql = ""
   Sql = Sql + " Select Row_Number() Over (Order By DTVALUE) [#], * From ("
   Sql = Sql + " Select AT.IDCLIENTE"
   Sql = Sql & ", Cast(Convert(Char(10),AT.DTATEND, 103) As SmallDatetime) [Data]"
   Sql = Sql + ", Case S.DSCSERVICO When 'Teste' Then ' Teste' Else S.DSCSERVICO End + ' ' + Right('00'+CAST(SS.NUMSESSAO As Varchar),2) + ' : ' + T.DSCTRATAMENTO [Item]"
   Sql = Sql + ", A.DSCAREA As [Área], SS.POTENCIA [Pot.]"
   Sql = Sql + ", SS.DISPAROS [Disp.]"
   If bFreq Then
      Sql = Sql + " , DateDiff(d, (Select AT0.DTATEND"
      Sql = Sql + "    From OATENDIMENTO AT0"
      Sql = Sql + "    JOIN OSESSAO SS0 On SS0.IDLOJA=AT0.IDLOJA And SS0.IDATENDIMENTO=AT0.IDATENDIMENTO"
      Sql = Sql + "    Where AT0.IDCLIENTE = AT.IDCLIENTE"
      Sql = Sql + "    And SS0.IDLOJA=SS.IDLOJA"
      Sql = Sql + "    And SS0.IDAREA=SS.IDAREA"
      Sql = Sql + "    And SS0.IDTPSERVICO=SS.IDTPSERVICO"
      Sql = Sql + "    And SS0.IDTPTRATAMENTO=SS.IDTPTRATAMENTO"
      Sql = Sql + "    And SS0.NUMSESSAO=SS.NUMSESSAO-1), AT.DTATEND) [Freq.]"
   Else
      Sql = Sql + " , '' [Freq.]"
   End If
   Sql = Sql + " , Right('00'+ Cast(DATEDIFF(mi,AT.HHINI, AT.HHFIM)/60 As varchar),2)+':'"
   Sql = Sql + " + Right('00'+ Cast(DATEDIFF(mi,AT.HHINI, AT.HHFIM)%60 As varchar),2) [Tempo]"
   Sql = Sql + " , F.NOME [Operadora]"
   Sql = Sql + " , AT.DTATEND [DTVALUE]"
   Sql = Sql + " , AT.IDATENDIMENTO, Null [IDVENDA]"
   Sql = Sql + " From (OSESSAO SS"
   Sql = Sql + " Join OATENDIMENTO AT  On SS.IDLOJA=AT.IDLOJA And SS.IDATENDIMENTO=AT.IDATENDIMENTO"
   Sql = Sql + " Join RFUNCIONARIO F   On AT.IDLOJA=F.IDLOJA AND AT.IDFUNCIONARIO=F.IDFUNCIONARIO"
   Sql = Sql + " Join OTPSERVICO S     On SS.IDLOJA=S.IDLOJA AND SS.IDTPSERVICO=S.IDTPSERVICO"
   Sql = Sql + " Join OTPTRATAMENTO T  On SS.IDLOJA=T.IDLOJA AND SS.IDTPTRATAMENTO=T.IDTPTRATAMENTO"
   Sql = Sql + " Join OAREA A          On SS.IDLOJA=A.IDLOJA AND SS.IDAREA=A.IDAREA)"
   Sql = Sql + " Left Join OATENDIMENTO_PRODUTO AP On AT.IDLOJA=AP.IDLOJA And AT.IDATENDIMENTO=AP.IDATENDIMENTO"
   Sql = Sql + " Where AT.IDLOJA = " & mvarIDLOJA
   Sql = Sql + " And AT.IDCLIENTE =" & mvarIDCLIENTE
   
   Sql = Sql + " Union All"
   Sql = Sql + " Select V.IDCLIENTE , Convert(VARCHAR, V.DTVENDA, 103 ) [Data]"
   Sql = Sql + " , ' '+PR.NMPROD+ ' ( '+ CAST(CAST(I.QTDVENDA AS INT) AS VARCHAR)+' )' [Item]"
   Sql = Sql + " , Null As [Área], Null [Pot.]"
   Sql = Sql + " , Null [Disp.]"
   Sql = Sql + " , Null [Freq.]"
   Sql = Sql + " , Null [Tempo]"
   Sql = Sql + " , Null [Operadora]"
   Sql = Sql + " , V.DTVENDA as [DTVALUE]"
   Sql = Sql + " , Null [IDATENDIMENTO], V.IDVENDA"
   Sql = Sql + " From CVENDA V"
   Sql = Sql + " Join CITENSVENDA I On V.IDLOJA=I.IDLOJA And V.IDVENDA=I.IDVENDA"
   Sql = Sql + " Join SPRODUTO PR   On I.IDLOJA=PR.IDLOJA AND I.IDPROD=PR.IDPROD And PR.ESERVICO=0"
   'Sql = Sql + " Left Join OATENDIMENTO_VENDA AV On V.IDLOJA=AV.IDLOJA And V.IDVENDA=AV.IDVENDA"
   Sql = Sql + " Where V.IDLOJA = " & mvarIDLOJA
   Sql = Sql + " And V.IDCLIENTE =" & mvarIDCLIENTE
   Sql = Sql + " ) Temp"
   
   GetSqlSessoes = Sql
End Function
Private Sub PopulaTrv()
   'Dim TbClasseCli As Object
   
   Dim iTab As Integer
   Dim xNode As TreeViewNode
   Dim Sql As String
   Dim MyRs As Object
   Dim MyRsClasse As Object
   
   If Not mvarMe.Visible Then Exit Sub
   'DoEvents
   iTab = mvarMe.TabContato.SelectedItem
   mvarMe.TabContato.Item(2).Selected = True
   Sql = "Select * "
   Sql = Sql & " From OCLASSE_CLIENTE"
   Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
   Sql = Sql & " And IDCLIENTE=" & mvarIDCLIENTE
   If mvarSys.xDb.AbreTabela(Sql, MyRs) Then
      While Not MyRs.EOF
         If Not ExisteNo(mvarMe.TrvClasse, "k" & MyRs("IDCLASSE")) Then
            Sql = "Select * "
            Sql = Sql & " From OCLASSE"
            Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
            Sql = Sql & " And IDCLASSE=" & MyRs("IDCLASSE")
            If mvarSys.xDb.AbreTabela(Sql, MyRsClasse) Then
               While Not MyRsClasse.EOF
                  If MyRsClasse("IDPAI") = 0 Then
                     Set xNode = mvarMe.TrvClasse.Nodes.Add(, , "k" & MyRsClasse("IDCLASSE"), MyRsClasse("DSCCLASSE"))
                  Else
                     Set xNode = mvarMe.TrvClasse.Nodes.Add("k" & MyRsClasse("IDPAI"), tvwChild, "k" & MyRsClasse("IDCLASSE"), MyRsClasse("DSCCLASSE"))
                     mvarMe.TrvClasse.Nodes("k" & MyRsClasse("IDPAI")).Bold = True
                  End If
                  xNode.Expanded = True
                  xNode.Checked = True
                  Call SetTag(xNode, "ISDIRT", "0")
                  Call SetTag(xNode, "VALUE", "0")
                  MyRsClasse.MoveNext
               Wend
            End If
         End If
         mvarMe.Refresh
         DoEvents
         With mvarMe.TrvClasse.Nodes("k" & CStr(MyRs("IDCLASSE")))
            .Tag = SetTag(.Tag, "VALUE", "1")
            .Checked = True
         End With
         MyRs.MoveNext
      Wend
   End If
   mvarMe.TabContato.Item(iTab).Selected = True
End Sub
Private Sub LimpaTela()
   Dim n As Variant
   On Error Resume Next
   With mvarMe
      For Each n In .Controls
         n.Text = ""
         n.MaskedText = ""
         n.ListIndex = -1
      Next
   
      mvarMe.LblMenor.Visible = False
      mvarMe.LblMenor2.Visible = False
      
'      .txtID.Text = ""
'      .txtTel.Text = ""
'      .txtCelular.Text = ""
'      .txtNome.Text = ""
'      .TxtRegistro.MaskedText = ""
'      .txtFax.Text = ""
'      .txtEmail.Text = ""
'      .txtEndereco.Text = ""
'      .CmbBairro.Text = ""
'      .CmbCidade.Text = ""
'      .cmbEstado.Text = ""
'      .cmbPais.Text = ""
'      .txtCEP.Text = ""
'      .TxtDTNASC.Text = ""
'      .CmbSexo.ListIndex = -1
'      .txtEmpresa.Text = ""
'      .txtCargo.Text = ""
'      .TxtDSCTPCONHEC.Text = ""
'      .TxtObs.Text = ""
'      .TxtMotivoInat.Text = ""
      
      .CmbConhecimento.ListIndex = 0
      .CmbFunc.ListIndex = 0
      If .CmbFototipo.ListCount = 0 Then
         .CmbFototipo.ListIndex = -1
      Else
         .CmbFototipo.ListIndex = 0
      End If
      If .CmbConhecimento.ListCount = 0 Then
         .CmbConhecimento.ListIndex = -1
      Else
         .CmbConhecimento.ListIndex = 0
      End If
      Dim xNode As TreeViewNode
      For Each xNode In .TrvClasse.Nodes
         xNode.Checked = False
      Next
      .ChkClasse.Value = xtpUnchecked
      .OptATIVO(1).Value = True

      
      .GrdSessao.Clear False
      .LblVALOR.Caption = "R$ 0,00"
      .LblTotal1.Left = .LblVALOR.Left - .LblTotal1.Width - 60
   
      .CmdExcluir.Enabled = False
      .LblProx.Caption = ""
      .CmdProxSessao.Caption = .LblProx.Caption
      Call SetTag(.LblProx, "IDEVENTO", "")
      
      .GrdCupom.Clear False
      .ChkIsento.Value = xtpUnchecked
      .ChkNFE.Value = xtpUnchecked
   End With
End Sub
Private Sub PopulaClasse()
   
   If clsCliente Is Nothing Then Call f_InstanciaClasseCliente
   With clsCliente
      If .IDCLIENTE = 0 Then .DTCADASTRO = Date
      If .IDLOJA = 0 Then .IDLOJA = mvarIDLOJA
      .IDCOLIGADA = mvarSys.Propriedades("IDCOLIGADA")
      .TEL1 = mvarMe.txtCelular.Text
      .TEL2 = mvarMe.txtTel.Text
      bNmDirt = (.NOME <> mvarMe.txtNome.Text)
      .NOME = mvarMe.txtNome.Text
      .REGISTRO = mvarMe.TxtRegistro.MaskedText
      .FAX = mvarMe.txtFax.Text
      .EMAIL = mvarMe.TxtEmail.Text
      .ENDERECO = mvarMe.txtEndereco.Text
      .Bairro = FormatarNome(LCase(Trim(mvarMe.CmbBairro.Text)))
      .CIDADE = FormatarNome(LCase(Trim(mvarMe.CmbCidade.Text)))
      .Estado = UCase(mvarMe.cmbEstado.Text)
      .PAIS = FormatarNome(LCase(Trim(mvarMe.cmbPais.Text)))
      .CEP = mvarMe.txtCEP.Text
      .DTNASC = mvarMe.TxtDTNASC.Text
      .SEXO = Mid(mvarMe.CmbSexo.Text, 1, 1)
      .EMPRESA = mvarMe.txtEmpresa.Text
      .CARGO = mvarMe.txtCargo.Text
      .FOTOTIPO = mvarMe.CmbFototipo.ItemData(mvarMe.CmbFototipo.ListIndex)
      .PROFISSAO = FormatarNome(LCase(Trim(mvarMe.CmbProfissao.Text)))
      .IDTPCONHEC = mvarMe.CmbConhecimento.ItemData(mvarMe.CmbConhecimento.ListIndex)
      .IDFUNC = mvarMe.CmbFunc.ItemData(mvarMe.CmbFunc.ListIndex)
      .DSCTPCONHEC = mvarMe.CmbDSCTPCONHEC.Text
      .FLGMARKETING = IIf(mvarMe.ChkFLGMARKETING.Value = xtpChecked, 0, 1)
      .FLGAGENDA = IIf(mvarMe.ChkFLGAGENDA.Value = xtpChecked, 0, 1)
      .OBS = mvarMe.TxtObs.Text
      .Ativo = IIf(mvarMe.OptATIVO(0).Value, 0, IIf(mvarMe.OptATIVO(1).Value, 1, 2))
      .MOTIVOINAT = mvarMe.TxtMotivoInat.Text
      .ISENTO = mvarMe.ChkIsento.Value
      .NFE = mvarMe.ChkNFE.Value
   End With
End Sub
Private Sub mvarMe_Activate()
'   mvarMe.TabContato.Item(2).Selected = True
'   Call MontaTela
'   Call MontaTrv
'Call AjustaGrpSessao(GetTag(mvarMe.CmdExtende, "ABERTO", "1"))
   If GetTag(mvarMe, "1VEZ", 1) = 1 Then
      Call SetTag(mvarMe, "1VEZ", 0)
'      Call PopulaTela
      Call PopulaTrv
   End If
   mvarMe.MousePointer = vbDefault
   Screen.MousePointer = vbDefault
End Sub
Private Sub MontaDscConhec()
   Dim Sql As String
   Dim MyRs As Object
   Dim sAux As String
   
   If xVal(mvarMe.CmbConhecimento.ItemData(mvarMe.CmbConhecimento.ListIndex)) <= 0 Then
      Exit Sub
   End If
   
   sAux = mvarMe.CmbDSCTPCONHEC.Text
   
   
   Sql = "Select Distinct DSCTPCONHEC"
   Sql = Sql & " From OCLIENTE"
   Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
   Sql = Sql & " And IDTPCONHEC=" & xVal(mvarMe.CmbConhecimento.ItemData(mvarMe.CmbConhecimento.ListIndex))
   Sql = Sql & " Order By DSCTPCONHEC"
   mvarMe.CmbDSCTPCONHEC.Clear
   If mvarSys.xDb.AbreTabela(Sql, MyRs) Then
      While Not MyRs.EOF
         mvarMe.CmbDSCTPCONHEC.AddItem (MyRs("DSCTPCONHEC") & "")
         MyRs.MoveNext
      Wend
      mvarMe.CmbDSCTPCONHEC.ListIndex = 0
   End If
   mvarMe.CmbDSCTPCONHEC.Text = sAux
End Sub
Private Sub mvarMe_CmbConhecimentoClick()
   Call MontaDscConhec
End Sub

Private Sub mvarMe_CmdExpHisClick()
   Dim nID As Integer
   Dim Sql As String
   Dim cQueries As Collection
   
   nID = xVal(InputBox("Favor Informar o identificador do Cliente Destino." & vbNewLine & vbNewLine & "Cliente: '" & Trim(mvarMe.txtNome.Text) & "'", "Exportar Histórico", xVal(mvarMe.CmdExpHis.Tag)))
   
   With mvarSys.xDb
      Sql = "Select * "
      Sql = Sql & " From OCLIENTE "
      Sql = Sql & " Where NOME = " & SqlStr(mvarMe.txtNome.Text)
      Sql = Sql & " And IDCLIENTE<>" & SqlNum(mvarMe.txtID.Text)
      Sql = Sql & " And IDLOJA=" & mvarIDLOJA
      Sql = Sql & " And IDCLIENTE=" & SqlNum(nID)
      If .AbreTabela(Sql) Then
         Set cQueries = New Collection
         
         Sql = "Update OEVENTOAGENDA "
         Sql = Sql & " Set IDCLIENTE=" & SqlNum(nID)
         Sql = Sql & ", ALTERSTAMP=1"
         Sql = Sql & ", TIMESTAMP=GETDATE()"
         Sql = Sql & " Where IDCLIENTE=" & SqlNum(mvarMe.txtID.Text)
         Sql = Sql & " And IDLOJA=" & mvarIDLOJA
         cQueries.Add Sql
         
         Sql = "Update OATENDIMENTO"
         Sql = Sql & " Set IDCLIENTE=" & SqlNum(nID)
         Sql = Sql & ", ALTERSTAMP=1"
         Sql = Sql & ", TIMESTAMP=GETDATE()"
         Sql = Sql & " Where IDCLIENTE=" & SqlNum(mvarMe.txtID.Text)
         Sql = Sql & " And IDLOJA=" & mvarIDLOJA
         cQueries.Add Sql

         Sql = "Update OTAREFAEVT"
         Sql = Sql & " Set IDCLIENTE=" & SqlNum(nID)
         Sql = Sql & ", ALTERSTAMP=1"
         Sql = Sql & ", TIMESTAMP=GETDATE()"
         Sql = Sql & " Where IDCLIENTE=" & SqlNum(mvarMe.txtID.Text)
         Sql = Sql & " And IDLOJA=" & mvarIDLOJA
         cQueries.Add Sql
         
         Sql = "Update CVENDA"
         Sql = Sql & " Set IDCLIENTE=" & SqlNum(nID)
         Sql = Sql & ", ALTERSTAMP=1"
         Sql = Sql & ", TIMESTAMP=GETDATE()"
         Sql = Sql & " Where IDCLIENTE=" & SqlNum(mvarMe.txtID.Text)
         Sql = Sql & " And IDLOJA=" & mvarIDLOJA
         cQueries.Add Sql
         
         Sql = "Update FFATURA"
         Sql = Sql & " Set IDCLIENTE=" & SqlNum(nID)
         Sql = Sql & ", ALTERSTAMP=1"
         Sql = Sql & ", TIMESTAMP=GETDATE()"
         Sql = Sql & " Where IDCLIENTE=" & SqlNum(mvarMe.txtID.Text)
         Sql = Sql & " And IDLOJA=" & mvarIDLOJA
         cQueries.Add Sql
         
         Call .Executa(cQueries)
         Call PopulaTela
      Else
        Call ExibirStop("Id. Inválido", "Histórico")
      End If
   End With
End Sub

Private Sub mvarMe_CmdNovoItemClick()
   Dim objNG As Object
   Dim nAux As Long
   
   Screen.MousePointer = vbHourglass
   Select Case mvarMe.TabSessao.SelectedItem
      Case 2
         Set objNG = CriarObjeto("Tarefa3R.NG_Tarefa")
         Set objNG.Sys = mvarSys
         Call objNG.ExibirTarefa(pCod:=0, pIDCLIENTE:=mvarIDCLIENTE)
         
         mvarMe.GrdSessao.Redraw = False
         Call PopulaGrdSessao(mvarMe.TabSessao.SelectedItem)
         mvarMe.GrdSessao.Redraw = True
      Case 3
         'ExibirVenda(pCod As Long, Optional pCalControl)
         Set objNG = CriarObjeto("Financ3R.NG_Financeiro")
         Set objNG.Sys = mvarSys
         Call objNG.ExibirVenda(pCod:=0, pIDCLIENTE:=mvarIDCLIENTE)
         
         mvarMe.GrdSessao.Redraw = False
         Call PopulaGrdSessao(mvarMe.TabSessao.SelectedItem)
         mvarMe.GrdSessao.Redraw = True
      Case 4
         nAux = mvarIDCLIENTE
         Set objNG = CriarObjeto("Financ3R.NG_Financeiro")
         Set objNG.Sys = mvarSys
         Call objNG.ExibirFatura(pIDFATURA:=0, pIDCLIENTE:=nAux)
         
         mvarMe.GrdSessao.Redraw = False
         Call PopulaGrdSessao(mvarMe.TabSessao.SelectedItem)
         mvarMe.GrdSessao.Redraw = True
   End Select
   Set objNG = Nothing
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF5 Then
      Screen.MousePointer = vbHourglass
      Call PopulaTela
      Screen.MousePointer = vbDefault
   End If
End Sub
Private Sub mvarMe_Load()
   mvarMe.MousePointer = vbHourglass
   Screen.MousePointer = vbHourglass
   
   Call MontaTela
   Call VerificaMensagemI
   Call PopulaTela
   beMailVazio = (mvarMe.TxtEmail.Text = "")
End Sub

Private Sub mvarMe_CmbBairroLostFocus()
   Dim Sql As String
   Dim Rs As Object
   If mvarMe.cmbEstado.Text = "" Or mvarMe.CmbCidade.Text = "" Then
      Sql = "Select Distinct Top(1) BAIRRO, CIDADE, ESTADO, DTCADASTRO "
      Sql = Sql & " From OCLIENTE "
      Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
      Sql = Sql & " And Bairro =" & SqlStr(mvarMe.CmbBairro.Text)
      Sql = Sql & " Order By DTCADASTRO DESC"
      
      If mvarSys.xDb.AbreTabela(Sql, Rs) Then
         Call LocalizarCombo(mvarMe.CmbCidade, Rs("CIDADE"))
         Call LocalizarCombo(mvarMe.cmbEstado, Rs("ESTADO"))
      End If
   End If
End Sub

Private Sub mvarMe_CmbCidadeLostFocus()
   Dim Sql As String
   Dim Rs As Object
   If mvarMe.cmbEstado.Text = "" Then
      Sql = "Select Distinct Top(1) BAIRRO, CIDADE, ESTADO, DTCADASTRO"
      Sql = Sql & " From OCLIENTE "
      Sql = Sql & " Where IDLOJA =" & mvarIDLOJA
      Sql = Sql & " And Bairro =" & SqlStr(mvarMe.CmbBairro.Text)
      Sql = Sql & " And Cidade =" & SqlStr(mvarMe.CmbCidade.Text)
      Sql = Sql & " Order By DTCADASTRO DESC"
      If mvarSys.xDb.AbreTabela(Sql, Rs) Then
         Call LocalizarCombo(mvarMe.cmbEstado, Rs("ESTADO"))
      End If
   End If
End Sub

Private Sub mvarMe_CmdEditarClasseClick()
   Dim TbClasse   As Object
   Dim sDscClasse As String
   Dim bOk        As Boolean
      
   If mvarMe.TrvClasse.SelectedItem Is Nothing Then
      ExibirAviso "Selecione uma classe"
      Exit Sub
   Else
      If mvarMe.TrvClasse.SelectedItem.Key = "" Then
         ExibirAviso "Selecione uma classe"
         Exit Sub
      End If
   End If
   
   Set TbClasse = CriarObjeto("BANCO_3R.TB_OCLASSE")
   With TbClasse
      Set .xDb = mvarSys.xDb
      If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDCLASSE:=SqlNum(Mid(mvarMe.TrvClasse.SelectedItem.Key, 2))) Then
         sDscClasse = Trim(Mid(InputBox("Entre com novo nome da classe '" & mvarMe.TrvClasse.SelectedItem.Text & "'.", "Classe de Cliente"), 1, 20))
         If Trim(sDscClasse) <> "" Then
            .DSCCLASSE = sDscClasse
            If .Salvar Then
               mvarMe.TrvClasse.Nodes("k" & .IDCLASSE).Text = sDscClasse
               'DoEvents
               mvarMe.TrvClasse.SetFocus
               mvarMe.TrvClasse.Nodes("k" & .IDCLASSE).EnsureVisible
               mvarMe.TrvClasse.Nodes("k" & .IDCLASSE).Selected = True
               mvarMe.TrvClasse.Nodes("k" & .IDCLASSE).Checked = True
            End If
         End If
      End If
   End With
End Sub

Private Sub mvarMe_CmdExcluirClick()
   Dim Queries As Collection
   Dim TbCliClasse As Object
   If f_VerificarDependencias() Then
      If ExibirPergunta("Confirma a exclusão do cliente.", "Atenção", False) = vbYes Then
         If clsCliente.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDCLIENTE:=xVal(mvarMe.txtID)) Then
            Set Queries = New Collection
            
            Set TbCliClasse = CriarObjeto("BANCO_3R.TB_OCLASSE_CLIENTE")
            Set TbCliClasse.xDb = mvarSys.xDb
            
            Queries.Add TbCliClasse.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDCLIENTE:=xVal(mvarMe.txtID))
            Queries.Add clsCliente.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDCLIENTE:=xVal(mvarMe.txtID))

            If mvarSys.xDb.Executa(Queries) Then
               Call ExibirResultado(mvarSys, True)
               
               mvarIsDirt = True
               
               Call LimpaTela
               If mvarMe.txtNome.Visible And mvarMe.txtNome.Enabled Then
                  mvarMe.txtNome.SetFocus
               End If
               'If mvarMe.TxtRegistro.Visible And mvarMe.TxtRegistro.Enabled Then
               '   mvarMe.TxtRegistro.SetFocus
               'End If
            Else
               Call ExibirResultado(mvarSys, False)
            End If
         Else
            Call ExibirResultado(mvarSys, False)
         End If
      End If
   End If
End Sub
Private Sub mvarMe_CmdExtendeClick()
   Call SetTag(mvarMe.CmdExtende, "ABERTO", IIf(GetTag(mvarMe.CmdExtende, "ABERTO", "1") = "1", "2", "1"))
   Call AjustaGrpSessao(GetTag(mvarMe.CmdExtende, "ABERTO", "1"))
   
End Sub
Private Sub mvarMe_CmdIncluirClasseClick()
   Dim TbClasse   As Object
   Dim sDscClasse As String
   Dim bOk        As Boolean
   
   Set TbClasse = CriarObjeto("BANCO_3R.TB_OCLASSE")
   With TbClasse
      Set .xDb = mvarSys.xDb
      sDscClasse = Trim(Mid(InputBox("Entre com o nome da Classe.", "Classe de Cliente"), 1, 20))
      If Trim(sDscClasse) <> "" Then
         If Not .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="DSCCLASSE=" & SqlStr(sDscClasse)) Then
            '.IDPAI
            .IDLOJA = mvarIDLOJA
            .DSCCLASSE = sDscClasse
            .Ativo = 1
         End If
         If .Salvar Then
            Call MontaTrv
            Call PopulaTrv
            'DoEvents
            mvarMe.TrvClasse.SetFocus
            mvarMe.TrvClasse.Nodes("k" & .IDCLASSE).EnsureVisible
            mvarMe.TrvClasse.Nodes("k" & .IDCLASSE).Selected = True
            mvarMe.TrvClasse.Nodes("k" & .IDCLASSE).Checked = True
         End If
      End If
   End With
End Sub
Private Sub mvarMe_cmdLovClick()
   Dim x As New TL_NovoLov
   
   Set x.Sys = mvarSys
   'x.Filtro = "Flavia"
   Call x.ShowClientes
   If x.IdSelecionado <> 0 Then
      mvarIDCLIENTE = x.IdSelecionado
      Call LimpaTela
      Call PopulaTela
   End If
   Set x = Nothing
End Sub
Private Sub mvarMe_cmdNovoClick()
   Dim bOk As Boolean
   
   If xVal(mvarMe.txtID) = 0 And Trim(mvarMe.txtNome) = "" Then
      Call LimpaTela
      mvarMe.txtNome.SetFocus
      'mvarMe.TxtRegistro.SetFocus
   Else
      If Salvar Then
         Call LimpaTela
         Call f_InstanciaClasseCliente
         mvarMe.txtNome.SetFocus
         'mvarMe.TxtRegistro.SetFocus
      Else
         mvarMe.txtNome.SetFocus
         'mvarMe.TxtRegistro.SetFocus
      End If
   End If
End Sub
Private Sub mvarMe_CmdOkClick()
   If Salvar Then Unload mvarMe
End Sub
Private Sub mvarMe_CmdSairClick()
   Unload mvarMe
End Sub
Private Sub mvarMe_cmdSalvarClick()
   Call Salvar
End Sub
Private Function Salvar() As Boolean
   Dim bOk     As Boolean
   Dim xNode   As TreeViewNode
   Dim bIsDirt As Boolean
   Dim iMouse  As Integer
   Dim Sql As String
  
   iMouse = Screen.MousePointer
   Screen.MousePointer = vbHourglass
   If f_ValidaCampos() Then
      Call PopulaClasse
      bIsDirt = clsCliente.IsDirt
      If Not bIsDirt Then
         For Each xNode In mvarMe.TrvClasse.Nodes
            If GetTag(xNode, "ISDIRT", "0") = "1" Then
               bIsDirt = True
               Exit For
            End If
         Next
      End If
      bOk = True
      If bIsDirt Then
         If beMailVazio And clsCliente.EMAIL <> "" Then
            Call VerificarBoasVindas
         End If
         bOk = False
         If clsCliente.Salvar() Then
            mvarIDCLIENTE = clsCliente.IDCLIENTE
            bOk = SalvarClasse
            beMailVazio = (mvarMe.TxtEmail.Text = "")
            
            '* Atualiza Título dos Eventos
            If bNmDirt Then
               Call AtualizarTituloEvt
            End If
         End If
         mvarIsDirt = mvarIsDirt Or bOk
      End If
      Call ExibirResultado(mvarSys, bOk, IIf(bOk, 1, 2))
   End If
   Salvar = bOk
   Screen.MousePointer = iMouse
End Function
Private Function SalvarClasse() As Boolean
   Dim xNode         As TreeViewNode
   Dim Queries       As Collection
   Dim TbClasseCli   As Object
   
   Set Queries = New Collection
   
   Set TbClasseCli = CriarObjeto("BANCO_3R.TB_OCLASSE_CLIENTE")
   Set TbClasseCli.xDb = mvarSys.xDb
   
   TbClasseCli.IDLOJA = mvarIDLOJA
   TbClasseCli.IDCLIENTE = mvarIDCLIENTE
   
   Queries.Add TbClasseCli.QryDelete(Ch_IDLOJA:=mvarIDLOJA, Ch_IDCLIENTE:=mvarIDCLIENTE)
   For Each xNode In mvarMe.TrvClasse.Nodes
      If xNode.Checked Then
         TbClasseCli.IDCLASSE = xVal(Mid(xNode.Key, 2))
         Queries.Add TbClasseCli.QrySave
      End If
   Next
   If Queries.Count > 0 Then
      SalvarClasse = mvarSys.xDb.Executa(Queries)
   Else
      SalvarClasse = True
   End If
End Function
Private Sub VerificarBoasVindas()
   Dim Sql As String
   Dim sMsg As String
   Dim NgTarefa As Object
   Dim oRs As Object
   Dim bOk As Boolean
   Dim sAux As String
   Dim i As Integer
   
   Sql = "Select Top 1 * "
   Sql = Sql & " From OATENDIMENTO"
   Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
   Sql = Sql & " And IDCLIENTE=" & clsCliente.IDCLIENTE
   Sql = Sql & " And DTATEND <= " & SqlDate(mvarSys.xDb.SysDate() & " 23:59:59")
   Sql = Sql & " Order By DTATEND Desc"
   If mvarSys.xDb.AbreTabela(Sql, oRs) Then
      If oRs.recordcount = 1 Then
         If DateDiff("d", oRs("DTATEND"), mvarSys.xDb.SysDate()) < 3 Then
            bOk = True
         Else
            sMsg = ""
            sMsg = sMsg & "O e-Mail estava desatualizado e o único " & vbNewLine
            sMsg = sMsg & "atendimento, deste cliente, ocorreu em: " & vbNewLine & vbNewLine
            sMsg = sMsg & DiaDaSemana(oRs("DTATEND")) & ", " & Format(oRs("DTATEND"), "dd/mm/yy") & vbNewLine & vbNewLine
            sMsg = sMsg & "Deseja criar tarefa de Boas Vindas?"
            
            bOk = (vbYes = ExibirPergunta(sMsg, "Boas Vindas!"))
            
         End If
         If bOk Then
            '* Salvar Tarefas
            Set NgTarefa = CreateObject("BANCO_3R.TB_OTAREFAEVT")
            With NgTarefa
               Set .xDb = mvarSys.xDb
               Sql = " IDCLIENTE=" & mvarIDCLIENTE
               Sql = Sql & " And IDTPTAREFA=" & eTpTarefa.TarBoasVindas
               Call .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:=Sql)
               .IDLOJA = mvarIDLOJA
               .IDEVENTO = oRs("IDEVENTO") & ""
               .DSCTAREFA = ""
               .TITULO = clsCliente.NOME & " - " & IIf(Trim(clsCliente.TEL1) = "", clsCliente.TEL2, clsCliente.TEL1)
               If .SITTAREFA <> 10 Then
                  i = 1
                  sAux = DateAdd("d", i, mvarSys.xDb.SysDate())
                  While eFeriado(mvarSys.xDb, sAux) Or Weekday(sAux) = 1 Or Weekday(sAux) = 7
                     i = i + 1
                     sAux = DateAdd("d", i, mvarSys.xDb.SysDate())
                  Wend
                  .DTTAREFA = sAux
               End If
               '.SITTAREFA = 0
               .IDCLIENTE = clsCliente.IDCLIENTE
               .PRIORIDADE = 1
               .IDTPTAREFA = eTpTarefa.TarBoasVindas
               If .IsDirt Then
                  Call .Salvar
               End If
            End With
            Set NgTarefa = Nothing
         End If
      End If
   End If
End Sub
Private Sub AtualizarTituloEvt()
   Dim Sql As String
   Dim sNome As String
   Dim sTel  As String
   Dim sAux  As String
   Dim i     As Integer
   
   '* Atualiza Título dos Eventos
   With mvarMe
      sNome = Trim(.txtNome.Text)
      sTel = IIf(Trim(.txtTel.Text) = "", Trim(.txtCelular.Text), Trim(.txtTel.Text))
      For i = 1 To Len(sTel)
         If isAlfaNum(Mid(sTel, i, 1)) Then
            sAux = sAux & Mid(sTel, i, 1)
         End If
      Next
      sTel = sAux
      sAux = sNome & IIf(sTel = "", "", " - ") & sTel
      Sql = "Update OEVENTOAGENDA "
      Sql = Sql & " Set Subject = " & SqlStr(sAux)
      Sql = Sql & " , ALTERSTAMP=1"
      Sql = Sql & " , TIMESTAMP=GETDATE()"
      Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
      Sql = Sql & " And  IDCLIENTE=" & mvarIDCLIENTE
      Call mvarSys.xDb.Executa(Sql)
   End With
End Sub
Private Function f_ValidaCampos() As Boolean
   ' *******************
   ' *******************
   ' *******************
   Dim bResult As Boolean
   
   bResult = True
   If Trim(mvarMe.txtTel.Text) = "" And Trim(mvarMe.txtCelular.Text) = "" And Trim(mvarMe.TxtEmail.Text) = "" Then
      mvarMe.TabContato.Item(0).Selected = True
      ExibirAviso "Ao menos um contato, telefone ou e-Mail, deve ser preenchido."
      bResult = False
      mvarMe.txtCelular.SetFocus
      GoTo Fim
   End If
   
   If Trim(mvarMe.txtNome.Text) = "" Then
      mvarMe.TabContato.Item(0).Selected = True
      ExibirAviso "O nome deve ser preenchido."
      bResult = False
      mvarMe.txtNome.SetFocus
      GoTo Fim
   End If
   
   If Trim(mvarMe.TxtEmail.Text) <> "" Then
      If InStr(1, mvarMe.TxtEmail.Text, "@", vbTextCompare) <= 1 Or Len(Trim(mvarMe.TxtEmail.Text)) < 6 Then
         mvarMe.TabContato.Item(0).Selected = True
         ExibirAviso "Email inválido."
         bResult = False
         mvarMe.TxtEmail.SetFocus
         GoTo Fim
      End If
   End If
   If Not mvarMe.OptATIVO(1).Value Then
      If Trim(mvarMe.TxtMotivoInat.Text) = "" Then
         mvarMe.TabContato.Item(2).Selected = True
         ExibirAviso "Campo inválido. O campo motivo deve ser preenchido."
         bResult = False
         mvarMe.TxtMotivoInat.SetFocus
         GoTo Fim
      End If
   End If
Fim:
   f_ValidaCampos = bResult
End Function

Public Function ExcluirCliente()
   If mvarIDCLIENTE <> 0 Then
      Call PopulaTela
      mvarMe_CmdExcluirClick
   End If
End Function

Private Function f_VerificarDependencias() As Boolean
   Dim bResult    As Boolean
   Dim TbEvento   As Object
   Dim TBAtend    As Object
   Dim TbVenda    As Object
   Dim sMsg       As String
   
   Set TbEvento = CriarObjeto("BANCO_3R.TB_OEVENTOAGENDA")
   With TbEvento
      Set .xDb = mvarSys.xDb
      If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="IDCLIENTE=" & SqlNum(mvarMe.txtID)) Then
         sMsg = vbNewLine & vbNewLine
         While Not .Rs.EOF
            .Popula
            sMsg = sMsg & "Evento : " & .IDEVENTO
            sMsg = sMsg & " - " & .Subject
            sMsg = sMsg & " (" & Format(.StartDateTime, "dd/mm/yyyy")
            sMsg = sMsg & " " & Format(.StartDateTime, "hh:mm")
            sMsg = sMsg & " - " & Format(.EndDateTime, "hh:mm") & ")"
            sMsg = sMsg & vbNewLine
            .Rs.MoveNext
         Wend
         Call ExibirAviso("Cliente possui evento associado e não pode ser excluído." & sMsg)
         GoTo Saida
      Else
         Set TBAtend = CriarObjeto("BANCO_3R.TB_OATENDIMENTO")
         With TBAtend
            Set .xDb = mvarSys.xDb
            If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="IDCLIENTE=" & SqlNum(mvarMe.txtID)) Then
               sMsg = vbNewLine & vbNewLine
               While Not .Rs.EOF
                  .Popula
                  sMsg = sMsg & "Atendimento : " & .IDATENDIMENTO
                  sMsg = sMsg & " (" & .DTATEND & " " & .HHINI & " - " & .HHFIM & ")"
                  sMsg = sMsg & vbNewLine
                  .Rs.MoveNext
               Wend
               
               Call ExibirAviso("Cliente possui atendimento associado e não pode ser excluído." & sMsg)
               GoTo Saida
            Else
               Set TbVenda = CriarObjeto("BANCO_3R.TB_CVENDA")
               With TbVenda
                  Set .xDb = mvarSys.xDb
                  If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="IDCLIENTE=" & SqlNum(mvarMe.txtID)) Then
                     sMsg = vbNewLine & vbNewLine
                     While Not .Rs.EOF
                        .Popula
                        sMsg = sMsg & "Venda : " & .IDVENDA
                        sMsg = sMsg & " (" & .DTVENDA & " " & .VLVENDA & ")"
                        sMsg = sMsg & vbNewLine
                        .Rs.MoveNext
                     Wend
                     Call ExibirAviso("Cliente possui venda associada e não pode ser excluído." & sMsg)
                     GoTo Saida
                  End If
               End With
            End If
         End With
      End If
   End With
   
   bResult = True
   
   f_VerificarDependencias = bResult
Saida:
   Set TbEvento = Nothing
   Set TBAtend = Nothing
   Set TbVenda = Nothing
End Function

Private Sub f_InstanciaClasseCliente()
   Set clsCliente = Nothing
   Set clsCliente = CriarObjeto("BANCO_3R.TB_OCLIENTE")
   Set clsCliente.xDb = mvarSys.xDb
End Sub
Private Sub mvarMe_GrdSessaoDblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)
   Dim objNG    As Object
   Dim nIDATEND As Long
   Dim nIDVENDA As Long
   Dim nIDEVENT As Long
   Dim nIDTAREFA As Long
   Dim nIDFATURA As Long

   
   If mvarMe.GrdSessao.RowCount <= 0 Then Exit Sub
   If lRow <= 0 Then Exit Sub
   
   Screen.MousePointer = vbHourglass
   If mvarMe.TabSessao.SelectedItem = 0 Then
      nIDATEND = xVal(mvarMe.GrdSessao.CellValue(lRow, "IDATENDIMENTO") & "")
      nIDVENDA = xVal(mvarMe.GrdSessao.CellValue(lRow, "IDVENDA") & "")
   ElseIf mvarMe.TabSessao.SelectedItem = 1 Then
      nIDEVENT = xVal(mvarMe.GrdSessao.CellValue(lRow, "IDEVENTO") & "")
   ElseIf mvarMe.TabSessao.SelectedItem = 2 Then
      nIDTAREFA = xVal(mvarMe.GrdSessao.CellValue(lRow, "IDTAREFA") & "")
   ElseIf mvarMe.TabSessao.SelectedItem = 3 Then
      nIDVENDA = xVal(mvarMe.GrdSessao.CellValue(lRow, "VENDA") & "")
   ElseIf mvarMe.TabSessao.SelectedItem = 4 Then
      nIDFATURA = xVal(mvarMe.GrdSessao.CellValue(lRow, "FATURA") & "")
   End If
   
   If nIDATEND > 0 Then
      Set objNG = CriarObjeto("Calendario3R.NG_Calendario")
      Set objNG.Sys = mvarSys
      Call objNG.ExibirAtendimento(nIDATEND)

   ElseIf nIDVENDA > 0 Then
      Set objNG = CriarObjeto("Financ3R.NG_Financeiro")
      Set objNG.Sys = mvarSys
      Call objNG.ExibirVenda(nIDVENDA)
   
   ElseIf nIDEVENT > 0 Then
      Set objNG = CriarObjeto("Calendario3R.NG_Calendario")
      Set objNG.Sys = mvarSys
      Call objNG.ExibirEvento(nIDEVENT)
      
   ElseIf nIDTAREFA > 0 Then
      Set objNG = CriarObjeto("Tarefa3R.NG_Tarefa")
      Set objNG.Sys = mvarSys
      Call objNG.ExibirTarefa(nIDTAREFA)
   
   ElseIf nIDFATURA > 0 Then
      Set objNG = CriarObjeto("Financ3R.NG_Financeiro")
      Set objNG.Sys = mvarSys
      Call objNG.ExibirFatura(nIDFATURA)
            
   End If
   mvarMe.GrdSessao.Redraw = False
   Call PopulaGrdSessao(mvarMe.TabSessao.SelectedItem)
   mvarMe.GrdSessao.Redraw = True
   
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_LblProxDblClick()
   Dim oNGCal As Object
   Dim nIDEVENTO  As Long
   Screen.MousePointer = vbHourglass
   
   nIDEVENTO = xVal(GetTag(mvarMe.LblProx, "IDEVENTO", 0))
   If nIDEVENTO > 0 Then
      Set oNGCal = CriarObjeto("Calendario3R.NG_Calendario")
      Set oNGCal.Sys = mvarSys
      Call oNGCal.ExibirEvento(nIDEVENTO)
   End If
   Set oNGCal = Nothing
   Screen.MousePointer = vbDefault
   mvarMe.txtNome.SetFocus
End Sub
Private Sub VerificaMensagemI()
   Dim Sql As String
   Dim sMsg As String
   
   Sql = "Select F.IDFATURA, C.NOME"
   Sql = Sql & " From FFATURA F"
   Sql = Sql & " Join OCLIENTE C On C.IDLOJA=F.IDLOJA And C.IDCLIENTE=F.IDCLIENTE"
   Sql = Sql & " Where F.SITFATURA = 0"
   Sql = Sql & " And F.IDLOJA = " & mvarIDLOJA
   Sql = Sql & " And F.IDCLIENTE = " & mvarIDCLIENTE
   
   If mvarSys.xDb.AbreTabela(Sql) Then
      With mvarSys.xDb.RsAux
         If .recordcount > 0 Then
            sMsg = "Cliente: " & !NOME
            sMsg = Mid(sMsg, 1, 28) & IIf(Len(sMsg) > 30, "...", "") & vbNewLine
            sMsg = sMsg & "Possui fatura " & StrZero(!IDFATURA, 6) & " 'Em Aberto'"
            Call ExibirMensagemI(mvarSys, sMsg)
         End If
      End With
   End If
End Sub

Private Sub MontaTela()
   Dim sAux       As String
   Dim Sql        As String
   Dim MyRs         As Object
   
   mvarMe.Icon = mvarSys.MDI.Icon
   mvarMe.CmdExpHis.Visible = False
   
   mvarMe.TxtRegistro.SetMask "000.000.000-00", "___.___.___-__"
   mvarMe.txtTel.MaxLength = gMaxLenTel
   mvarMe.txtCelular.MaxLength = gMaxLenTel
   mvarMe.txtFax.MaxLength = 20 'gMaxLenTel
   
   mvarMe.TabContato.Item(3).Visible = True
   mvarMe.TabContato.Item(0).Selected = True
   'mvarMe.TabContato.Color = xtpTabColorOffice2003
   
   mvarMe.CmdExtende.Visible = (mvarSys.Propriedades("TPAREA") <> "0")
   
   Call AjustaGrpSessao(1)
   
   '*******
   '* Carrega Fototipo
   mvarMe.CmbFototipo.Clear
   mvarMe.CmbFototipo.Style = xtpComboDropDownList
   
   Sql = "Select * "
   Sql = Sql & " From OFOTOTIPO"
   Sql = Sql & " Order By FOTOTIPO"
   If mvarSys.xDb.AbreTabela(Sql, MyRs) Then
      While Not MyRs.EOF
         sAux = GetFOTOTIPOText(MyRs("FOTOTIPO"), MyRs("PELE"))
         mvarMe.CmbFototipo.AddItem sAux
         mvarMe.CmbFototipo.ItemData(mvarMe.CmbFototipo.NewIndex) = MyRs("FOTOTIPO")
         MyRs.MoveNext
      Wend
      mvarMe.CmbFototipo.ListIndex = 0
   End If
   
   '*******
   '* Carrega Funcionarios
   mvarMe.CmbFunc.Clear
   mvarMe.CmbFunc.Style = xtpComboDropDownList
   mvarMe.CmbFunc.AddItem " "
   mvarMe.CmbFunc.ItemData(mvarMe.CmbFunc.NewIndex) = 0
   
   Sql = "Select IDFUNCIONARIO, NOME "
   Sql = Sql & " From RFUNCIONARIO"
   Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
   Sql = Sql & " And FLGCERTIFICADO=1"
   Sql = Sql & " And SITFUNC='A'"
   Sql = Sql & " Order By NOME"
   If mvarSys.xDb.AbreTabela(Sql, MyRs) Then
      While Not MyRs.EOF
         sAux = MyRs("NOME")
         mvarMe.CmbFunc.AddItem sAux
         mvarMe.CmbFunc.ItemData(mvarMe.CmbFunc.NewIndex) = MyRs("IDFUNCIONARIO")
         MyRs.MoveNext
      Wend
      mvarMe.CmbFunc.ListIndex = 0
   End If
   
   '*******
   '* Carrega Conhecimento
   mvarMe.CmbConhecimento.Clear
   mvarMe.CmbConhecimento.Style = xtpComboDropDownList
   mvarMe.CmbConhecimento.AddItem " "
   mvarMe.CmbConhecimento.ItemData(mvarMe.CmbConhecimento.NewIndex) = 0
   
   Sql = "Select * "
   Sql = Sql & " From OTPCONHEC"
   Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
   Sql = Sql & " And IDTPCONHEC<>0"
   Sql = Sql & " Order By IDTPCONHEC"
   If mvarSys.xDb.AbreTabela(Sql, MyRs) Then
      While Not MyRs.EOF
         sAux = CStr(MyRs("IDTPCONHEC")) + " - " + MyRs("NMCONHEC")
         mvarMe.CmbConhecimento.AddItem sAux
         mvarMe.CmbConhecimento.ItemData(mvarMe.CmbConhecimento.NewIndex) = MyRs("IDTPCONHEC")
         MyRs.MoveNext
      Wend
      mvarMe.CmbConhecimento.ListIndex = 0
   End If
   
   '*******
   '* Carrega Profissao
   mvarMe.CmbProfissao.Clear
   mvarMe.CmbProfissao.Style = xtpComboDropDownList
   mvarMe.CmbProfissao.AddItem ""
   Sql = "Select Distinct RTrim(LTrim(Profissao)) [Profissao] "
   Sql = Sql & " From OCLIENTE "
   Sql = Sql & " Where IDCOLIGADA=" & mvarSys.Propriedades("IDCOLIGADA")
   Sql = Sql & " And RTrim(LTrim(Profissao))<>''"
   Sql = Sql & " Order by 1"
   If mvarSys.xDb.AbreTabela(Sql, MyRs) Then
      While Not MyRs.EOF
         If Trim(MyRs("Profissao")) <> "" And Trim(MyRs("Profissao")) <> "-1" Then
            mvarMe.CmbProfissao.AddItem Trim(MyRs("Profissao"))
         End If
         MyRs.MoveNext
      Wend
   End If
   If mvarMe.CmbProfissao.ListCount = 2 Then
      mvarMe.CmbProfissao.ListIndex = 1
   End If
   
   '*******
   '* Carrega Bairro
   mvarMe.CmbBairro.Clear
   mvarMe.CmbBairro.Style = xtpComboDropDownList
   mvarMe.CmbBairro.AddItem ""
   Sql = "Select Distinct RTrim(LTrim(Bairro)) [Bairro] "
   Sql = Sql & " From OCLIENTE "
   Sql = Sql & " Where IDCOLIGADA=" & mvarSys.Propriedades("IDCOLIGADA")
   Sql = Sql & " And RTrim(LTrim(Bairro))<>''"
   Sql = Sql & " Order by 1"
   If mvarSys.xDb.AbreTabela(Sql, MyRs) Then
      While Not MyRs.EOF
         If Trim(MyRs("Bairro")) <> "" Then
            mvarMe.CmbBairro.AddItem Trim(MyRs("Bairro"))
         End If
         MyRs.MoveNext
      Wend
   End If
   If mvarMe.CmbBairro.ListCount = 2 Then
      mvarMe.CmbBairro.ListIndex = 1
   End If
   
   '*******
   '* Carrega Cidade
   mvarMe.CmbCidade.Clear
   mvarMe.CmbCidade.Style = xtpComboDropDownList
   mvarMe.CmbCidade.AddItem ""
   Sql = "Select Distinct RTrim(LTrim(Cidade)) [Cidade] "
   Sql = Sql & " From OCLIENTE "
   Sql = Sql & " Where IDCOLIGADA=" & mvarSys.Propriedades("IDCOLIGADA")
   Sql = Sql & " And RTrim(LTrim(Cidade))<>''"
   Sql = Sql & " Order by 1"
   If mvarSys.xDb.AbreTabela(Sql, MyRs) Then
      While Not MyRs.EOF
         If Trim(MyRs("Cidade")) <> "" Then
            mvarMe.CmbCidade.AddItem Trim(MyRs("Cidade"))
         End If
         MyRs.MoveNext
      Wend
   End If
   If mvarMe.CmbCidade.ListCount = 2 Then
      mvarMe.CmbCidade.ListIndex = 1
   End If
   
   '*******
   '* Carrega Estado
   mvarMe.cmbEstado.Clear
   mvarMe.cmbEstado.Style = xtpComboDropDownList
   mvarMe.cmbEstado.AddItem ""
   Sql = "Select Distinct RTrim(LTrim(Estado)) [Estado] "
   Sql = Sql & " From OCLIENTE "
   Sql = Sql & " Where IDCOLIGADA=" & mvarSys.Propriedades("IDCOLIGADA")
   Sql = Sql & " And RTrim(LTrim(Estado))<>''"
   Sql = Sql & " Order by 1"
   If mvarSys.xDb.AbreTabela(Sql, MyRs) Then
      While Not MyRs.EOF
         If Trim(MyRs("Estado")) <> "" Then
            mvarMe.cmbEstado.AddItem Trim(MyRs("Estado"))
         End If
         MyRs.MoveNext
      Wend
   End If
   If mvarMe.cmbEstado.ListCount = 2 Then
      mvarMe.cmbEstado.ListIndex = 1
   End If
   '*******
   '* Carrega Sexo
   mvarMe.CmbSexo.Clear
   mvarMe.CmbSexo.Style = xtpComboDropDownList
   mvarMe.CmbSexo.AddItem "Feminino"
   mvarMe.CmbSexo.AddItem "Masculino"
   
   '*******
   '* Carrega Classe
   Call MontaTrv
   
   Call LimpaTela
End Sub
Private Sub MontaTrv()
   Dim xNode      As TreeViewNode
   Dim Sql        As String
   Dim MyRs       As Object
   
   mvarMe.TrvClasse.Nodes.Clear
   mvarMe.TrvClasse.Scroll = False
   
   Sql = "Select *"
   Sql = Sql & " From OCLASSE"
   Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
   Sql = Sql & " Order By IDPAI, DSCCLASSE"
   If mvarSys.xDb.AbreTabela(Sql, MyRs) Then
      While Not MyRs.EOF
         If MyRs("IDPAI") = 0 Then
            Set xNode = mvarMe.TrvClasse.Nodes.Add(, , "k" & MyRs("IDCLASSE"), MyRs("DSCCLASSE"))
         Else
            Set xNode = mvarMe.TrvClasse.Nodes.Add("k" & MyRs("IDPAI"), tvwChild, "k" & MyRs("IDCLASSE"), MyRs("DSCCLASSE"))
            mvarMe.TrvClasse.Nodes("k" & MyRs("IDPAI")).Bold = True
         End If
         xNode.Expanded = True
         Call SetTag(xNode, "ISDIRT", "0")
         Call SetTag(xNode, "VALUE", "0")
         MyRs.MoveNext
      Wend
   End If
   mvarMe.TrvClasse.Scroll = True
End Sub
Private Function GetFOTOTIPOText(pFOTOTIPO As Integer, Optional pPele) As String
   Dim sAux As String
   Dim TbFoto As Object
   
   Select Case pFOTOTIPO
      Case 1: sAux = "I"
      Case 2: sAux = "II"
      Case 3: sAux = "III"
      Case 4: sAux = "IV"
      Case 5: sAux = "V"
      Case 6: sAux = "VI"
   End Select
   If IsMissing(pPele) Then
      Set TbFoto = CriarObjeto("BANCO_3R.TB_OFOTOTIPO")
      Set TbFoto.xDb = mvarSys.xDb
      If TbFoto.Pesquisar(Ch_FOTOTIPO:=pFOTOTIPO) Then
         GetFOTOTIPOText = sAux & " - " & TbFoto.PELE
      Else
         GetFOTOTIPOText = sAux
      End If
   Else
      GetFOTOTIPOText = sAux & " - " & pPele
   End If
   Set TbFoto = Nothing
End Function

Private Sub mvarMe_TabSessaoSelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
   
   mvarMe.GrdSessao.Redraw = False
   mvarMe.CmdExtende.Visible = (Item.Index = 0) And (mvarSys.Propriedades("TPAREA") <> "0")
   mvarMe.GrpValor.Visible = InArray(Item.Index, Array(0, 3, 4))
   mvarMe.LblTotal1.Visible = mvarMe.GrpValor.Visible
   mvarMe.CmdNovoItem.Visible = InArray(Item.Index, Array(3, 4))
   Select Case Item.Index
      Case 0
         'mvarMe.CmdExtende.Visible = True
         mvarMe.GrpValor.Visible = True
         mvarMe.LblTotal1.Visible = True
         mvarMe.LblVALOR.Visible = mvarMe.LblTotal1.Visible
         mvarMe.GrdSessao.Height = mvarMe.GrpValor.Top + mvarMe.GrpValor.Height - 960
         mvarMe.GrdSessao.ForeColor = vbBlack
      Case 1
         mvarMe.GrdSessao.Height = mvarMe.GrpValor.Top + 30
         'mvarMe.GrdSessao.ForeColor = vbRed
      Case 2
         mvarMe.CmdNovoItem.Visible = True
         mvarMe.CmdNovoItem.Caption = "&Gerar Tarefa"
         mvarMe.GrpValor.Visible = True
         mvarMe.LblTotal1.Visible = False
         mvarMe.LblVALOR.Visible = mvarMe.LblTotal1.Visible
         mvarMe.LblTotal1.Left = mvarMe.LblVALOR.Left - mvarMe.LblTotal1.Width - 60
         mvarMe.GrdSessao.Height = mvarMe.GrpValor.Top + mvarMe.GrpValor.Height - 960
         mvarMe.GrdSessao.ForeColor = vbBlack
      Case 3
         'mvarMe.CmdExtende.Visible = True
         mvarMe.CmdNovoItem.Visible = True
         mvarMe.CmdNovoItem.Caption = "&Gerar Venda"
         mvarMe.GrpValor.Visible = True
         mvarMe.LblTotal1.Visible = True
         mvarMe.LblVALOR.Visible = mvarMe.LblTotal1.Visible
         mvarMe.LblTotal1.Left = mvarMe.LblVALOR.Left - mvarMe.LblTotal1.Width - 60
         mvarMe.GrdSessao.Height = mvarMe.GrpValor.Top + mvarMe.GrpValor.Height - 960
         mvarMe.GrdSessao.ForeColor = vbBlack
      Case 4
         mvarMe.CmdNovoItem.Visible = True
         mvarMe.CmdNovoItem.Caption = "&Gerar Fatura"
         mvarMe.GrpValor.Visible = True
         mvarMe.LblTotal1.Visible = True
         mvarMe.LblVALOR.Visible = mvarMe.LblTotal1.Visible
         mvarMe.LblTotal1.Left = mvarMe.LblVALOR.Left - mvarMe.LblTotal1.Width - 60
         mvarMe.GrdSessao.Height = mvarMe.GrpValor.Top + mvarMe.GrpValor.Height - 960
         mvarMe.GrdSessao.ForeColor = vbBlack
   End Select
   Call PopulaGrdSessao(Item.Index)
   mvarMe.GrdSessao.Redraw = True
End Sub

Private Sub mvarMe_TrvClasseDblClick()
   If Not mvarMe.TrvClasse.SelectedItem Is Nothing Then
      mvarMe.TrvClasse.SelectedItem.Checked = Not mvarMe.TrvClasse.SelectedItem.Checked
   End If
End Sub
Private Sub mvarMe_TrvClasseNodeCheck(ByVal Node As XtremeSuiteControls.TreeViewNode)
   Dim sValue  As String
   
   sValue = GetTag(Node, "VALUE", "0")
   Call SetTag(Node, "ISDIRT", "1")
   
   If Node.Checked And sValue = "1" Then Call SetTag(Node, "ISDIRT", "0")
   If Not Node.Checked And sValue = "0" Then Call SetTag(Node, "ISDIRT", "0")
End Sub
Private Sub mvarMe_TxtDTNASCLostFocus()
   mvarMe.TxtDTNASC.Text = FormatarData(mvarMe.TxtDTNASC.Text)
   Call DefineIdade
End Sub
Private Sub DefineIdade()
   Dim nIdade As Integer
   
   mvarMe.LblMenor.Visible = False
   mvarMe.LblMenor2.Visible = False
   If IsDate(mvarMe.TxtDTNASC.Text) Then
      nIdade = DateDiff("yyyy", mvarMe.TxtDTNASC.Text, mvarSys.xDb.SysDate)
      If nIdade <= 18 Then
         If nIdade = 18 Then
            If CDate(mvarSys.xDb.SysDate) < CDate(Mid(mvarMe.TxtDTNASC.Text, 1, 6) & Year(mvarSys.xDb.SysDate)) Then
               mvarMe.LblMenor.ForeColor = vbRed
               mvarMe.LblMenor2.ForeColor = vbRed
               mvarMe.LblMenor.Visible = True
               mvarMe.LblMenor2.Visible = True
               mvarMe.LblMenor.Caption = "Menor de"
               mvarMe.LblMenor2.Caption = "Idade"
            Else
               mvarMe.LblMenor.ForeColor = &H808080
               mvarMe.LblMenor2.ForeColor = &H808080
               mvarMe.LblMenor.Visible = True
               mvarMe.LblMenor2.Visible = True
               mvarMe.LblMenor.Caption = nIdade
               mvarMe.LblMenor2.Caption = "anos"
            End If
         Else
            mvarMe.LblMenor.Visible = True
            mvarMe.LblMenor2.Visible = True
            mvarMe.LblMenor.Caption = "Menor de"
            mvarMe.LblMenor2.Caption = "Idade"
         End If
      Else
         mvarMe.LblMenor.ForeColor = &H808080
         mvarMe.LblMenor2.ForeColor = &H808080
         mvarMe.LblMenor.Visible = True
         mvarMe.LblMenor2.Visible = True
         mvarMe.LblMenor.Caption = nIdade
         mvarMe.LblMenor2.Caption = "anos"
      End If
   End If
End Sub
Private Sub mvarMe_TxtEmailLostFocus()
   mvarMe.TxtEmail.Text = Replace(Trim(mvarMe.TxtEmail.Text), " ", "")
End Sub
Private Sub mvarMe_TxtNomeLostFocus()
   Dim Sql As String
   Dim Rs As Object
   
   mvarMe.txtNome.Text = FormatarNome(mvarMe.txtNome.Text)
   If mvarMe.txtNome.Text <> "" And mvarMe.CmbSexo.Text = "" Then
      Sql = "Select Distinct Top(1) SEXO, DTCADASTRO"
      Sql = Sql & " From OCLIENTE "
      Sql = Sql & " Where Like '" & Trim(GetWords(mvarMe.txtNome.Text)(1)) & "%'"
      Sql = Sql & " Order By DTCADASTRO DESC"
      If mvarSys.xDb.AbreTabela(Sql, Rs) Then
         Call LocalizarCombo(mvarMe.CmbSexo, Rs("SEXO"))
      End If
   End If
End Sub

Private Sub mvarMe_TxtRegistroGotFocus()
   'mvarMe.TxtRegistro.Text = UnFormat(mvarMe.TxtRegistro.Text)
   mvarMe.TxtRegistro.SelStart = 0
   mvarMe.TxtRegistro.SelLength = Len(mvarMe.TxtRegistro.Text)
End Sub
Private Sub mvarMe_TxtRegistroLostFocus()
   'mvarMe.TxtRegistro.Text = Format(mvarMe.TxtRegistro.Text, "###.###.###-##")
   'mvarMe.TxtRegistro.Text = mvarMe.TxtRegistro.Text
   If Not ValidarCPF(mvarMe.TxtRegistro.Text) Then
      Call ExibirAviso("CPF Inválido", "Clientes")
      If mvarMe.TxtRegistro.Visible And mvarMe.TxtRegistro.Enabled Then
         mvarMe.TxtRegistro.SetFocus
      End If
   End If
End Sub
Private Sub AjustaGrpSessao(pPos As Integer)

   With mvarMe
      .GrpProxSessao.Width = 2150
      .LblProx.AutoSize = True
      .LblLblProx.AutoSize = True
      With .GrdSessao
         If pPos = 1 Then 'If .Width > 8000 Then
            mvarMe.CmdExtende.Caption = "»"
            .Width = 6425
            If .Visible And .RowCount > 0 Then Call .Sort(Array("#"))
         ElseIf pPos = 2 Then
            .Width = 8025
            mvarMe.CmdExtende.Caption = "«"
            If .Visible And .RowCount > 0 Then Call .Sort(Array("Área", "Item"))
            
         End If
         If .RowCount > 0 Then
            On Error Resume Next
            .ColVisible("Pot.") = (pPos = 2)
            .ColVisible("Disp.") = (pPos = 2)
            .ColVisible("Freq.") = (pPos = 2)
            '.ColVisible("Tempo") = (pPos = 2)
         End If
      End With
      
      .GrpSessao.Width = .GrdSessao.Width + 5855
      .Width = .GrdSessao.Width + 5875

      .GrpSessao.Width = .GrdSessao.Width + 180
      .GrpValor.Width = .GrdSessao.Width
      .CmdNovoItem.Left = 240
      .LblVALOR.Left = .GrpValor.Width - .LblVALOR.Width - 240
      .LblTotal1.Left = .LblVALOR.Left - .LblTotal1.Width - 60
      .Width = .GrdSessao.Width + 5875
      .Left = (Screen.Width / 2) - (.Width / 2)
      
      .GrpProxSessao.Width = 2160
      .CmdProxSessao.Width = 2160
      .GrpProxSessao.Left = .GrpSessao.Left + .GrpSessao.Width - .GrpProxSessao.Width - 120
      .CmdProxSessao.Left = .GrpProxSessao.Width - .CmdProxSessao.Width
      .CmdProxSessao.Top = 230
      .LblLblProx.Left = .GrpProxSessao.Width - .LblLblProx.Width
      .LblProx.Left = .GrpProxSessao.Width - .LblProx.Width
   End With
End Sub
Private Sub ExibirVenda(pIDVENDA As Long)
   Dim TlVenda    As Object
   
   Set TlVenda = CriarObjeto("Financ3R.TL_Venda")
   With TlVenda
      Set .Sys = mvarSys
      
      .IDLOJA = mvarIDLOJA
      .IDCLIENTE = mvarIDCLIENTE
      .IDVENDA = pIDVENDA
      
      .Show vbModal
      
   End With
   Set TlVenda = Nothing
End Sub
