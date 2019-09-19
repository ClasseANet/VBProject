VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_CADRPONTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'option explicit
Private WithEvents mvarMe  As FrmCADRPONTO
Attribute mvarMe.VB_VarHelpID = -1
Private mvarSys            As Object
Private mvarTbRBANCOHH     As Object
Private mvarACUMULADO0     As Double
Private mvarDIAFOLGA       As Integer
Private mvarOk             As Boolean
Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
   Set mvarTbRBANCOHH = CriarObjeto("BANCO_3R.TB_RBANCOHH")
   Set mvarTbRBANCOHH.xdb = mvarSys.xdb
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Public Property Let Ok(ByVal vData As Boolean)
   mvarOk = vData
End Property
Public Property Get Ok() As Boolean
    Ok = mvarOk
End Property
Public Property Let ACUMULADO0(ByVal vData As Double)
   mvarACUMULADO0 = vData
End Property
Public Property Get ACUMULADO0() As Double
    ACUMULADO0 = mvarACUMULADO0
End Property
Public Property Let DIAFOLGA(ByVal vData As Integer)
   mvarDIAFOLGA = vData
End Property
Public Property Get DIAFOLGA() As Integer
    DIAFOLGA = mvarDIAFOLGA
End Property
Public Property Set TbRBANCOHH(ByVal vData As Object)
   Set mvarTbRBANCOHH = vData
   mvarTbRBANCOHH.xdb = mvarSys.xdb
End Property
Public Property Get TbRBANCOHH() As Object
    Set TbRBANCOHH = mvarTbRBANCOHH
End Property
Public Sub Show()
   mvarMe.Show vbModal
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmCADRPONTO
   Set mvarTbRBANCOHH = CriarObjeto("BANCO_3R.TB_RBANCOHH")
End Sub
Private Sub Class_Terminate()
   Set mvarMe = Nothing
   Set mvarTbRBANCOHH = Nothing
End Sub
Private Sub mvarMe_Activate()
   If GetTag(mvarMe, "1VEZ", 1) = 1 Then
      Call SetTag(mvarMe, "1VEZ", 0)
      Call PopulaTela
   End If
   mvarMe.MousePointer = vbDefault
   Screen.MousePointer = vbDefault
End Sub

Private Sub mvarMe_ChkFeriadoMouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim TLFeriado As TL_CADGFERIADO
   
   Set TLFeriado = New TL_CADGFERIADO
   With TLFeriado
      Set .Sys = mvarSys
      .Id = mvarMe.TxtDTPONTO.Text
      .Show
      Call PopulaTela
   End With
   Set TLFeriado = Nothing
End Sub

Private Sub mvarMe_ChkFLGFALTAClick()
   If mvarMe.ChkFLGFALTA.Value = xtpChecked Then
      mvarMe.TxtHHINI.Text = ""
      mvarMe.TxtHHFIM.Text = ""
   Else
      mvarMe.TxtHHINI.Text = Format(TbRBANCOHH.HHINI, "hh:mm")
      mvarMe.TxtHHFIM.Text = Format(TbRBANCOHH.HHFIM, "hh:mm")
   End If
   Call CalculaCampos
End Sub
Private Sub mvarMe_CmbSENTIDOKeyPress(KeyAscii As Integer)
   Call SendTab(mvarMe, KeyAscii)
End Sub
Private Sub mvarMe_CmbUNIDADEKeyPress(KeyAscii As Integer)
   Call SendTab(mvarMe, KeyAscii)
End Sub
Private Sub mvarMe_CmdAtend1Click()
   mvarMe.TxtHHFIM.Text = UltAtend
   Call mvarMe_TxtHHFIMLostFocus
End Sub
Public Function UltAtend() As String
   If IsDate(mvarMe.TxtHHATEND1.Text) Then
      UltAtend = Format(DateAdd("n", 10, CDate(mvarMe.TxtHHATEND1.Text)), "HH:MM")
   End If
End Function

Private Sub mvarMe_CmdExcluirClick()
   Dim bOk As Boolean
   
   If Excluir() Then
      Unload mvarMe
   End If
End Sub

Private Sub mvarMe_CmdSairClick()
   mvarOk = False
   Unload mvarMe
End Sub
Private Sub mvarMe_CmdSalvarClick()
   Dim bOk As Boolean
   bOk = True
   If mvarMe.ChkFLGFALTA.Value = xtpUnchecked And (mvarMe.TxtHHINI.Text = "" Or mvarMe.TxtHHFIM.Text = "") Then
      Call ExibirAviso("Batidas inválidas.")
      bOk = False
   End If
   
   If bOk Then
      mvarOk = Salvar(True)
      If mvarOk Then
         Unload mvarMe
      Else
         Call ExibirStop("Erro ao registrar ponto.", "Banco de Horas")
      End If
   End If
End Sub
Public Sub mvarMe_Load()
   If GetTag(mvarMe, "LIMPOU", 0) = 1 Then
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   Call LimpaTela
   Call MontarTela
   Call PopulaTela
   Call SetTag(mvarMe, "LIMPOU", 1)
End Sub
Private Sub mvarMe_TxtCHAPAGotFocus()
   'Call SelecionarTexto(Me.ActiveControl)
   Call LimpaTela
End Sub
Private Sub LimpaTela()
   Dim nIDLOJA As Integer
   DoEvents
   Call LimparTela(mvarMe)
   mvarMe.CmbIDABONO.ListIndex = -1
   mvarMe.TxtCHAPA.Tag = ""
   mvarMe.TxtOBS.Text = ""
   mvarMe.TxtHHBATIDA0.Text = ""
   mvarMe.TxtHHBATIDA1.Text = ""
   mvarMe.TxtHHATEND0.Text = ""
   mvarMe.TxtHHATEND1.Text = ""
   
   mvarMe.CmdSalvar.Enabled = True
   mvarMe.TxtCHAPA.Tag = ""
End Sub
Private Sub mvarMe_TxtCHAPAKeyPress(KeyAscii As Integer)
   Call SendTab(mvarMe, KeyAscii)
End Sub
Private Sub mvarMe_TxtHHABONOLostFocus()
   mvarMe.TxtHHABONO.Text = ValBr(mvarMe.TxtHHABONO.Text)
   Call CalculaCampos
End Sub
Private Sub mvarMe_TxtHHESPERADOLostFocus()
   mvarMe.TxtHHESPERADO.Text = ValBr(mvarMe.TxtHHESPERADO.Text)
   Call CalculaCampos
End Sub
Private Sub mvarMe_TxtHHFIMLostFocus()
   Dim sHHFim As String
   If mvarMe.TxtHHFIM.Text = "" And mvarMe.ChkFLGFALTA.Value = xtpUnchecked And mvarMe.ChkFLGDIA.Value = xtpChecked Then
      sHHFim = DateAdd("h", mvarSys.GetParam("HHESPERADO" & Weekday(TbRBANCOHH.DTPONTO), Default:=9) + 1, mvarSys.GetParam("WORKDAYSTARTTIME", Default:="09:00"))
      mvarMe.TxtHHFIM.Text = sHHFim  'Format(mvarSys.GetParam("WORKDAYENDTIME", pIDLOJA:=mvarSys.IDLOJA), "hh:mm")
   Else
      If mvarMe.TxtHHFIM.Text <> "" Then
         mvarMe.TxtHHFIM.Text = FormatarHora(mvarMe.TxtHHFIM.Text)
      End If
   End If
   If xVal(mvarMe.TxtHHREFEICAO.Text) = 0 Then
      Call mvarMe_TxtHHREFEICAOLostFocus
   End If
   If mvarMe.TxtHHINI.Text = "" Then
      Call mvarMe_TxtHHINILostFocus
   End If
   
   Call CalculaCampos
End Sub
Private Sub mvarMe_TxtHHINILostFocus()
   If mvarMe.TxtHHINI.Text = "" And mvarMe.ChkFLGFALTA.Value = xtpUnchecked And mvarMe.ChkFLGDIA.Value = xtpChecked Then
      mvarMe.TxtHHINI.Text = Format(mvarSys.GetParam("WORKDAYSTARTTIME", pIDLOJA:=mvarSys.IDLOJA), "hh:mm")
   Else
      If mvarMe.TxtHHINI.Text <> "" Then
         mvarMe.TxtHHINI.Text = FormatarHora(mvarMe.TxtHHINI.Text)
      End If
   End If
   Call CalculaCampos
End Sub
Private Sub mvarMe_TxtHHREFEICAOLostFocus()
   If xVal(mvarMe.TxtHHREFEICAO.Text) = 0 Then
      If IsDate(mvarMe.TxtHHFIM.Text) Then
         If CDate(mvarMe.TxtHHFIM.Text) >= CDate("13:00") Then
            mvarMe.TxtHHREFEICAO.Text = "1,00"
         End If
      End If
   End If
   mvarMe.TxtHHREFEICAO.Text = ValBr(mvarMe.TxtHHREFEICAO.Text)
   Call CalculaCampos
End Sub
Private Sub mvarMe_TxtSENHAKeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      Call mvarMe_CmdSalvarClick
   End If
   Call SendTab(mvarMe, KeyAscii)
End Sub
Private Sub MontarTela()
   Dim Sql As String
   Dim MyRs As Object
   Dim sDia As String
   
   With mvarMe
      Select Case Weekday(TbRBANCOHH.DTPONTO)
         Case 1: sDia = "Domingo"
         Case 2: sDia = "Segunda-Feira"
         Case 3: sDia = "Terça-Feira"
         Case 4: sDia = "Quarta-Feira"
         Case 5: sDia = "Quinta-Feira"
         Case 6: sDia = "Sexta-Feira"
         Case 7: sDia = "Sábado"
      End Select
      .Caption = " Registro de Ponto [" & TbRBANCOHH.DTPONTO & " - " & sDia & "]"
      
      Sql = "Select * " & vbNewLine
      Sql = "Select IDBATIDA [Id.], Convert(char(5), DTBATIDA, 108) [Hora]" & vbNewLine
      Sql = Sql & ", Case SENTIDO When 0 Then 'E' Else 'S' End [Sentido]" & vbNewLine
      Sql = Sql & ", Case FLGMANUAL When 0 Then '--' Else 'S' End [Manual] " & vbNewLine
      Sql = Sql & " From RBATIDA" & vbNewLine
      Sql = Sql & " Where IDLOJA = " & TbRBANCOHH.IDLOJA & vbNewLine
      Sql = Sql & " And IDFUNCIONARIO= " & TbRBANCOHH.IDFUNCIONARIO & vbNewLine
      Sql = Sql & " And Year(DTBATIDA)= " & Year(TbRBANCOHH.DTPONTO) & vbNewLine
      Sql = Sql & " And Month(DTBATIDA)= " & Month(TbRBANCOHH.DTPONTO) & vbNewLine
      Sql = Sql & " And Day(DTBATIDA)= " & Day(TbRBANCOHH.DTPONTO) & vbNewLine
      Sql = Sql & " Order by IDBATIDA"
      .GrdBatidas.Clear
      If mvarSys.xdb.AbreTabela(Sql, MyRs) Then
         .GrdBatidas.FillFromRS MyRs
         .GrdBatidas.AutoWidthCol 1
         .GrdBatidas.AutoWidthCol 2
         .GrdBatidas.AutoWidthCol 3
         .GrdBatidas.AutoWidthCol 4
         
         .TxtHHBATIDA0.Text = ""
         .TxtHHBATIDA1.Text = ""
         For i = 1 To .GrdBatidas.RowCount
            .TxtHHBATIDA0.Text = IIf((.TxtHHBATIDA0.Text <= .GrdBatidas.CellValue(i, "Hora") And Trim(.TxtHHBATIDA0.Text) <> ""), .TxtHHBATIDA0.Text, .GrdBatidas.CellValue(i, "Hora"))
            .TxtHHBATIDA1.Text = IIf((.TxtHHBATIDA1.Text >= .GrdBatidas.CellValue(i, "Hora") And Trim(.TxtHHBATIDA1.Text) <> ""), .TxtHHBATIDA1.Text, .GrdBatidas.CellValue(i, "Hora"))
         Next
      End If
            
      Sql = "Select A.IDATENDIMENTO [Id.]"
      Sql = Sql & ", Convert(char(5), A.HHINI, 108) [Inicio]"
      Sql = Sql & ", Convert(char(5), A.HHFIM, 108) [Fim]"
      Sql = Sql & ", F.NOME [Funcionáro] , C.NOME [Cliente]"
      Sql = Sql & " From OATENDIMENTO A " & vbNewLine
      Sql = Sql & " Join OCLIENTE C On A.IDLOJA=C.IDLOJA And A.IDCLIENTE=C.IDCLIENTE" & vbNewLine
      Sql = Sql & " Join RFUNCIONARIO F On A.IDLOJA=F.IDLOJA And A.IDFUNCIONARIO=F.IDFUNCIONARIO" & vbNewLine
      Sql = Sql & " Where A.IDLOJA = " & TbRBANCOHH.IDLOJA & vbNewLine
      Sql = Sql & " And Year(A.DTATEND)= " & Year(TbRBANCOHH.DTPONTO) & vbNewLine
      Sql = Sql & " And Month(A.DTATEND)= " & Month(TbRBANCOHH.DTPONTO) & vbNewLine
      Sql = Sql & " And Day(A.DTATEND)= " & Day(TbRBANCOHH.DTPONTO) & vbNewLine
      Sql = Sql & " Order by A.HHFIM"
      .GrdAtend.Clear
      If mvarSys.xdb.AbreTabela(Sql, MyRs) Then
         .GrdAtend.FillFromRS MyRs
         .GrdAtend.AutoWidthCol 1
         .GrdAtend.AutoWidthCol 2
         .GrdAtend.AutoWidthCol 3
         .GrdAtend.AutoWidthCol 4
         .GrdAtend.AutoWidthCol 5
         
         .TxtHHATEND0.Text = ""
         .TxtHHATEND1.Text = ""
         For i = 1 To .GrdAtend.RowCount
            .TxtHHATEND0.Text = IIf((.TxtHHATEND0.Text <= .GrdAtend.CellValue(i, "Inicio") And Trim(.TxtHHATEND0.Text) <> ""), .TxtHHATEND0.Text, .GrdAtend.CellValue(i, "Inicio"))
            .TxtHHATEND1.Text = IIf((.TxtHHATEND1.Text >= .GrdAtend.CellValue(i, "Fim") And Trim(.TxtHHATEND1.Text) <> ""), .TxtHHATEND1.Text, .GrdAtend.CellValue(i, "Fim"))
         Next
      End If
   
   End With
   Set MyRs = Nothing
End Sub
Private Sub PopulaTela()
   Dim Sql As String
   
   mvarMe.TxtIDMOVHH.Text = IIf(TbRBANCOHH.IDMOVHH = 0, "", TbRBANCOHH.IDMOVHH)
   mvarMe.TxtCHAPA.Tag = TbRBANCOHH.IDFUNCIONARIO
   
   Sql = "Select CHAPA, NOME"
   Sql = Sql & " From RFUNCIONARIO"
   Sql = Sql & " Where IDLOJA=" & mvarSys.IDLOJA
   Sql = Sql & " And IDFUNCIONARIO=" & TbRBANCOHH.IDFUNCIONARIO
   If mvarSys.xdb.AbreTabela(Sql) Then
      mvarMe.TxtCHAPA.Text = mvarSys.xdb.RsAux("CHAPA")
      mvarMe.TxtNOME.Text = mvarSys.xdb.RsAux("NOME")
   End If
   mvarMe.ChkFeriado.Value = IIf(eFeriado(mvarSys.xdb, TbRBANCOHH.DTPONTO), xtpChecked, xtpUnchecked)
   mvarMe.TxtDTPONTO.Text = TbRBANCOHH.DTPONTO
   mvarMe.TxtHHINI.Text = Format(TbRBANCOHH.HHINI, "hh:mm")
   mvarMe.TxtHHFIM.Text = Format(TbRBANCOHH.HHFIM, "hh:mm")
   mvarMe.TxtOBS.Text = TbRBANCOHH.OBS
   
   If TbRBANCOHH.IDMOVHH = 0 Then
      If eFeriado(mvarSys.xdb, TbRBANCOHH.DTPONTO) Or Weekday(TbRBANCOHH.DTPONTO) = 1 Then
         mvarMe.ChkFLGDIA.Value = 0
         mvarMe.ChkFLGDIA.Enabled = False
      Else
         mvarMe.ChkFLGDIA.Value = 1
      End If
      mvarMe.ChkFLGFALTA.Value = 0
      mvarMe.TxtHHESPERADO.Text = mvarSys.GetParam("HHESPERADO" & Weekday(TbRBANCOHH.DTPONTO), Default:=0)
      If eFeriado(mvarSys.xdb, TbRBANCOHH.DTPONTO) Or Weekday(TbRBANCOHH.DTPONTO) = 1 Or Weekday(TbRBANCOHH.DTPONTO) = mvarDIAFOLGA Then
         mvarMe.TxtHHESPERADO.Text = 0
      End If
   Else
      mvarMe.ChkFLGFALTA.Value = TbRBANCOHH.FLGFALTA
      mvarMe.ChkFLGDIA.Value = TbRBANCOHH.FLGDIA
      mvarMe.TxtHHESPERADO.Text = TbRBANCOHH.HHESPERADO
      mvarMe.TxtHHTRAB.Text = TbRBANCOHH.HHTrab
   End If
   mvarMe.TxtHHESPERADO.Text = ValBr(xVal(mvarMe.TxtHHESPERADO.Text))
   TbRBANCOHH.HHREFEICAO = CalculaRefeicao(mvarMe.TxtHHINI.Text, mvarMe.TxtHHFIM.Text)
   mvarMe.TxtHHREFEICAO.Text = ValBr(TbRBANCOHH.HHREFEICAO)
   mvarMe.TxtHHTRAB.Text = ValBr(xVal(mvarMe.TxtHHTRAB.Text))
         
   'If mvarACUMULADO0 = 0 Then
   '   mvarACUMULADO0 = TbRBANCOHH.ACUMULADO - TbRBANCOHH.SALDODIA
   'End If
   
   mvarMe.TxtAcumulado0.Text = ValBr(mvarACUMULADO0)
   mvarMe.TxtSALDODIA.Text = ValBr(TbRBANCOHH.SALDODIA)
   If TbRBANCOHH.IDMOVHH = 0 Then
      mvarMe.TxtACUMULADO.Text = mvarACUMULADO0 + ValBr(mvarMe.TxtSALDODIA.Text)
   Else
      mvarMe.TxtACUMULADO.Text = ValBr(TbRBANCOHH.ACUMULADO)
   End If

   ''mvarSaldoParcial = xVal(mvarMe.TxtSaldoParcial.Text)
   ''TbRBANCOHH.IDABONO = mvarMe.CmbIDABONO.ItemData(mvarMe.CmbIDABONO.ListIndex)
   ''TbRBANCOHH.HHABONADO = xVal(mvarMe.TxtHHABONO.Text)
   ''FLGZERASALDO


'   mvarMe.TxtHHATEND0
'   mvarMe.TxtHHATEND1
'   mvarMe.TxtHHBATIDA0
'   mvarMe.TxtHHBATIDA1
   Call CalculaCampos
End Sub
Private Sub HabilitaAbono()
   
   If xVal(mvarMe.TxtSaldoParcial.Text) >= 0 Then
      mvarMe.LblIDABONO.Enabled = False
      mvarMe.LblHHABONO.Enabled = False
      mvarMe.CmbIDABONO.Enabled = False
      mvarMe.TxtHHABONO.Enabled = False
      mvarMe.CmbIDABONO.ListIndex = -1
      mvarMe.TxtHHABONO.Text = ""
   Else
      mvarMe.LblIDABONO.Enabled = True
      mvarMe.LblHHABONO.Enabled = True
      mvarMe.CmbIDABONO.Enabled = True
      mvarMe.TxtHHABONO.Enabled = True
      Call LocalizarCombo(mvarMe.CmbIDABONO, "IDABONO", True, True)
      mvarMe.TxtHHABONO.Text = ValBr(TbRBANCOHH.HHABONADO)
   End If
End Sub
Private Sub PopulaClasse()
   Dim Sql As String
            
   TbRBANCOHH.IDMOVHH = xVal(mvarMe.TxtIDMOVHH.Text)
   TbRBANCOHH.IDFUNCIONARIO = xVal(mvarMe.TxtCHAPA.Tag)
   TbRBANCOHH.DTPONTO = mvarMe.TxtDTPONTO.Text
   TbRBANCOHH.HHINI = mvarMe.TxtDTPONTO.Text & " " & mvarMe.TxtHHINI.Text
   TbRBANCOHH.HHFIM = mvarMe.TxtDTPONTO.Text & " " & mvarMe.TxtHHFIM.Text
   TbRBANCOHH.OBS = mvarMe.TxtOBS.Text
   TbRBANCOHH.FLGDIA = mvarMe.ChkFLGDIA.Value
   TbRBANCOHH.FLGFALTA = mvarMe.ChkFLGFALTA.Value
   TbRBANCOHH.HHREFEICAO = xVal(mvarMe.TxtHHREFEICAO.Text)
   TbRBANCOHH.HHESPERADO = xVal(mvarMe.TxtHHESPERADO.Text)
   TbRBANCOHH.HHTrab = xVal(mvarMe.TxtHHTRAB.Text)
   mvarSaldoParcial = xVal(mvarMe.TxtSaldoParcial.Text)
   TbRBANCOHH.IDABONO = mvarMe.CmbIDABONO.ItemData(mvarMe.CmbIDABONO.ListIndex)
   TbRBANCOHH.HHABONADO = xVal(mvarMe.TxtHHABONO.Text)
   TbRBANCOHH.SALDODIA = xVal(mvarMe.TxtSALDODIA.Text)
   TbRBANCOHH.ACUMULADO = xVal(mvarMe.TxtACUMULADO.Text)
   TbRBANCOHH.FLGZERASALDO = (0 = xVal(mvarMe.TxtACUMULADO.Text))
End Sub
Public Function Excluir() As Boolean
   Dim sMsg As String
   mvarOk = False
   With TbRBANCOHH
      If .IDLOJA <> 0 And .IDFUNCIONARIO <> 0 And .IDMOVHH <> 0 Then
         sMsg = "Funcionario: " & mvarMe.TxtNOME.Text & vbNewLine
         sMsg = sMsg & "Data: " & mvarMe.TxtDTPONTO.Text & vbNewLine & vbNewLine
         If ExibirPergunta(sMsg & "Deseja excluir este ponto?", "Ponto") = vbYes Then
            If .Excluir(True) Then
               Excluir = True
               mvarOk = True
               Call ExibirResultado(mvarSys, True, 1)
            Else
               Call ExibirStop("Erro ao excluir ponto.", "Banco de Horas")
            End If
         End If
      End If
   End With
End Function
Public Function Salvar(Optional bAPROVADO = False) As Boolean
   Dim bOk     As Boolean
   Dim Queries As Collection
   Dim nValorA As Currency
   
   If Not mvarMe.Visible Then
      Call LimpaTela
      Call MontarTela
      Call PopulaTela
   End If
   Set Queries = New Collection
   If ValidaCampos() Then
      nValorA = xVal(mvarMe.TxtACUMULADO.Text) - TbRBANCOHH.ACUMULADO
      If nValorA <> 0 Then
         Sql = ""
         Sql = Sql & "Update RBANCOHH" & vbNewLine
         Sql = Sql & " Set ACUMULADO = ACUMULADO + (" & SqlNum(nValorA) & ")" & vbNewLine
         Sql = Sql & " Where IDLOJA =" & mvarSys.IDLOJA & vbNewLine
         Sql = Sql & " And IDFUNCIONARIO = " & SqlNum(TbRBANCOHH.IDFUNCIONARIO) & vbNewLine
         Sql = Sql & " And DTPONTO > " & SqlDate(mvarMe.TxtDTPONTO.Text) & vbNewLine
         Queries.Add Sql
      End If
      Call PopulaClasse
      TbRBANCOHH.APROVADO = IIf(bAPROVADO, 1, 0)
      If TbRBANCOHH.isDirt Then
         Queries.Add TbRBANCOHH.QrySave
        
         bOk = mvarSys.xdb.Executa(Queries) 'TbRBANCOHH.Salvar()
         If bOk Then
            Call TbRBANCOHH.Popula(mvarSys.xdb.ADORs)
         End If
      Else
         bOk = True
      End If
      mvarMe.CmdSalvar.Enabled = Not bOk
   End If
   Set Queries = Nothing
   
   DoEvents
   Call ExibirResultado(mvarSys, bOk, 1)
   
   Salvar = bOk
   
   If Not mvarMe.Visible Then mvarOk = bOk
End Function
Private Function ValidaCampos() As Boolean
   ValidaCampos = True
End Function
Private Sub CalculaCampos()
   
   mvarMe.TxtHHTRAB.Text = ValBr(CalculaHHTrab(mvarMe.TxtHHINI.Text, mvarMe.TxtHHFIM.Text, xVal(mvarMe.TxtHHREFEICAO.Text)))
   mvarMe.TxtSALDODIA.Text = ValBr(CalculaSaldoDia(mvarMe.TxtHHINI.Text, mvarMe.TxtHHFIM.Text, xVal(mvarMe.TxtHHREFEICAO.Text), xVal(mvarMe.TxtHHESPERADO.Text), xVal(mvarMe.TxtHHABONO.Text)))
   
   If mvarMe.TxtHHINI.Text <> "" And mvarMe.TxtHHFIM.Text <> "" Then
      mvarMe.TxtSaldoParcial.Text = ValBr(xVal(mvarMe.TxtHHTRAB.Text) - xVal(mvarMe.TxtHHESPERADO.Text))
   Else
      mvarMe.TxtSaldoParcial.Text = ValBr(0)
   End If
   mvarMe.TxtACUMULADO.Text = ValBr(mvarACUMULADO0 + ValBr(mvarMe.TxtSALDODIA.Text))
   
   
   Call HabilitaAbono
End Sub
Public Function CalculaHHTrab(pHHINI As String, pHHFIM As String, pHHRefeicao As Double) As Double
   Dim HHTrab As Double
   
   If Not IsDate(pHHINI) Then pHHINI = "0:00"
   If Not IsDate(pHHFIM) Then pHHFIM = "0:00"
   If DateDiff("h", pHHINI, pHHFIM) <= 0 Then
      CalculaHHTrab = DateDiff("n", pHHINI, pHHFIM)
   Else
      CalculaHHTrab = DateDiff("n", pHHINI, pHHFIM) - (pHHRefeicao * 60)
   End If
   CalculaHHTrab = xVal(CalculaHHTrab / 60, 2)
End Function
Public Function CalculaSaldoDia(pHHINI As String, pHHFIM As String, pHHRefeicao As Double, pHHEsperado As Double, pHHAbono As Double) As Double
   Dim HHTrab As Double
   Dim HHSaldoParcial As Double
   
   HHTrab = CalculaHHTrab(pHHINI, pHHFIM, pHHRefeicao)
   HHSaldoParcial = HHTrab - pHHEsperado
   
   CalculaSaldoDia = xVal(HHSaldoParcial + pHHAbono, 2)
End Function
Public Function CalculaRefeicao(pHHINI As String, pHHFIM As String) As Double
   If (Trim(pHHINI) = "" And Trim(pHHFIM) = "") Then
      CalculaRefeicao = 0
   Else
      If CDate(Trim(pHHFIM)) < CDate("12:30") Then
         CalculaRefeicao = 0
      Else
         CalculaRefeicao = 1
      End If
   End If
End Function
