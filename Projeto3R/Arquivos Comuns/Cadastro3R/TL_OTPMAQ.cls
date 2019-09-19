VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_CADOTPMAQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe  As FrmCADOTPMAQ
Attribute mvarMe.VB_VarHelpID = -1

Private mvarShortBar As TL_ShortBar
'Private WithEvents mvarMDI As FrmMDI

Private mvarRptRec        As ReportRow
Private TbCad             As Object 'TB_OTPMAQUINA

Private mvarIDTPMAQ        As Integer
Private mvarIDLOJA         As Integer
Private mvarSys            As Object
Public Property Set TLCadastro(ByRef vData As Object)
   Dim Sql As String
   Dim sTag As String
   
   With vData
      .Titulo = "Tipo de Máquinas"
      
      Sql = "Select IDTPMAQ [Id.], DSCMAQ [Descrição]"
      'Sql = Sql & ", [Ordem Num.]=Case TPNUM When 1 Then 'Crescente' Else 'Decrescente' End"
      'Sql = Sql & ", NUMTRATAMENTO"
      Sql = Sql & " From OTPMAQ"
      Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
      .Consulta = Sql
      
      sTag = ""
      sTag = sTag & "|CAMPO=Id.|CHAVE=S|GRUPO=0|VISIBLE=S:"
      sTag = sTag & "|CAMPO=Descrição|CHAVE=N|GRUPO=0|VISIBLE=S:"
      .TAGCAMPOS = sTag
   End With
End Property
Public Property Set ShortBar(ByVal vData As TL_ShortBar)
   Set mvarShortBar = vData
End Property
Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
   mvarIDLOJA = mvarSys.IDLOJA
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Public Property Let IDTPMAQ(ByVal vData As String)
   mvarIDTPMAQ = vData
End Property
Public Property Get IDTPMAQ() As String
   IDTPMAQ = mvarIDTPMAQ
End Property
Public Property Set RptRec(ByVal vData As ReportRow)
   Set mvarRptRec = vData
'   mvarIDLOJA = RptRec.Record.Item(0).Value
   mvarIDTPMAQ = RptRec.Record.Item(0).Value
End Property
Public Property Get RptRec() As ReportRow
   Set RptRec = mvarRptRec
End Property
Public Sub Show()
   mvarMe.Show vbModal
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmCADOTPMAQ
End Sub
Private Sub mvarMe_Activate()
   If GetTag(mvarMe, "1VEZ", 1) = 1 Then
      Call SetTag(mvarMe, "1VEZ", 0)
      'Call MontaTela
      Call PopulaTela
   End If
End Sub
Private Sub InstanciaClasse()
   Set TbCad = Nothing
   Set TbCad = CriarObjeto("BANCO_3R.TB_OTPMAQ")
   Set TbCad.Xdb = mvarSys.Xdb
End Sub
Private Sub PopulaTela()
   If TbCad Is Nothing Then
      Call InstanciaClasse
   End If
   
   With TbCad
      If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDTPMAQ:=mvarIDTPMAQ) Then
         mvarMe.CmdExcluir.Enabled = True
         
         mvarMe.TxtID.Text = .IDTPMAQ
         mvarMe.TxtDSCMAQ.Text = .DSCMAQ
         mvarMe.OptNumeracao(.TPNUM).Value = True
         mvarMe.TxtNUMTRATAMENTO.Text = .NUMTRATAMENTO
         mvarMe.ChkTPMANIPULO.Value = .TPMANIPULO
         
         Call PopulaGrid
      Else
         Call LimpaTela
      End If
   End With
End Sub
Private Sub PopulaGrid()
   Dim sTag As String
   Dim i As Integer
   Dim sNum As String
   Dim nRow As Integer
   
   sTag = TbCad.NUMTRATAMENTO
   With mvarMe.GrdTrat
      For i = 1 To xVal(GetTag(sTag, "GR", 0))
         sNum = GetTag(sTag, "GR" & i, "")
         While sNum <> ""
            If InStr(sNum, ",") = 0 Then
               nRow = xVal(sNum)
               sNum = ""
            Else
               nRow = xVal(Mid(sNum, 1, InStr(sNum, ",")))
               sNum = Mid(sNum, InStr(sNum, ",") + 1)
            End If
            If nRow < mvarMe.GrdTrat.RowCount Then
               mvarMe.GrdTrat.CellValue(nRow, "Grupo") = i
            End If
         Wend
      Next
   End With
End Sub
Public Function Excluir() As Boolean
   If mvarIDTPMAQ <> 0 Then
      If VerificarDependencias() Then
         If ExibirPergunta("Confirma a exclusão do tipo de máquina?", "Atenção", False) = vbYes Then
            Call TbCad.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDTPMAQ:=xVal(mvarMe.TxtID))
            If TbCad.Excluir() Then
               Call ExibirResultado(mvarSys, True)
               
               'mvarIsDirt = True
               
               Call LimpaTela
               If Not mvarShortBar Is Nothing Then
                  mvarShortBar.TLCadastro.F_Refresh
               End If
               If mvarMe.TxtDSCMAQ.Visible And mvarMe.TxtDSCMAQ.Enabled Then
                  mvarMe.TxtDSCMAQ.SetFocus
               End If
               Excluir = True
            Else
               Call ExibirResultado(mvarSys, False)
            End If
         End If
      End If
   End If
End Function
Private Function ValidaCampos() As Boolean
   ' *******************
   Dim bResult As Boolean
   
   bResult = True
   If Trim(mvarMe.TxtDSCMAQ.Text) = "" Then
      ExibirAviso "A Descrição deve ser preenchida."
      bResult = False
      mvarMe.TxtDSCMAQ.SetFocus
      GoTo Fim
   End If
Fim:
   ValidaCampos = bResult
End Function
Private Sub LimpaTela()
   Dim n As Object
   On Error Resume Next
   With mvarMe
      For Each n In mvarMe.Controls
         n.Text = ""
      Next
      .GrdTrat.Clear
   End With
End Sub
Private Sub MontaTela()
   Dim TbTrat As Object
   Dim i As Integer
   
   mvarMe.ChkTPMANIPULO.Value = xtpChecked
   mvarMe.OptNumeracao(1).Value = True
   
   Set TbTrat = CriarObjeto("BANCO_3R.TB_OTPTRATAMENTO")
   Set TbTrat.Xdb = mvarSys.Xdb
   If TbTrat.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="IDTPTRATAMENTO<>0") Then
      With mvarMe.GrdTrat
         .ShowControlsInAllCells = True
         .Header.DragCols = False
        
         With .Combos
            With .Add("GRUPO")
               .AddItem " ", 0
               .AddItem "Grupo 01", 1
               .AddItem "Grupo 02", 2
               .AddItem "Grupo 03", 3
               .AddItem "Grupo 04", 4
               
               .AutoAdjustWidth
            End With
         End With
         
         Call .FillFromRS(TbTrat.Rs)
         For i = 1 To .ColCount
            .ColVisible(i) = False
         Next
         .ColVisible("DSCTRATAMENTO") = True
         .ColHeaderText("DSCTRATAMENTO") = "Tratamento"
         .ColWidth("DSCTRATAMENTO") = 120
         
         With .AddCol("GRUPO", sHeader:="Grupo", lWidth:=50, bVisible:=True)
            .eType = igCellCombo
            .eTypeFlags = igComboBtnFlat
            .sCtrlKey = "GRUPO"
         End With
         .ColWidth("GRUPO") = 80
         For i = 1 To .RowCount
            .CellType(i, "GRUPO") = igCellCombo
            .CellTypeFlags(i, "GRUPO") = igComboBtnFlat
            .CellCtrlKey(i, "GRUPO") = "GRUPO"
         Next
      End With
   End If
End Sub
Private Sub PopulaClasse(ByRef Queries As Collection)
   Dim i As Integer
   Dim sTrat As String
   Dim nGrp  As Integer
   Dim sAux  As String
   
   If TbCad Is Nothing Then Call InstanciaClasse
   
   With TbCad
      If .IDTPMAQ = 0 Then
'         .DTCADASTRO = mvarSys.Xdb.SysDate()
      End If
      .IDLOJA = mvarIDLOJA
      .DSCMAQ = mvarMe.TxtDSCMAQ.Text
      .TPNUM = IIf(mvarMe.OptNumeracao(1).Value, 1, 2)
      sTrat = ""
      For i = 1 To mvarMe.GrdTrat.RowCount
         nGrp = xVal(mvarMe.GrdTrat.CellValue(i, "Grupo"))
         If nGrp <> 0 Then
            If GetTag(sTrat, "GR" & nGrp, "") = "" Then
               sTrat = SetTag(sTrat, "GR", xVal(GetTag(sTrat, "GR", 0)) + 1)
            End If
            sAux = GetTag(sTrat, "GR" & nGrp, "")
            If InStr(sAux, mvarMe.GrdTrat.CellValue(i, "IDTPTRATAMENTO")) = 0 Then
               sAux = sAux & IIf(Trim(sAux) = "", "", ",") & mvarMe.GrdTrat.CellValue(i, "IDTPTRATAMENTO")
            End If
            sTrat = SetTag(sTrat, "GR" & nGrp, sAux)
         End If
      Next
      .NUMTRATAMENTO = sTrat
     
      If .isDirt Then
         If .Salvar Then
            mvarIDTPMAQ = .IDTPMAQ
         End If
      End If
   End With
End Sub
Private Function Salvar() As Boolean
   Dim bOk     As Boolean
   Dim xNode   As TreeViewNode
   Dim bIsDirt As Boolean
   Dim Queries As Collection
   
   If ValidaCampos() Then
      Set Queries = New Collection
      Call PopulaClasse(Queries)
      
      'bIsDirt = TbCad.isDirt
      'If bIsDirt Then
      '   If TbCad.Salvar() Then
      '      bOk = True
      '   End If
      'Else
      '   bOk = True
      'End If
      bOk = False
      If Queries.Count = 0 Then
         bOk = True
      Else
         If mvarSys.Xdb.Executa(Queries) Then
            bOk = True
         End If
      End If
      
      Call ExibirResultado(mvarSys, bOk, IIf(bOk, 1, 2))
   End If
   Salvar = bOk
End Function
Private Function VerificarDependencias()
   Dim bResult    As Boolean
   Dim TbMaq    As Object
   Dim sMsg       As String
   
   Set TbMaq = CriarObjeto("BANCO_3R.TB_OMAQUINA")
   With TbMaq
      Set .Xdb = mvarSys.Xdb
      If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:="IDTPMAQ=" & SqlNum(mvarMe.TxtID)) Then
         sMsg = "Este item já possui máquina associada e não pode ser excluído."
         Call ExibirAviso(sMsg)
         GoTo Saida
      End If
   End With
   
   bResult = True
   
   VerificarDependencias = bResult
Saida:
   Set TbMaq = Nothing
End Function

Private Sub mvarMe_CmdExcluirClick()
   If mvarIDTPMAQ <> 0 Then
      If mvarMe.Visible Then
         Call PopulaTela
      End If
      Call Excluir
   End If
End Sub

Private Sub mvarMe_CmdLovClick()
'   Dim SQL As String
'   Dim xAct As XActive.XLOV
'
'   Set xAct = New XActive.XLOV
'
'
'   SQL = "Select IDTPMAQ [Id], DSCMAQ [Nome], TPNUM, NUMTRATAMENTO"
'   SQL = SQL & " From OTPMAQ"
'   With xAct
'      .dBase = mvarSys.XDb
'      .Query = SQL
'      .CAMPO_ID = "IDTPMAQ"
'      .CAMPO_NOME = "DSCMAQ"
'      .Show
'   End With
'   Call F_LOV(mvarSys.XDb, "OTPMAQ", SQL)
End Sub

Private Sub mvarMe_CmdNovoClick()
   If xVal(mvarMe.TxtID) = 0 And Trim(mvarMe.TxtDSCMAQ) = "" Then
      Call LimpaTela
      mvarMe.TxtDSCMAQ.SetFocus
   Else
      If Salvar Then
         Call LimpaTela
         Call InstanciaClasse
         mvarMe.TxtDSCMAQ.SetFocus
      Else
         mvarMe.TxtID.SetFocus
      End If
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
End Sub


