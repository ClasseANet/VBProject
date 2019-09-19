VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_FecharDia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim nTotalErros   As Integer
Dim nTotalAvisos  As Integer
Dim bExeFechamento As Boolean
Private Enum Bandeira
    Verde = 1
    Vermelha = 2
    Amarela = 3
    Ok = 4
    Erro = 5
    Aviso = 6
End Enum
Private WithEvents mvarMe  As FrmFecharDia
Attribute mvarMe.VB_VarHelpID = -1

Private mvarSys         As Object
Private mvarCalControl  As Object
Private mvarNGCal       As Object
Private mvarIDLOJA   As Integer
Public Property Set Sys(ByVal vData As Object)
    Set mvarSys = vData
    Set gSys = mvarSys
    mvarIDLOJA = mvarSys.Propriedades("IDLOJA")
gDebug = (ReadIniFile(mvarSys.ExePath & "Debug.txt", "DEBUG", "UTILITARIO", "0") = "1")
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
Public Property Set CalControl(ByVal vData As Object)
    Set mvarCalControl = vData
End Property
Public Property Get CalControl() As Object
    Set CalControl = mvarCalControl
End Property
Public Property Set NgCal(ByVal vData As Object)
    Set mvarNGCal = vData
End Property
Public Property Get NgCal() As Object
    Set NgCal = mvarNGCal
End Property
Public Sub Show()
   On Error Resume Next
   If mvarSys.Propriedades("FECHATELAAG") = 1 Then
      Call CapturarTelaSis(Environ("TEMP") & "\" & "AG" & Format(mvarSys.xDb.SysDate, "ddmmyy") & ".jpg")
      If Not ExisteArquivo(Environ("TEMP") & "\" & "AG" & Format(mvarSys.xDb.SysDate, "ddmmyy") & ".jpg") Then
         Call CapturarTelaSis(Environ("TEMP") & "\" & "AG" & Format(mvarSys.xDb.SysDate, "ddmmyy") & ".jpg")
      End If
   End If
   '* Carregar Form
   mvarMe.Show vbModal
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmFecharDia
End Sub
Private Sub mvarMe_Activate()
If gDebug Then MsgBox "Activate"
   If GetTag(mvarMe, "1VEZ", 1) = 1 Then
      Call SetTag(mvarMe, "1VEZ", 0)
      'Call VerificarDia(mvarMe.CmbDTFECHAR.Value, False)
   End If
End Sub
Private Sub mvarMe_ChkExibirClick()
   Call ExibirItens
End Sub
Private Sub mvarMe_CmdFecharDia()
   Unload mvarMe
End Sub
Private Sub mvarMe_CmbDTFECHARChange()
   mvarMe.MvwDTFECHAR.Value = mvarMe.CmbDTFECHAR.Value
   mvarMe.wndDatePicker.Select mvarMe.CmbDTFECHAR.Value
   
   Call LimparGrid
   Call ExibeValoresDia(mvarMe.CmbDTFECHAR.Value)
End Sub
Private Sub mvarMe_CmdFecharDiaClick()
   Call VerificarDia(mvarMe.CmbDTFECHAR.Value)
   Call FecharDia(mvarMe.CmbDTFECHAR.Value)
End Sub
Private Sub mvarMe_CmdSairClick()
   If mvarMe.Timer.Enabled Then
      mvarMe.Timer.Enabled = False
      If vbYes = ExibirPergunta("O Sitema está tentando enviar fechamento aos sócios." & vbNewLine & vbNewLine & "Tem certeza que deseja cancelar?", pDefaultYes:=False) Then
         Unload mvarMe
      Else
         mvarMe.Timer.Enabled = True
      End If
   Else
      Unload mvarMe
   End If
End Sub
Private Sub mvarMe_CmdVerificarClick()
   Call VerificarDia(mvarMe.CmbDTFECHAR.Value)
End Sub
Private Sub mvarMe_GrdFecharDiaHyperlinkClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal HyperlinkIndex As Long)
   Dim nCod As Long
   
   nCod = xVal(Mid(Item.Value, Item.Hyperlinks.Item(0).HyperTextBegin + 1, Item.Hyperlinks.Item(0).HyperTextLen))
   
   Select Case GetTag(Item, "TABELA", "")
      Case "OEVENTOAGENDA": Call ExibirEvento(nCod)
      Case "OATENDIMENTO": Call ExibirAtendimento(nCod)
      Case "CVENDA": Call ExibirVenda(nCod)
   End Select
   'Call FecharDia(mvarMe.CmbDTFECHAR.Value)
End Sub
Private Sub ExibirEvento(pCod As Long)
   If mvarNGCal Is Nothing Then
      Set mvarNGCal = CriarObjeto("Calendario3R.NG_Calendario")
      Set mvarNGCal.Sys = mvarSys
   End If
   Call mvarNGCal.ExibirEvento(pCod, mvarCalControl)
End Sub
Private Sub ExibirAtendimento(pCod As Long)
   If mvarNGCal Is Nothing Then
      Set mvarNGCal = CriarObjeto("Calendario3R.NG_Calendario")
      Set mvarNGCal.Sys = mvarSys
   End If
   Call mvarNGCal.ExibirAtendimento(pCod, mvarCalControl)
End Sub
Private Sub ExibirVenda(pCod As Long)
   Dim NgFinanc As Object
   
   Set NgFinanc = CriarObjeto("Financ3R.NG_Financeiro")
   Set NgFinanc.Sys = mvarSys
   Call NgFinanc.ExibirVenda(pCod, mvarCalControl)
End Sub
Private Sub mvarMe_Load()
   Dim sDT  As String
If gDebug Then MsgBox "Load"
   mvarMe.LblVLTotCT.Caption = ""
   mvarMe.LblVLDin.Caption = ""
   mvarMe.LblVLChe.Caption = ""
   mvarMe.LbVLDia.Caption = ""
   mvarMe.LblVLCred.Caption = ""
   mvarMe.LblVLDeb.Caption = ""
   mvarMe.LblVLCaixa.Caption = ""
      
   If Not mvarCalControl Is Nothing Then sDT = Format(mvarCalControl.DayView.Selection.Begin, "dd/mm/yyyy")
   If Not IsDate(sDT) Then sDT = mvarSys.xDb.SysDate(1)
   mvarMe.CmbDTFECHAR.Value = sDT
   mvarMe.MvwDTFECHAR.Value = sDT
   mvarMe.wndDatePicker.Select sDT
   
   If Format(sDT, "dd/mm/yyyy") = Format(mvarSys.xDb.SysDate(), "dd/mm/yyyy") Then
      Call VerificarDia(sDT, False)
   Else
      mvarMe.GrdFecharDia.Enabled = False
   End If
End Sub
Private Sub VerificarDia(pData As String, Optional pGravaHtml As Boolean = True)
If gDebug Then MsgBox "VerificarDia"
   Dim sDate As String
   Dim i As Integer
   
   Screen.MousePointer = vbHourglass
   mvarMe.MousePointer = vbHourglass
   
   sDate = Format(pData, "dd/mm/yyyy")
   
   Call LimparGrid
   Call ExibeValoresDia(mvarMe.CmbDTFECHAR.Value)
   mvarMe.GrdFecharDia.Enabled = True
   
   mvarMe.ProgressBar1.Scrolling = xtpProgressBarMarquee
   mvarMe.ProgressBar1.Enabled = True
   mvarMe.ProgressBar1.Visible = True
   DoEvents
   
   nTotalErros = 0
   nTotalAvisos = 0
   
   Call ResizeGrid
   Call VerificarEventos(sDate)
   Call VerificarAtendimentos(sDate)
   Call VerificarVendas(sDate)
   Call VerificarTarefas(sDate)
   
   For i = (mvarMe.GrdFecharDia.Rows.Count - 1) To 0 Step -1
      If mvarMe.GrdFecharDia.Rows(i).GroupRow Then
         If GetTag(mvarMe.GrdFecharDia, "Erro" & i) <> "" Then
            mvarMe.GrdFecharDia.Rows(i).Expanded = GetTag(mvarMe.GrdFecharDia, "Erro" & i)
         End If
      End If
   Next
   
   Call ResizeGrid
   
   mvarMe.ProgressBar1.Value = 0
   mvarMe.ProgressBar1.Scrolling = xtpProgressBarSmooth
   mvarMe.ProgressBar1.Enabled = False
   mvarMe.ProgressBar1.Visible = False
   
   'If pGravaHtml Then
   '   Call EscreverFechamento(pData)
   'End If
  
   Screen.MousePointer = vbDefault
   mvarMe.MousePointer = vbDefault
End Sub
Private Sub ExibeValoresDia(pDataIni As String, Optional pDataFim As String)
   
'   Exit Sub
   
   Dim Sql As String
   Dim sAux As String
        
   '************
   '* Valores do Dia
   Dim sVLDINDIA As String
   Dim sVLDEBDIA As String
   Dim sVLCREDIA As String
   Dim sVLCHEDIA As String
   Dim sVLCAIXADIA As String
   Dim nVLCartao As Currency
   Dim oMovCC As Object
   Dim sData As String
           
   If pDataFim = pDataIni Or pDataFim = "" Then
      mvarMe.GrpValoresDia.Caption = " Valores do Dia " & pDataIni
   Else
      mvarMe.GrpValoresDia.Caption = " Valores do Período " & pDataIni & " à " & pDataFim
   End If
   pDataFim = IIf(pDataFim = "", pDataIni, pDataFim)
   pDataIni = pDataIni & " 00:00"
   pDataFim = pDataFim & " 23:59"
   
   mvarMe.GrpValoresDia.Font.Bold = True
   mvarMe.GrpValoresDia.ForeColor = vbBlack
           
   mvarMe.LblVLTotCT.Caption = ""
   mvarMe.LblVLDin.Caption = ""
   mvarMe.LblVLChe.Caption = ""
   mvarMe.LbVLDia.Caption = ""
   mvarMe.LblVLCred.Caption = ""
   mvarMe.LblVLDeb.Caption = ""
   If xVal(Replace(mvarMe.LblVLCaixa.Caption, "R$", "")) = 0 Then
      mvarMe.LblVLCaixa.Caption = ""
   End If
   
   sVLDINDIA = "R$ 0,00"
   sVLDEBDIA = "R$ 0,00"
   sVLCREDIA = "R$ 0,00"
   sVLCHEDIA = "R$ 0,00"
   nVLCartao = 0
   
   Sql = "Select P.IDFORMAPGTO, SUM(P.VLPGTO)-SUM(V.VLTROCO) [Valor]" & vbNewLine
   Sql = Sql & " From CPGTOSVENDA P "
   Sql = Sql & " Join CVENDA V On P.IDLOJA=V.IDLOJA And P.IDVENDA=V.IDVENDA" & vbNewLine
   Sql = Sql & " Where V.IDLOJA=" & mvarSys.Propriedades("IDLOJA") & vbNewLine
   'Sql = Sql & " And Convert(varchar, P.DTPGTO,103)=" & SqlStr(Format(pDataIni, "dd/mm/yyyy")) & vbNewLine
   If pDataFim = "" Then
      Sql = Sql & " And Convert(varchar, V.DTVENDA,103)=" & SqlStr(Format(pDataIni, "dd/mm/yyyy")) & vbNewLine
   Else
      Sql = Sql & " And V.DTVENDA>=" & SqlDate(pDataIni) & vbNewLine
      Sql = Sql & " And V.DTVENDA<=" & SqlDate(pDataFim) & vbNewLine
   End If
   Sql = Sql & " And isNull(V.IDFATURA,0)=0" & vbNewLine
   Sql = Sql & " And isNull(V.FLGCANCELADA,0)=0" & vbNewLine
   Sql = Sql & " And V.SITVENDA=10" & vbNewLine
   Sql = Sql & " Group By P.IDFORMAPGTO" & vbNewLine
   With mvarSys.xDb
      If .Abretabela(Sql) Then
         While Not .RsAux.EOF
            If .RsAux("IDFORMAPGTO") = 1 Then '* Dinheiro
               sVLDINDIA = "R$ " & ValBr(.RsAux("Valor"))
            ElseIf .RsAux("IDFORMAPGTO") = 2 Then '* Débito
               sVLDEBDIA = "R$ " & ValBr(.RsAux("Valor"))
               nVLCartao = nVLCartao + xVal(.RsAux("Valor"))
            ElseIf .RsAux("IDFORMAPGTO") = 3 Then '* Crédito
               sVLCREDIA = "R$ " & ValBr(.RsAux("Valor"))
               nVLCartao = nVLCartao + xVal(.RsAux("Valor"))
            ElseIf .RsAux("IDFORMAPGTO") = 4 Then '* Cheque
               sVLCHEDIA = "R$ " & ValBr(.RsAux("Valor"))
            End If
            .RsAux.MoveNext
         Wend
      End If
   
      '************
      '* Conta Caixa
      Dim pTbConta As Object
      
      If xVal(Replace(mvarMe.LblVLCaixa.Caption, "R$", "")) = 0 Then
         Set oMovCC = CriarObjeto("Financ3R.TL_MOVCC")
         Set oMovCC.Sys = mvarSys
         oMovCC.IDLOJA = mvarSys.Propriedades("IDLOJA")
         If DefineContaLAN(pTbConta, 1, mvarSys) Then
            oMovCC.IDCONTA = pTbConta.IDCONTA
         Else
            oMovCC.IDCONTA = mvarSys.Propriedades("IDCONTA")
         End If
         
         Sql = "Select Max(DTEMISSAO) as [DTEMISSAO] "
         Sql = Sql & " From FLAN "
         Sql = Sql & " Where IDLOJA =" & pTbConta.IDLOJA
         Sql = Sql & " And IDCONTA=" & pTbConta.IDCONTA
         If .Abretabela(Sql) Then
            sData = .RsAux("DTEMISSAO")
         Else
            sData = mvarSys.xDb.SysDate()
         End If
         sData = Format(sData, "dd/mm/yyyy")
         
         Sql = oMovCC.QueryGrid
         Sql = Sql & " And F.DTVENCIMENTO>=" & SqlDate(sData)
         Sql = Sql & " Order By NLAN"
         
         If .Abretabela(Sql) Then
            .RsAux.MoveLast
            sVLCAIXADIA = "R$ " & ValBr(.RsAux("Saldo"))
         Else
            sVLCAIXADIA = "R$ 0,00"
         End If
      Else
         sVLCAIXADIA = xVal(Replace(mvarMe.LblVLCaixa.Caption, "R$", ""))
      End If
   End With
   mvarMe.LblDeb.AutoSize = True
   
   'mvarMe.LblCred.Caption = "CRÉDITO : R$"
   'mvarMe.LblDeb.Caption = "DÉBITO  : R$"
   'mvarMe.LblDin.Caption = "DINHEIRO: R$"
   'mvarMe.LblChe.Caption = "CHEQUE  : R$"
   'mvarMe.LblDia.Caption = "TOTAL   : R$"
   
   'mvarMe.LblTotCT.Caption = "TOT. CARTÃO: R$"
   'mvarMe.LblCaixa.Caption = "CAIXA   : R$ "
   
   
   mvarMe.LblVLTotCT.Caption = ValBr(nVLCartao)
   mvarMe.LblVLDin.Caption = ValBr(sVLDINDIA)
   mvarMe.LblVLChe.Caption = ValBr(sVLCHEDIA)
   mvarMe.LbVLDia.Caption = ValBr(xVal(nVLCartao) + xVal(sVLDINDIA) + xVal(sVLCHEDIA))
   
   mvarMe.LblVLCred.Caption = ValBr(sVLCREDIA)
   mvarMe.LblVLDeb.Caption = ValBr(sVLDEBDIA)
   
   mvarMe.LblVLCaixa.Caption = ValBr(sVLCAIXADIA)

   Screen.MousePointer = vbDefault
End Sub
Private Sub FecharDia(pData As String)
   Dim Sql As String
     
   With mvarMe.TaskDialog
      .Reset
      .AddButton "OK", 101
      .DefaultButton = 101
      
      .MessageBoxStyle = True
      .WindowTitle = "Valores do Dia"
      .VerifyCheckState = False
      
      
      '************
      '* Valores do Dia
      Dim sVLDINDIA As String
      Dim sVLDEBDIA As String
      Dim sVLCREDIA As String
      Dim sVLCHEDIA As String
      Dim sVLCAIXADIA As String
      Dim nVLCartao As Currency
      Dim oMovCC As Object
            
If Not mvarMe.GrpValoresDia.Visible Then
      sVLDINDIA = "R$ 0,00"
      sVLDEBDIA = "R$ 0,00"
      sVLCREDIA = "R$ 0,00"
      sVLCHEDIA = "R$ 0,00"
      nVLCartao = 0
      
      Sql = "Select P.IDFORMAPGTO, SUM(P.VLPGTO)-SUM(V.VLTROCO) [Valor]"
      Sql = Sql & " From CPGTOSVENDA P "
      Sql = Sql & " Join CVENDA V On P.IDLOJA=V.IDLOJA And P.IDVENDA=V.IDVENDA"
      Sql = Sql & " Where V.IDLOJA=" & mvarSys.Propriedades("IDLOJA")
      'Sql = Sql & " And Convert(varchar, P.DTPGTO,103)=" & SqlStr(Format(pData, "dd/mm/yyyy"))
      Sql = Sql & " And Convert(varchar, V.DTVENDA,103)=" & SqlStr(Format(pData, "dd/mm/yyyy"))
      Sql = Sql & " And isNull(V.IDFATURA,0)=0"
      Sql = Sql & " And isNull(V.FLGCANCELADA,0)=0"
      Sql = Sql & " And V.SITVENDA=10"
      Sql = Sql & " Group By P.IDFORMAPGTO"
      With mvarSys.xDb
         If mvarSys.xDb.Abretabela(Sql) Then
            While Not .RsAux.EOF
               If .RsAux("IDFORMAPGTO") = 1 Then '* Dinheiro
                  sVLDINDIA = "R$ " & ValBr(.RsAux("Valor"))
               ElseIf .RsAux("IDFORMAPGTO") = 2 Then '* Débito
                  sVLDEBDIA = "R$ " & ValBr(.RsAux("Valor"))
                  nVLCartao = nVLCartao + xVal(.RsAux("Valor"))
               ElseIf .RsAux("IDFORMAPGTO") = 3 Then '* Crédito
                  sVLCREDIA = "R$ " & ValBr(.RsAux("Valor"))
                  nVLCartao = nVLCartao + xVal(.RsAux("Valor"))
               ElseIf .RsAux("IDFORMAPGTO") = 4 Then '* Cheque
                  sVLCHEDIA = "R$ " & ValBr(.RsAux("Valor"))
               End If
               .RsAux.MoveNext
            Wend
         End If
      
         '************
         '* Conta Caixa
         Dim pTbConta As Object
         
         Set oMovCC = CriarObjeto("Financ3R.TL_MOVCC")
         Set oMovCC.Sys = mvarSys
         oMovCC.IDLOJA = mvarSys.Propriedades("IDLOJA")
         If DefineContaLAN(pTbConta, 1, mvarSys) Then
            oMovCC.IDCONTA = pTbConta.IDCONTA
         Else
            oMovCC.IDCONTA = mvarSys.Propriedades("IDCONTA")
         End If
         Sql = oMovCC.QueryGrid
         Sql = Sql & " Order By NLAN"
         If .Abretabela(Sql) Then
            .RsAux.MoveLast
            sVLCAIXADIA = "R$ " & ValBr(.RsAux("Saldo"))
         Else
            sVLCAIXADIA = "R$ 0,00"
         End If
      End With
      .ContentText = ""
      .ContentText = .ContentText & "Fechamento Crédito  : " & sVLCREDIA & vbNewLine
      .ContentText = .ContentText & "Fechamento Débito   : " & sVLDEBDIA & vbNewLine
      .ContentText = .ContentText & "                                   =========" & vbNewLine
      .ContentText = .ContentText & "                     Total       R$ " & ValBr(nVLCartao) & vbNewLine & vbNewLine
      '.ContentText = .ContentText & "Fechamento Dinheiro : " & sVLDINDIA & vbNewLine
      .ContentText = .ContentText & "Fechamento Cheque  : " & sVLCHEDIA & vbNewLine & vbNewLine
      .ContentText = .ContentText & "Fechamento Caixa : " & sVLCAIXADIA & vbNewLine
      
      
      '.FooterIcon = xtpTaskIconInformation
      '.FooterText = "Conta Caixa : R$ 535,00"
      '.FooterText = .FooterText &      "Para enviar o fechamento aos sócios é necessário estar conectado à Internet. Caso a conexão não esteja ativa, desmarque a opção acima e continue o fechamento normalmente."
      .CommonButtons = 0
      .DialogWidth = 50
      '.RtlLayout = True
      .ShowDialog
End If
   
      .Reset
      .MessageBoxStyle = True
      .WindowTitle = "Fechamento do Dia"
      .VerifyCheckState = True
      
      .FooterIcon = xtpTaskIconInformation
      .FooterText = "Para enviar o fechamento aos sócios é necessário estar conectado à Internet. Caso a conexão não esteja ativa, desmarque a opção acima e continue o fechamento normalmente."
      .CommonButtons = 0
      If nTotalErros = 0 Then
         'From General Tab
         '.EnableCallbackTimer = True
         '.ProgressBarStyle = xtpTaskProgressBarNormal
         '.AllowCancellation = True
         '.SetProgressBarRange 0, 5000
         
         
         .VerificationText = "Comunicação aos sócios"
         If nTotalAvisos = 0 Then
            .CommonButtons = xtpTaskButtonCancel Or xtpTaskButtonOk
            .DefaultButton = xtpTaskButtonOk
            .MainInstructionText = "Fechamento Ok."
            .ContentText = "A seguir será realizado o fechamento das atividades e a cópia de segurança. Esta operação poderá levar alguns minutos, por favor aguarde."
         
            On Error Resume Next
            If mvarSys.MDI.Icon = 0 Then
               .MainIcon = xtpTaskIconShield
            Else
               .MainIcon = xtpTaskIconCustom
               .MainIconHandle = mvarSys.MDI.Icon
            End If
         Else
            .CommonButtons = xtpTaskButtonYes Or xtpTaskButtonNo
            
            .MainInstructionText = "Existem alguns comentários no fechamento." & vbNewLine & vbNewLine & "Deseja continuar?"
            .ContentText = "Os comentários existentes não impedem o fechamento do dia. Caso continue, o fechamento das atividades e a cópia de segurança será realizada normalmente."
            '.FooterText = "Os comentários existentes não impedem o fechamento do dia. Caso continue, o fechamento das atividades e a cópia de segurança será realizada normalmente. Lembrando que esta operação poderá levar alguns minutos, por favor aguarde."
            .ExpandedInformationText = "Total de Erros: " & StrZero(nTotalErros, 2) & vbNewLine & "Total de Avisos: " & StrZero(nTotalAvisos, 2)
            .ExpandedControlText = "Detalhes..."
            .MainIcon = xtpTaskIconWarning
         End If
                           
         .ShowDialog
               
      Else
         'From General Tab
         .AddButton "OK", 101
         .DefaultButton = 101
         If nTotalErros = 1 Then
            .MainInstructionText = "Foi encontrado um erro e o fechamento não poderá continuar."
         Else
            .MainInstructionText = "Foram encontrados alguns erros e o fechamento não poderá continuar."
         End If
         '.ContentText = ""
         .FooterText = "Por favor retorne a tela anterior, ajuste as sinalizações e execute a verificação novamente."
         .VerificationText = ""
         .VerifyCheckState = True
         
         'From Expanded Tab
         .ExpandedInformationText = "Total de Erros: " & StrZero(nTotalErros, 2) & vbNewLine & "Total de Avisos: " & StrZero(nTotalAvisos, 2)
         .ExpandedControlText = "Detalhes..."
         .CollapsedControlText = .ExpandedControlText
         .FooterIcon = xtpTaskIconInformation
         .MainIcon = xtpTaskIconError
                     
         .ShowDialog
      End If
      
   End With
         
   Screen.MousePointer = vbDefault
End Sub
Private Sub ResizeGrid()
   With mvarMe.GrdFecharDia
      .AllowColumnResize = True
      .Columns(0).Width = 4
      .Columns(1).Width = 2
      .Columns(2).Width = 30
      .AllowColumnResize = False
   End With
End Sub
Private Sub ExibirItens()
   Dim xRec As ReportRecord
   Dim xSubRec As ReportRecord
   Dim nVisible As Integer
   
   For Each xRec In mvarMe.GrdFecharDia.Records
      nVisible = 0
      For Each xSubRec In xRec.Childs
         If xSubRec.Item(1).Icon = Bandeira.Ok Then xSubRec.Visible = (mvarMe.ChkOk.Value = xtpChecked)
         If xSubRec.Item(1).Icon = Bandeira.Erro Then xSubRec.Visible = (mvarMe.ChkErros.Value = xtpChecked)
         If xSubRec.Item(1).Icon = Bandeira.Aviso Then xSubRec.Visible = (mvarMe.ChkAvisos.Value = xtpChecked)
         
         If Not xSubRec.Visible Then nVisible = nVisible + 1
      Next
      xRec.Visible = True
      If xRec.Childs.Count = nVisible Then
         If xRec.Item(0).Icon = Bandeira.Verde Then xRec.Visible = (mvarMe.ChkOk.Value = xtpChecked)
         If xRec.Item(0).Icon = Bandeira.Vermelha Then xRec.Visible = (mvarMe.ChkErros.Value = xtpChecked)
         If xRec.Item(0).Icon = Bandeira.Amarela Then xRec.Visible = (mvarMe.ChkAvisos.Value = xtpChecked)
      End If
   Next
   mvarMe.GrdFecharDia.Populate
End Sub
Private Sub LimparGrid()
   Dim xCol As ReportColumn
   Dim i    As Integer
   
   With mvarMe.GrdFecharDia
      .Tag = ""
      .Icons = mvarMe.ImageManager1.Icons
      .ShowGroupBox = False
      .BorderStyle = xtpBorderFrame
      .AllowColumnRemove = False
      .AllowColumnReorder = False
      .AllowColumnResize = False
      .AllowColumnSort = False
      .AllowEdit = False
      .MultipleSelection = False
      .ShowHeaderRows = False

      
      .Records.DeleteAll
      .Columns.DeleteAll
      Set xCol = .Columns.Add(0, " ", 30, False)
      Set xCol = .Columns.Add(1, " ", 30, False)
      Set xCol = .Columns.Add(2, "Itens", 1, True)
      
      For i = 0 To .Columns.Count - 1
         .Columns(i).TreeColumn = (i = 0)
         .Columns(i).AutoSortWhenGrouped = True
         .Columns(i).AllowDrag = False
         .Columns(i).AllowRemove = False
         .Columns(i).AutoSize = True
         .Columns(i).Editable = False
      Next
      
      .GroupsOrder.Add .Columns(0)
      .Populate
   End With
End Sub
Private Sub VerificarEventos(pData As String)
   Dim xSubRec    As ReportRecord
   Dim xRecord    As ReportRecord
   Dim Sql        As String
   Dim MyRs       As Object
   Dim sIDAntes   As String
   Dim sGrupo     As String
   Dim i          As Integer
   Dim sCaption   As String
   Dim QtdIni     As Integer
   Dim QtdTotal   As Integer
   Dim QtdErro    As Integer
   Dim QtdAviso   As Integer
   Dim bNewId     As Boolean
   Dim nTreeColumn As Integer
   Dim RGrp       As Integer
   
   Sql = "Select E.IDLOJA, E.IDEVENTO, E.FLGCONFIRMADO, E.FLGCANCELADO, E.STARTDATETIME" & vbNewLine
   Sql = Sql & ", C.IDCLIENTE, C.NOME, A.IDATENDIMENTO" & vbNewLine
   Sql = Sql & " From OEVENTOAGENDA E" & vbNewLine
   Sql = Sql & " Left Join OCLIENTE      C On E.IDLOJA=C.IDLOJA AND E.IDCLIENTE=C.IDCLIENTE" & vbNewLine
   Sql = Sql & " Left Join OATENDIMENTO  A On A.IDLOJA=E.IDLOJA And A.IDEVENTO=E.IDEVENTO" & vbNewLine
   Sql = Sql & " Where CONVERT(VARCHAR(10), E.STARTDATETIME, 103 )=" & SqlDate(pData, 1) & vbNewLine
   Sql = Sql & " And E.ISMEETING = 1" & vbNewLine
   Sql = Sql & " And E.IDLOJA=" & mvarSys.Propriedades("IDLOJA") & vbNewLine
   'Sql = Sql & " And E.FLGCANCELADO <> 1"
   If Val(mvarSys.Propriedades("FECHAPARCIAL", True)) Then
      Sql = Sql & " And E.ENDDATETIME <=" & SqlDate(mvarSys.xDb.SysDate(3)) & vbNewLine
   End If
   Sql = Sql & " Order by E.STARTDATETIME" & vbNewLine
      
   sGrupo = "OEVENTOAGENDA"
   QtdIni = mvarMe.GrdFecharDia.Records.Count
   nTreeColumn = 0
   If mvarSys.xDb.Abretabela(Sql, MyRs) Then
      sIDAntes = ""
      QtdTotal = 0
      QtdErro = 0
      QtdAviso = 0
      With MyRs
         While Not .EOF
            If sIDAntes <> .FIELDS("IDEVENTO") & "" Then
               QtdTotal = QtdTotal + 1
            
               sCaption = "Evento " & StrZero(.FIELDS("IDEVENTO") & "", xVal(mvarSys.Propriedades("LENIDVENDA")))
               sCaption = sCaption & " - " & Format(.FIELDS("STARTDATETIME") & "", "HH:MM")
               sCaption = sCaption & " - " & Trim(Mid(.FIELDS("NOME") & "", 1, 35))
               
               Set xRecord = AddReportRecord(mvarMe.GrdFecharDia, Nothing, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
               xRecord.Expanded = True
               xRecord.Item(0).Icon = Bandeira.Verde
               If .FIELDS("FLGCANCELADO") = 1 Then
                  sCaption = "Evento " & StrZero(.FIELDS("IDEVENTO") & "", xVal(mvarSys.Propriedades("LENIDVENDA")))
                  sCaption = sCaption & " - " & Format(.FIELDS("STARTDATETIME") & "", "HH:MM")
                  sCaption = sCaption & " - " & Trim(Mid(.FIELDS("NOME") & "", 1, 25))
                  sCaption = sCaption & "... (Cancelado)"
                  
                  xRecord.Item(2).Caption = sCaption
               End If
               xRecord.Item(2).Hyperlinks.AddHyperlink Len("Evento "), xVal(mvarSys.Propriedades("LENIDVENDA"))
               Call SetTag(xRecord.Item(2), "TABELA", sGrupo)
            
               '**************
               '* ERROS
               '**************
               '*********
               '* Verificar Atendimentos
               If .FIELDS("IDATENDIMENTO") & "" = "" And .FIELDS("FLGCANCELADO") <> 1 Then
                  sCaption = "Sem atendimento associado."
                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                  xSubRec.Item(1).Icon = Bandeira.Erro
               End If
               '*********
               '* Verificar Integridades
               If Val(.FIELDS("IDEVENTO") & "") < 0 Then
                  sCaption = "Sem integridade no campo [Identificador] (" & .FIELDS("IDEVENTO") & "" & ")"
                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                  xSubRec.Item(1).Icon = Bandeira.Erro
               End If
               
            End If
            
            sIDAntes = .FIELDS("IDEVENTO") & ""
            .MoveNext
            
            bNewId = .EOF
            If Not bNewId Then bNewId = (sIDAntes <> .FIELDS("IDATENDIMENTO") & "")
            If bNewId Then
               If xRecord.Item(0).Icon <> Bandeira.Vermelha Then
                  For Each xSubRec In xRecord.Childs
                     If xSubRec.Item(1).Icon = Bandeira.Aviso Then
                        If xRecord.Item(0).Icon <> Bandeira.Vermelha Then
                           xRecord.Item(0).Icon = Bandeira.Amarela
                        End If
                        QtdAviso = QtdAviso + 1
                     End If
                     If xSubRec.Item(1).Icon = Bandeira.Erro Or xSubRec.Item(1).Icon = Bandeira.Aviso Then
                        xRecord.Item(0).Icon = Bandeira.Vermelha
                        QtdErro = QtdErro + 1
                     End If
                  Next
               End If
               If xRecord.Item(0).Icon = Bandeira.Verde Then
                  xRecord.Childs.DeleteAll
                  xRecord.Item(2).Hyperlinks.RemoveAll
               End If
            End If
         Wend
      End With
   Else
      Set xRecord = AddReportRecord(mvarMe.GrdFecharDia, Nothing, Array("", "", "Sem eventos para analisar"), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
      xRecord.Expanded = False
      xRecord.Item(0).Icon = Bandeira.Verde
   End If
   
   '* Determina Grupo
   With mvarMe.GrdFecharDia
      For i = QtdIni To .Records.Count - 1
         Set xSubRec = .Records(i)
         With xSubRec.Item(0)
            If .GroupCaption = sGrupo Then
               .GroupCaption = ""
               .GroupCaption = .GroupCaption & StrZero(QtdTotal, IIf(QtdTotal = 0, 0, 2)) & " Evento" & IIf(QtdTotal = 1, "", "s")
               .GroupCaption = .GroupCaption & Space(14) & " - " & StrZero(QtdErro, IIf(QtdErro = 0, 0, 2)) & " Erro" & IIf(QtdErro = 1, "", "s")
               .GroupCaption = .GroupCaption & " / " & StrZero(QtdAviso, IIf(QtdAviso = 0, 0, 2)) & " Aviso" & IIf(QtdAviso = 1, "", "s")
               .GroupPriority = 1
            End If
         End With
      Next
      nTotalErros = nTotalErros + QtdErro
      nTotalAvisos = nTotalAvisos + QtdAviso
      
      .Populate
      Call SetTag(mvarMe.GrdFecharDia, "Erro" & (QtdIni), IIf(QtdErro + QtdAviso = 0, False, True))
   End With
End Sub
Private Sub VerificarAtendimentos(pData As String)
   Dim xSubRec    As ReportRecord
   Dim xRecord    As ReportRecord
   Dim Sql        As String
   Dim MyRs       As Object
   Dim sIDAntes   As String
   Dim sGrupo     As String
   Dim i          As Integer
   Dim sCaption   As String
   Dim QtdIni     As Integer
   Dim QtdTotal   As Integer
   Dim QtdErro    As Integer
   Dim QtdAviso   As Integer
   Dim bNewId     As Boolean
   
   Dim bSessaoOK   As Boolean
   Dim bDirecaoOK  As Boolean
   Dim bFotoOK     As Boolean
   Dim bPotenciaOK As Boolean
   Dim bDisparosOK As Boolean
   Dim bTempoOK    As Boolean
   
   Dim bVendaOk    As Boolean
   Dim bFaturaOk   As Boolean
   Dim bDeveTerVenda As Boolean
         
   Sql = ""
   Sql = Sql & "Select A.IDLOJA, A.IDATENDIMENTO, A.IDEVENTO, A.IDCLIENTE, A.IDFUNCIONARIO" & vbNewLine
   Sql = Sql & ", A.IDMAQUINA, A.IDSALA, A.HHINI, A.HHFIM, S.IDSESSAO" & vbNewLine
   Sql = Sql & ", S.IDTPSERVICO, S.IDTPTRATAMENTO, S.IDAREA, S.IDDIRECAO, S.FOTOTIPO" & vbNewLine
   Sql = Sql & ", S.POTENCIA, S.DISPAROS, S.NUMSESSAO, S.TEMPOSESSAO, V.IDVENDA" & vbNewLine
   Sql = Sql & ", F.IDFATURA, P.IDPROD, C.IDCLIENTE, C.NOME, E.IDEVENTO" & vbNewLine
   Sql = Sql & ", P2.EVENDA, P2.VLVENDA, T.FLGDISPARO" & vbNewLine
   Sql = Sql & " From OATENDIMENTO               A" & vbNewLine
   Sql = Sql & " Join OCLIENTE                   C On  A.IDLOJA = C.IDLOJA And A.IDCLIENTE=C.IDCLIENTE" & vbNewLine
   Sql = Sql & " Left Join OSESSAO               S On  A.IDLOJA = S.IDLOJA And A.IDATENDIMENTO=S.IDATENDIMENTO" & vbNewLine
   Sql = Sql & " Left Join OTPTRATAMENTO         T On  S.IDLOJA = T.IDLOJA And S.IDTPTRATAMENTO=T.IDTPTRATAMENTO " & vbNewLine
   Sql = Sql & " Left Join OATENDIMENTO_VENDA    V On  A.IDLOJA = V.IDLOJA And A.IDATENDIMENTO=V.IDATENDIMENTO" & vbNewLine
   Sql = Sql & " Left Join OATENDIMENTO_PRODUTO AP On  A.IDLOJA =AP.IDLOJA And A.IDATENDIMENTO=AP.IDATENDIMENTO" & vbNewLine
   Sql = Sql & " Left Join FFATURA               F On  A.IDLOJA = F.IDLOJA And A.IDATENDIMENTO=F.IDATENDIMENTO" & vbNewLine
   Sql = Sql & " Left Join SPRODUTO              P On AP.IDLOJA = P.IDLOJA And AP.IDPROD=P.IDPROD  And P.EVENDA=1" & vbNewLine
   Sql = Sql & " Left Join OEVENTOAGENDA         E On  E.IDLOJA = A.IDLOJA And E.IDEVENTO=A.IDEVENTO" & vbNewLine
   Sql = Sql & " Left Join OTRATAMENTO_PROD     TP On TP.IDLOJA = S.IDLOJA And TP.IDTPSERVICO=S.IDTPSERVICO And TP.IDTPTRATAMENTO=S.IDTPTRATAMENTO" & vbNewLine
   Sql = Sql & " Left Join SPRODUTO             P2 On TP.IDLOJA =P2.IDLOJA And TP.IDPROD =P2.IDPROD And P2.EVENDA=1" & vbNewLine
   Sql = Sql & " Where CONVERT(VARCHAR(10), A.DTATEND, 103 )=" & SqlDate(pData, 1) & vbNewLine
   Sql = Sql & " And A.SITATEND Not In ('9X')" ', '9C')"& vbNewLine
   Sql = Sql & " And A.IDLOJA=" & mvarSys.Propriedades("IDLOJA") & vbNewLine
   If Val(mvarSys.Propriedades("FECHAPARCIAL")) Then
      Sql = Sql & " And A.HHFIM <=" & SqlDate(mvarSys.xDb.SysDate(3)) & vbNewLine
   End If
   Sql = Sql & " And A.IDATENDIMENTO<>0" & vbNewLine
   Sql = Sql & " Order by A.DTATEND, A.HHINI"
      
   sGrupo = "OATENDIMENTO"
   QtdIni = mvarMe.GrdFecharDia.Records.Count
   If mvarSys.xDb.Abretabela(Sql, MyRs) Then
      sIDAntes = ""
      QtdTotal = 0
      QtdErro = 0
      QtdAviso = 0
      
      bSessaoOK = True
      bDirecaoOK = True
      bFotoOK = True
      bPotenciaOK = True
      bDisparosOK = True
      bTempoOK = True
      
      bVendaOk = True
      bDeveTerVenda = False
      
      Dim nTreeColumn As Integer
      nTreeColumn = 0
      With MyRs
         While Not .EOF
            If .FIELDS("IDATENDIMENTO") & "" = "12138" Then
               sIDAntes = sIDAntes
            End If
            If sIDAntes <> .FIELDS("IDATENDIMENTO") & "" Then
               QtdTotal = QtdTotal + 1
            
               sCaption = "Atendimento " & StrZero(.FIELDS("IDATENDIMENTO") & "", xVal(mvarSys.Propriedades("LENIDVENDA")))
               sCaption = sCaption & " - " & Format(.FIELDS("HHINI") & "", "HH:MM")
               sCaption = sCaption & " - " & Trim(Mid(.FIELDS("NOME") & "", 1, 35))
               
               Set xRecord = AddReportRecord(mvarMe.GrdFecharDia, Nothing, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
               xRecord.Expanded = True
               xRecord.Item(0).Icon = Bandeira.Verde
               xRecord.Item(2).Hyperlinks.AddHyperlink Len("Atendimento "), xVal(mvarSys.Propriedades("LENIDVENDA"))
               Call SetTag(xRecord.Item(2), "TABELA", sGrupo)
'mvarMe.GrdFecharDia.Populate
               '**************
               '* ERROS
               '**************
               
               '**************
               '* Verificar Cliente
               If xVal(.FIELDS("IDCLIENTE") & "") <= 0 Then
                  sCaption = "Sem cliente definido."
                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                  xSubRec.Item(1).Icon = Bandeira.Erro
               End If
               
               '**************
               '* Verificar Sala
               If xVal(.FIELDS("IDSALA") & "") <= 0 Then
                  sCaption = "Sem sala/local definido."
                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                  xSubRec.Item(1).Icon = Bandeira.Erro
               End If
               
               '**************
               '* Verificar Máquina
               If xVal(.FIELDS("FLGDISPARO") & "") = 1 Then
                  If xVal(.FIELDS("IDMAQUINA") & "") <= 0 Then
                     sCaption = "Sem máquina definida."
                     Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                     xSubRec.Item(1).Icon = Bandeira.Erro
                  End If
               End If
               '**************
               '* Verificar Operadora
               If xVal(.FIELDS("IDFUNCIONARIO") & "") <= 0 Then
                  sCaption = "Sem funcionária definida."
                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                  xSubRec.Item(1).Icon = Bandeira.Erro
               End If
               
               '**************
               '* Verificar Evento
               If xVal(.FIELDS("IDEVENTO") & "") <= 0 Then
                  sCaption = "Sem evento associado."
                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                  xSubRec.Item(1).Icon = Bandeira.Erro
               End If
               
               '**************
               '* Verificar Existencia Venda
               bVendaOk = (xVal(.FIELDS("IDVENDA") & "") > 0) Or (xVal(.FIELDS("IDTPSERVICO") & "") = 1)
               bFaturaOk = (xVal(.FIELDS("IDFATURA") & "") > 0)
               bDeveTerVenda = (xVal(.FIELDS("EVENDA") & "") = 1 And xVal(.FIELDS("VLVENDA") & "") > 0)
               
               
               '**************
               '* Verificar Registro de Tempo
               If Not IsDate(.FIELDS("HHINI") & "") Or Not IsDate(.FIELDS("HHINI") & "") Then
                  sCaption = "Inconsistência no horário."
                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                  xSubRec.Item(1).Icon = Bandeira.Erro
               End If
            End If
               
            If sIDAntes = .FIELDS("IDATENDIMENTO") & "" Or Trim(sIDAntes) = "" Then
               bDeveTerVenda = (xVal(.FIELDS("EVENDA") & "") = 1 And xVal(.FIELDS("VLVENDA") & "") > 0)
               
               '**************
               '* AVISOS
               '**************
               
               bSessaoOK = IIf(bSessaoOK, (xVal(.FIELDS("IDSESSAO") & "") > 0), bSessaoOK)            '* Verificar Nº da Sessao
               If xVal(.FIELDS("FLGDISPARO") & "") = 1 Then
                  bDirecaoOK = IIf(bSessaoOK, (xVal(.FIELDS("IDSESSAO") & "") > 0), bSessaoOK)        '* Verificar Nº Direção
                  bFotoOK = IIf(bFotoOK, (xVal(.FIELDS("FOTOTIPO") & "") > 0), bFotoOK)               '* Verificar Nº Fototipo
                  bPotenciaOK = IIf(bPotenciaOK, (xVal(.FIELDS("POTENCIA") & "") > 0), bPotenciaOK)   '* Verificar Nº Potência
                  bDisparosOK = IIf(bSessaoOK, (xVal(.FIELDS("IDSESSAO") & "") > 0), bSessaoOK)       '* Verificar Nº Disparos
                  bTempoOK = IIf(bSessaoOK, (xVal(.FIELDS("IDSESSAO") & "") > 0), bSessaoOK)          '* Verificar Nº Tempo
               Else
                  bDirecaoOK = True    '* Verificar Nº Direção
                  bFotoOK = True       '* Verificar Nº Fototipo
                  bPotenciaOK = True   '* Verificar Nº Potência
                  bDisparosOK = True   '* Verificar Nº Disparos
                  bTempoOK = True      '* Verificar Nº Tempo
               End If
            End If
            
            sIDAntes = .FIELDS("IDATENDIMENTO") & ""
            .MoveNext
            
            bNewId = .EOF
            If Not bNewId Then bNewId = (sIDAntes <> .FIELDS("IDATENDIMENTO") & "")
            If bNewId Then
               If Not bVendaOk Then
                  If bFaturaOk Then
                     sCaption = "Fatura em aberto."
                     Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                     xSubRec.Item(1).Icon = Bandeira.Aviso
                  End If
                  If bDeveTerVenda Then '* DeveExistirVenda
                     sCaption = "Sem venda associada."
                     Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                     xSubRec.Item(1).Icon = Bandeira.Erro
                  End If
               End If
            
               If Not bSessaoOK Then
                  sCaption = "Erro no campo Sessão."
                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                  xSubRec.Item(1).Icon = Bandeira.Erro
               End If
               If Not bDirecaoOK Then
                  sCaption = "Sem integridade no campo [Direção]."
                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                  xSubRec.Item(1).Icon = Bandeira.Aviso
               End If
               If Not bFotoOK Then
                  sCaption = "Sem integridade no campo [Fototipo]."
                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                  xSubRec.Item(1).Icon = Bandeira.Erro
               End If
               If Not bPotenciaOK Then
                  sCaption = "Sem integridade no campo [Potência]."
                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                  xSubRec.Item(1).Icon = Bandeira.Erro
               End If
               If Not bDisparosOK Then
                  sCaption = "Sem integridade no campo [Nº Disparos]."
                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                  xSubRec.Item(1).Icon = Bandeira.Aviso
               End If
               If Not bTempoOK Then
                  sCaption = "Sem integridade no campo [Tempo da Sessão]."
                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                  xSubRec.Item(1).Icon = Bandeira.Aviso
               End If
                              
               bSessaoOK = True
               bDirecaoOK = True
               bFotoOK = True
               bPotenciaOK = True
               bDisparosOK = True
               bTempoOK = True
               
               bVendaOk = True
               bDeveTerVenda = False
                              
               If xRecord.Item(0).Icon <> Bandeira.Vermelha Then
                  For Each xSubRec In xRecord.Childs
                     If xSubRec.Item(1).Icon = Bandeira.Aviso Then
                        If xRecord.Item(0).Icon <> Bandeira.Vermelha Then
                           xRecord.Item(0).Icon = Bandeira.Amarela
                        End If
                        QtdAviso = QtdAviso + 1
                     End If
                     If xSubRec.Item(1).Icon = Bandeira.Erro Then
                        xRecord.Item(0).Icon = Bandeira.Vermelha
                        QtdErro = QtdErro + 1
                     End If
                  Next
               End If
               If xRecord.Item(0).Icon = Bandeira.Verde Then
                  xRecord.Childs.DeleteAll
                  'xRecord.Item(2).Hyperlinks.RemoveAll
               End If
            End If
         Wend
      End With
   Else
      sCaption = "Sem atendimentos para analisar"
      Set xRecord = AddReportRecord(mvarMe.GrdFecharDia, Nothing, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
      xRecord.Expanded = False
      xRecord.Item(0).Icon = Bandeira.Verde
   End If
   
   '* Determina Grupo
   With mvarMe.GrdFecharDia
      For i = QtdIni To .Records.Count - 1
         Set xSubRec = .Records(i)
         With xSubRec.Item(0)
            If .GroupCaption = sGrupo Then
               .GroupCaption = ""
               .GroupCaption = .GroupCaption & StrZero(QtdTotal, IIf(QtdTotal = 0, 0, 2)) & " Atendimento" & IIf(QtdTotal = 1, "", "s")
               .GroupCaption = .GroupCaption & Space(5) & " - " & StrZero(QtdErro, IIf(QtdErro = 0, 0, 2)) & " Erro" & IIf(QtdErro = 1, "", "s")
               .GroupCaption = .GroupCaption & " / " & StrZero(QtdAviso, IIf(QtdAviso = 0, 0, 2)) & " Aviso" & IIf(QtdAviso = 1, "", "s")

               .GroupPriority = 1
            End If
         End With
      Next
      nTotalErros = nTotalErros + QtdErro
      nTotalAvisos = nTotalAvisos + QtdAviso
      .Populate
      Call SetTag(mvarMe.GrdFecharDia, "Erro" & (QtdIni + 1), IIf(QtdErro + QtdAviso = 0, False, True))
   End With
End Sub
Private Sub VerificarTarefas(pData As String)
   With mvarMe.GrdFecharDia
   End With
End Sub
Private Sub VerificarVendas(pData As String)
   Dim xSubRec    As ReportRecord
   Dim xRecord    As ReportRecord
   Dim Sql        As String
   Dim MyRs       As Object
   Dim RsAux      As Object
   Dim RsAux2     As Object
   Dim RsCorr     As Object
   Dim sIDAntes   As String
   Dim sGrupo     As String
   Dim i          As Integer
   Dim sCaption   As String
   Dim QtdIni     As Integer
   Dim QtdTotal   As Integer
   Dim QtdErro    As Integer
   Dim QtdAviso   As Integer
   Dim bNewId     As Boolean
   
   Dim bQtdVendaOk   As Boolean
   Dim bVLVendaOK    As Boolean
   Dim bPgtosOK      As Boolean
   Dim bFlanOK       As Boolean
   Dim bFlanExiste   As Boolean
   Dim nQdPAux       As Double
   Dim nQdSAux       As Double
   Dim nQdProd       As Double
   Dim nQdServ       As Double
   Dim nValor        As Currency
   Dim nValor2       As Currency
   Dim nTaxa         As Double
   Dim nIDATEND      As Long
   Dim sId           As String
   Dim nVLPGTO       As Currency
   Dim nVLFAT        As Currency
   Dim nQTDFAT       As Double
   Dim nIDFormaPgto  As Integer
   Dim nIDCondPgto   As Integer
         
   Sql = "Select V.IDLOJA, V.IDVENDA, V.IDCLIENTE, V.IDFUNCIONARIO, V.SITVENDA"
   Sql = Sql & ", C.NOME, V.VLITENS, V.VLVENDA, V.VLPGTO, V.VLDESC"
   Sql = Sql & ", V.VLTROCO"
   Sql = Sql & " From CVENDA     V"
   Sql = Sql & " Join OCLIENTE   C On V.IDLOJA=C.IDLOJA AND V.IDCLIENTE=C.IDCLIENTE"
   Sql = Sql & " Where CONVERT(VARCHAR(10), V.DTVENDA, 103 )=" & SqlDate(pData, 1)
   Sql = Sql & " And V.SITVENDA Not In ('9X')"
   Sql = Sql & " And V.IDLOJA=" & mvarSys.Propriedades("IDLOJA")
   Sql = Sql & " Order by V.IDVENDA"
      
   sGrupo = "CVENDA"
   QtdIni = mvarMe.GrdFecharDia.Records.Count

   If mvarSys.xDb.Abretabela(Sql, MyRs) Then
      sIDAntes = ""
      QtdTotal = 0
      QtdErro = 0
      QtdAviso = 0
      
      bVLVendaOK = True
      
      Dim nTreeColumn As Integer
      nTreeColumn = 0
      With MyRs
         While Not .EOF
            If .FIELDS("IDVENDA") & "" = 4532 Then
               sIDAntes = sIDAntes
            End If
            If sIDAntes <> .FIELDS("IDVENDA") & "" Then
               nVLFAT = 0
               QtdTotal = QtdTotal + 1
            
               sCaption = "Venda " & StrZero(.FIELDS("IDVENDA") & "", xVal(mvarSys.Propriedades("LENIDVENDA")))
               sCaption = sCaption & " - R$" & ValBr(xVal(.FIELDS("VLPGTO") & "") - xVal(.FIELDS("VLTROCO") & "")) 'ValBr(.FIELDS("VLVENDA") & "")
               sCaption = sCaption & " - " & Trim(Mid(.FIELDS("NOME") & "", 1, 35))
               
               Set xRecord = AddReportRecord(mvarMe.GrdFecharDia, Nothing, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
               xRecord.Expanded = True
               xRecord.Item(0).Icon = Bandeira.Verde
               xRecord.Item(2).Hyperlinks.AddHyperlink Len("Venda "), xVal(mvarSys.Propriedades("LENIDVENDA"))
               Call SetTag(xRecord.Item(2), "TABELA", sGrupo)
               '**************
               '* ERROS
               '**************
               '* Verificar Cliente
               If xVal(.FIELDS("IDCLIENTE") & "") <= 0 Then
                  sCaption = "Sem cliente definido."
                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                  xSubRec.Item(1).Icon = Bandeira.Erro
               End If
               
               '* Verificar Operadora
'               If xVal(.Fields("IDFUNCIONARIO") & "") <= 0 Then
'                  sCaption = "Sem funcionária definida."
'                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
'                  xSubRec.Item(1).Icon = Bandeira.Erro
'               End If
                                                            
               '*******************
               '* Verificar Valores na Venda
               nVLPGTO = xVal(.FIELDS("VLPGTO") & "")
               If xVal(.FIELDS("VLPGTO") & "") = 0 Then
                  Sql = "Select Sum(P.VLPGTO) [PG]"
                  Sql = Sql & " From CPGTOSVENDA P "
                  Sql = Sql & " Where P.IDVENDA=" & .FIELDS("IDVENDA")
                  Sql = Sql & " And P.IDLOJA=" & SqlNum(.FIELDS("IDLOJA") & "")
                  If mvarSys.xDb.Abretabela(Sql, RsCorr) Then
                     nVLPGTO = xVal(RsCorr("PG") & "")
                     Sql = "Update CVENDA"
                     Sql = Sql & " Set VLPGTO=" & SqlNum(xVal(RsCorr("PG") & ""))
                     Sql = Sql & " Where IDVENDA=" & .FIELDS("IDVENDA")
                     Sql = Sql & " And IDLOJA=" & SqlNum(.FIELDS("IDLOJA") & "")
                     If mvarSys.xDb.executa(Sql) Then
                        nVLPGTO = xVal(RsCorr("PG") & "")
                     End If
                  End If
               End If

               If nVLFAT + nVLPGTO - xVal(.FIELDS("VLTROCO") & "") <> xVal(.FIELDS("VLVENDA") & "") Then
                  Sql = "Select Sum(F.VALOR) [PG]" & vbNewLine
                  Sql = Sql & " From FFATURA F" & vbNewLine
                  'Sql = Sql & " Join OATENDIMENTO A On F.IDLOJA=A.IDLOJA And F.IDATENDIMENTO=A.IDATENDIMENTO" & vbNewLine
                  'Sql = Sql & " Join OATENDIMENTO_VENDA V On V.IDLOJA=A.IDLOJA And V.IDATENDIMENTO=A.IDATENDIMENTO" & vbNewLine
                  Sql = Sql & " Where F.IDLOJA=" & SqlNum(.FIELDS("IDLOJA") & "") & vbNewLine
                  Sql = Sql & " And F.IDVENDAORIGEM=" & .FIELDS("IDVENDA") & vbNewLine
                  ''Sql = Sql & " And F.SITFATURA<>10" Retirado pois a ftura pode estar paga no mesmo dia
                  'Sql = Sql & " And V.IDVENDA=" & .FIELDS("IDVENDA") & vbNewLine
                  If mvarSys.xDb.Abretabela(Sql, RsCorr) Then
                     nVLFAT = xVal(RsCorr("PG") & "")
                  End If
                  If nVLFAT = 0 Then
                     Sql = "Select IDFATURA" & vbNewLine
                     Sql = Sql & " From FFATURA " & vbNewLine
                     Sql = Sql & " Where IDLOJA=" & SqlNum(.FIELDS("IDLOJA") & "") & vbNewLine
                     Sql = Sql & " And IDVENDAORIGEM=" & .FIELDS("IDVENDA") & vbNewLine
                     Sql = Sql & " Order By IDFATURA"
                     If Not mvarSys.xDb.Abretabela(Sql) Then
                        Sql = "Select V.IDFATURA, isNull(F.VALOR,0) [VALOR]" & vbNewLine
                        Sql = Sql & " From CVENDA V" & vbNewLine
                        Sql = Sql & " Left Join FFATURA F On V.IDLOJA=F.IDLOJA And V.IDFATURA=F.IDFATURA"
                        Sql = Sql & " Where V.IDLOJA=" & SqlNum(.FIELDS("IDLOJA") & "") & vbNewLine
                        Sql = Sql & " And V.IDVENDA=" & .FIELDS("IDVENDA") & vbNewLine
                        Sql = Sql & " And Year(V.DTVENDA)>2012 And V.IDFATURA>1"
                        If mvarSys.xDb.Abretabela(Sql, RsAux) Then
                           Sql = "Update FFATURA"
                           Sql = Sql & " Set IDVENDAORIGEM=" & SqlNum(.FIELDS("IDVENDA") & "")
                           Sql = Sql & " Where IDFATURA=" & RsAux("IDFATURA") & ""
                           Sql = Sql & " And IDLOJA=" & SqlNum(.FIELDS("IDLOJA") & "")
                           If mvarSys.xDb.executa(Sql) Then
                              nVLFAT = xVal(RsAux("VALOR") & "")
                           End If
                        End If
                     End If
                  End If
               End If
               Dim nVLPAGO As Currency
               Dim nVLAPAGAR As Currency
               nVLPAGO = nVLFAT + nVLPGTO - xVal(.FIELDS("VLTROCO") & "")
               nVLAPAGAR = xVal(.FIELDS("VLVENDA") & "")
               If nVLPAGO <> nVLAPAGAR Then
                  sCaption = "Valor pago"
                  sCaption = sCaption & " (R$" & ValBr(nVLPGTO) & ")"
                  sCaption = sCaption & " difere do calculado à pagar"
                  sCaption = sCaption & " (R$" & ValBr(xVal(.FIELDS("VLVENDA") & "")) & ")."
                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                  xSubRec.Item(1).Icon = IIf(nVLPGTO > nVLAPAGAR, Bandeira.Aviso, Bandeira.Erro)
               End If
                                                            
               '*******************
               '* Verificar Registro dos Pagamentos
               Sql = "Select Sum(VLPGTO) As [SOMA_PGTO]"
               Sql = Sql & " From CPGTOSVENDA"
               Sql = Sql & " Where IDLOJA=" & SqlNum(.FIELDS("IDLOJA") & "")
               Sql = Sql & " And IDVENDA=" & SqlNum(.FIELDS("IDVENDA") & "")
               If mvarSys.xDb.Abretabela(Sql, RsAux) Then
                  bPgtosOK = (nVLPGTO = xVal(RsAux("SOMA_PGTO") & ""))
               Else
                  bPgtosOK = (nVLPGTO = 0)
               End If
               If Not bPgtosOK Then
                  sCaption = "Valor de venda pago"
                  sCaption = sCaption & " (R$" & ValBr(nVLPGTO) & ")"
                  sCaption = sCaption & " difere dos pagamentos"
                  sCaption = sCaption & " (R$" & ValBr(xVal(RsAux("SOMA_PGTO") & "")) & ")."
                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                  xSubRec.Item(1).Icon = Bandeira.Erro
               End If
                              
                              
               '*******************************************
               '***** Verificar Atendimento da Venda  *****
               '*******************************************
               Sql = "Select A.IDATENDIMENTO" & vbNewLine
               Sql = Sql & " From OATENDIMENTO     A" & vbNewLine
               Sql = Sql & " Join OATENDIMENTO_VENDA   AV On A.IDLOJA =AV.IDLOJA  And A.IDATENDIMENTO=AV.IDATENDIMENTO" & vbNewLine
               Sql = Sql & " Where AV.IDLOJA=" & SqlNum(.FIELDS("IDLOJA") & "") & vbNewLine
               Sql = Sql & " And AV.IDVENDA=" & SqlNum(.FIELDS("IDVENDA") & "") & vbNewLine
               If mvarSys.xDb.Abretabela(Sql, RsAux) Then
                  '**************
                  '* Verificar quantidade de itens e Valor do Atendimento
                  nIDATEND = RsAux("IDATENDIMENTO")
                  nQdProd = 0
                  nQdServ = 0
                  nQdPAux = 0
                  nQdSAux = 0
                  nValor = 0
                  sId = ""
                  RsAux.MoveFirst
                  While Not RsAux.EOF
                     nIDATEND = RsAux("IDATENDIMENTO")
                     nValor = nValor + CalculaValorAtend((.FIELDS("IDLOJA") & ""), nIDATEND, nQdSAux, nQdPAux)
                     sId = sId & RsAux("IDATENDIMENTO") & " ,"
                     RsAux.MoveNext
                  Wend
                  nQdServ = nQdServ + nQdSAux
                  nQdProd = nQdProd + nQdPAux
                  bVLVendaOK = (nValor <= .FIELDS("VLITENS"))
                  sId = Mid(sId, 1, Len(sId) - 2)
                  If sId <> "" Then
                     If Not bVLVendaOK Then
                        Sql = "Select Sum(V.VLITENS) [VLITENS]" & vbNewLine
                        Sql = Sql & " From CVENDA V" & vbNewLine
                        Sql = Sql & " Join OATENDIMENTO_VENDA   AV On V.IDLOJA =AV.IDLOJA  And V.IDVENDA=AV.IDVENDA" & vbNewLine
                        Sql = Sql & " Where AV.IDLOJA=" & SqlNum(.FIELDS("IDLOJA") & "") & vbNewLine
                        Sql = Sql & " And AV.IDATENDIMENTO IN (" & sId & ")" & vbNewLine
                        If mvarSys.xDb.Abretabela(Sql, RsAux2) Then
                           bVLVendaOK = (nValor <= RsAux2("VLITENS"))
                        End If
                     End If
                     '***********
                     '* Verifica Fatura
                     If Not bVLVendaOK Then
                        nVLFAT = 0
                        Sql = "Select Sum(VALOR) [PG]"
                        Sql = Sql & " From FFATURA F"
                        Sql = Sql & " Join OATENDIMENTO A On F.IDLOJA=A.IDLOJA And F.IDATENDIMENTO=A.IDATENDIMENTO"
                        Sql = Sql & " Join OATENDIMENTO_VENDA V On V.IDLOJA=A.IDLOJA And V.IDATENDIMENTO=A.IDATENDIMENTO"
                        Sql = Sql & " Where F.SITFATURA<>10"
                        Sql = Sql & " And V.IDLOJA=" & SqlNum(.FIELDS("IDLOJA") & "")
                        Sql = Sql & " And V.IDATENDIMENTO IN (" & sId & ")"
                        If mvarSys.xDb.Abretabela(Sql, RsCorr) Then
                           nVLFAT = xVal(RsCorr("PG") & "")
                        End If
                        bVLVendaOK = (nValor <= RsAux2("VLITENS") + nVLFAT)
                     End If
                  
                     Sql = " Select IsNull(Sum(QTDVENDA),0) as [QTDVENDA]"
                     Sql = Sql & " From CITENSVENDA"
                     Sql = Sql & " Where IDLOJA = " & SqlNum(.FIELDS("IDLOJA") & "")
                     Sql = Sql & " And   IDVENDA =" & .FIELDS("IDVENDA")
                     If mvarSys.xDb.Abretabela(Sql, RsAux2) Then
                        bQtdVendaOk = (nQdProd + nQdServ = RsAux2("QTDVENDA"))
                        If Not bQtdVendaOk Then
                           Sql = " Select IsNull(Sum(QTDVENDA),0) as [QTDVENDA],  IsNull(Sum(QTDVENDA*VLUNIT),0) as [VALOR]"
                           Sql = Sql & " From CITENSVENDA"
                           Sql = Sql & " Where IDLOJA = " & SqlNum(.FIELDS("IDLOJA") & "")
                           Sql = Sql & " And   IDVENDA IN (Select IDVENDA"
                           Sql = Sql & "             From OATENDIMENTO_VENDA"
                           Sql = Sql & "             Where IDLOJA=" & SqlNum(.FIELDS("IDLOJA") & "")
                           Sql = Sql & "             And IDATENDIMENTO in (" & sId & "))"
                           If mvarSys.xDb.Abretabela(Sql, RsAux2) Then
                              bQtdVendaOk = (nQdProd + nQdServ <= Val(RsAux2("QTDVENDA")))
                           End If
                        End If
                     End If
                     '*********
                     '* Verifica Fatura
                     If Not bQtdVendaOk Then
                        nQTDFAT = 0
                        Sql = " Select IsNull(Sum(I.QTDITEM),0) as [QTD]"
                        Sql = Sql & " From FITEMFATURA I"
                        Sql = Sql & " Join FFATURA F On F.IDLOJA=I.IDLOJA AnD I.IDFATURA=F.IDFATURA"
                        Sql = Sql & " Where F.SITFATURA <> 10"
                        Sql = Sql & " And F.IDLOJA = " & SqlNum(.FIELDS("IDLOJA") & "")
                        Sql = Sql & " And F.IDATENDIMENTO In (" & sId & ")"
                        
                        If mvarSys.xDb.Abretabela(Sql, RsCorr) Then
                           nQTDFAT = xVal(RsCorr("QTD") & "")
                        End If
                        bQtdVendaOk = (nQdProd + nQdServ <= Val(RsAux2("QTDVENDA")) + nQTDFAT)
                     End If
                     
                  End If
                  If Not bVLVendaOK Or Not bQtdVendaOk Then
                     If Not bVLVendaOK Then
                        '* Verificar Valor de Venda
                        sCaption = "Valor de venda não confere com o valor do atendimento."
                        Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                        xSubRec.Item(1).Icon = Bandeira.Erro
                     End If
                     If Not bQtdVendaOk Then
                        '* Verificar Valor de Venda
                        sCaption = "Quantidade de venda não cenfere com a quantidade do atendimento."
                        Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                        xSubRec.Item(1).Icon = Bandeira.Erro
                     End If
                     
                     
                     '* Procurar e Marcar Atendimento
                     With mvarMe.GrdFecharDia
                        For i = 0 To .Records.Count - 1
                           If GetTag(.Records(i).Item(2), "TABELA") = "OATENDIMENTO" Then
                              If InStr(.Records(3).Item(2).Value, StrZero(nIDATEND, mvarSys.Propriedades("LENIDVENDA"))) <> 0 Then
                                 .Records(i).Item(2).Hyperlinks.AddHyperlink Len("Atendimento "), mvarSys.Propriedades("LENIDVENDA")
                              End If
                           End If
                        Next
                     End With
                  End If
               Else
                  Sql = "Select * "
                  Sql = Sql & " From FFATURA "
                  Sql = Sql & " Where IDLOJA=" & SqlNum(.FIELDS("IDLOJA") & "")
                  Sql = Sql & " And IDVENDA=" & SqlNum(.FIELDS("IDVENDA") & "")
                  If Not mvarSys.xDb.ExisteReg(Sql) Then '* Pagamento de Fatura
                     '* Verificar se Existe Atendimento
                     sCaption = "Venda sem atendimento associado."
                     Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                     xSubRec.Item(1).Icon = Bandeira.Aviso
                  End If
               End If
                              
               
               '* Verificar Situação da Venda
               If .FIELDS("SITVENDA") & "" <> "10" Then
                  sCaption = "Venda sem fechamento."
                  Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                  xSubRec.Item(1).Icon = Bandeira.Erro
               End If
                              
               '* Verificar Lançamento Financeiro
               Sql = "Select P.IDVENDA, P.IDPGTO, P.VLPGTO, P.IDFORMAPGTO, P.IDCONDPGTO"
               Sql = Sql & ", V.VLTROCO, F.IDVENDA, Sum(F.VALOR) [VALOR]"
               Sql = Sql & " From CPGTOSVENDA P"
               Sql = Sql & " Join CVENDA V    On P.IDLOJA =V.IDLOJA And P.IDVENDA=V.IDVENDA "
               Sql = Sql & " Left Join FLAN F On P.IDLOJA =F.IDLOJA And P.IDVENDA=F.IDVENDA And P.IDPGTO=F.IDPGTO AND F.FLGDELETE<>1"
               Sql = Sql & " Where P.IDLOJA=" & SqlNum(.FIELDS("IDLOJA") & "")
               Sql = Sql & " And P.IDVENDA=" & SqlNum(.FIELDS("IDVENDA") & "")
               Sql = Sql & " Group By P.IDVENDA, P.IDPGTO, P.VLPGTO, P.IDFORMAPGTO, P.IDCONDPGTO"
               Sql = Sql & ", V.VLTROCO, F.IDVENDA"
               bFlanOK = False
               bFlanExiste = False
               
               bFlanExiste = mvarSys.xDb.Abretabela(Sql, RsAux)
               If bFlanExiste Then
                  nValor = 0
                  nValor2 = 0
                  RsAux.MoveFirst
                  While Not RsAux.EOF
                     nIDFormaPgto = xVal(RsAux("IDFORMAPGTO") & "")
                     nIDCondPgto = xVal(RsAux("IDCONDPGTO") & "")
                     nValor = xVal(RsAux("VLPGTO") & "") - IIf(nIDFormaPgto = 1, xVal(RsAux("VLTROCO") & ""), 0)
                     If nIDFormaPgto = 2 Or nIDFormaPgto = 3 Then
                        nTaxa = 0
                        Sql = "Select TXSERV, TXPARC "
                        Sql = Sql & " From CFORMAPGTO "
                        Sql = Sql & " Where IDLOJA=" & SqlNum(.FIELDS("IDLOJA") & "")
                        Sql = Sql & " And IDFORMAPGTO=" & nIDFormaPgto
                        If mvarSys.xDb.Abretabela(Sql) Then
                           If nIDCondPgto <= 2 Then
                              nTaxa = xVal(mvarSys.xDb.RsAux("TXSERV") & "")
                           Else
                              nTaxa = xVal(mvarSys.xDb.RsAux("TXPARC") & "")
                           End If
                              
                        End If
                        nValor = nValor - ValBr(nValor * nTaxa / 100, 2, 2)
                     End If
                     nValor2 = xVal(RsAux("Valor") & "")
                     bFlanOK = (nValor = nValor2)
                     If Not bFlanOK Then
                        RsAux.MoveLast
                     End If
                     RsAux.MoveNext
                  Wend
                  If Not bFlanOK Then
                     sCaption = "Valor do pagamento (R$" & ValBr(nValor) & ") difere do lançamento (R$" & ValBr(nValor2) & ")."
                     Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                     xSubRec.Item(1).Icon = Bandeira.Erro
                  End If
                  
               Else
                  If nVLFAT <> .FIELDS("VLVENDA") Then
                     sCaption = "Lançamento financeiro sem pagamento."
                     Set xSubRec = AddReportRecord(mvarMe.GrdFecharDia, xRecord, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
                     If .FIELDS("VLVENDA") <> 0 Then
                        xSubRec.Item(1).Icon = Bandeira.Erro
                     Else
                        xSubRec.Item(1).Icon = Bandeira.Aviso
                     End If
                  End If
               End If
            End If
               
            sIDAntes = .FIELDS("IDVENDA") & ""
            .MoveNext
            
            bNewId = .EOF
            If Not bNewId Then bNewId = (sIDAntes <> .FIELDS("IDVENDA") & "")
            If bNewId Then
               If xRecord.Item(0).Icon <> Bandeira.Vermelha Then
                  For Each xSubRec In xRecord.Childs
                     If xSubRec.Item(1).Icon = Bandeira.Aviso Then
                        If xRecord.Item(0).Icon <> Bandeira.Vermelha Then
                           xRecord.Item(0).Icon = Bandeira.Amarela
                        End If
                        QtdAviso = QtdAviso + 1
                     End If
                     If xSubRec.Item(1).Icon = Bandeira.Erro Then
                        xRecord.Item(0).Icon = Bandeira.Vermelha
                        QtdErro = QtdErro + 1
                     End If
                  Next
               End If
               If xRecord.Item(0).Icon = Bandeira.Verde Then
                  xRecord.Childs.DeleteAll
                  'xRecord.Item(2).Hyperlinks.RemoveAll
               End If
            End If
         Wend
      End With
   Else
      sCaption = "Sem vendas para analisar"
      Set xRecord = AddReportRecord(mvarMe.GrdFecharDia, Nothing, Array("", "", sCaption), TreeColumn:=nTreeColumn, GroupCaption:=sGrupo)
      xRecord.Expanded = False
      xRecord.Item(0).Icon = Bandeira.Verde
   End If
   
   '* Determina Grupo
   With mvarMe.GrdFecharDia
      For i = QtdIni To .Records.Count - 1
         Set xSubRec = .Records(i)
         With xSubRec.Item(0)
            If .GroupCaption = sGrupo Then
               .GroupCaption = ""
               .GroupCaption = .GroupCaption & StrZero(QtdTotal, IIf(QtdTotal = 0, 0, 2)) & " Venda" & IIf(QtdTotal = 1, "", "s")
               .GroupCaption = .GroupCaption & Space(16) & " - " & StrZero(QtdErro, IIf(QtdErro = 0, 0, 2)) & " Erro" & IIf(QtdErro = 1, "", "s")
               .GroupCaption = .GroupCaption & " / " & StrZero(QtdAviso, IIf(QtdAviso = 0, 0, 2)) & " Aviso" & IIf(QtdAviso = 1, "", "s")

               .GroupPriority = 1
            End If
         End With
      Next
      nTotalErros = nTotalErros + QtdErro
      nTotalAvisos = nTotalAvisos + QtdAviso
      .Populate
      Call SetTag(mvarMe.GrdFecharDia, "Erro" & (QtdIni + 2), IIf(QtdErro + QtdAviso = 0, False, True))
   End With
End Sub
Private Sub FecharVendas(pData As String)
   With mvarMe.GrdFecharDia
   End With
End Sub

Private Function FecharAtendimentos(pData As String) As Boolean
   Dim Sql        As String
   Dim MyRs       As Object
   Dim TB_Atend   As Object
   Dim Queries    As Collection
   Dim CollEvent  As Collection
   
   Dim n          As Variant
   Dim oEvent     As Object
            
   Set Queries = New Collection
   Set CollEvent = New Collection
   
   Sql = "Select IDLOJA, IDATENDIMENTO, IDEVENTO" & vbNewLine
   Sql = Sql & " From OATENDIMENTO" & vbNewLine
   Sql = Sql & " Where IDLOJA=" & SqlNum(Sys.Propriedades("IDLOJA")) & vbNewLine
   Sql = Sql & " And Day(DTATEND)=" & Day(pData) & vbNewLine
   Sql = Sql & " And Month(DTATEND)=" & Month(pData) & vbNewLine
   Sql = Sql & " And Year(DTATEND)=" & Year(pData) & vbNewLine
   Sql = Sql & " And SITATEND <> " & SqlStr(gSITATEND_Excluido) & vbNewLine
   Sql = Sql & " And SITATEND <> " & SqlStr(gSITATEND_Fechado) & vbNewLine
   If mvarSys.xDb.Abretabela(Sql, MyRs) Then
      Sql = "Update OATENDIMENTO " & vbNewLine
      Sql = Sql & " Set SITATEND=" & SqlStr(gSITATEND_Fechado) & vbNewLine
      Sql = Sql & " , ALTERSTAMP=1" & vbNewLine
      Sql = Sql & " , TIMESTAMP=GetDate()" & vbNewLine
      Sql = Sql & " Where IDLOJA=" & SqlNum(Sys.Propriedades("IDLOJA")) & vbNewLine
      Sql = Sql & " And Day(DTATEND)=" & Day(pData) & vbNewLine
      Sql = Sql & " And Month(DTATEND)=" & Month(pData) & vbNewLine
      Sql = Sql & " And Year(DTATEND)=" & Year(pData) & vbNewLine
      Sql = Sql & " And SITATEND <> " & SqlStr(gSITATEND_Excluido)
      Queries.Add Sql
      While Not MyRs.EOF
         CollEvent.Add MyRs("IDEVENTO").Value
         MyRs.MoveNext
      Wend
   Else
      If Queries.Count > 0 Then
         Dim nRec As ReportRecord
         For Each nRec In mvarMe.GrdFecharDia.Records
            If GetTag(nRec.Item(2).Tag, "TABELA") = "OEVENTOAGENDA" Then
               CollEvent.Add xVal(Mid(nRec.Item(2).Value, Len("Evento ") + 1, 6))
            End If
         Next
      End If
   End If
   If Queries.Count = 0 Then
      FecharAtendimentos = True
   Else
      If mvarSys.xDb.executa(Queries) Then
         FecharAtendimentos = True
         For Each n In CollEvent
            DoEvents
            If Not mvarCalControl Is Nothing Then
               Set oEvent = mvarCalControl.DataProvider.GetEvent(n)
            End If
            If Not oEvent Is Nothing Then
               oEvent.CustomProperties("EXIBEMSG") = False
               If oEvent.CustomProperties("SITATEND") <> "10" And oEvent.CustomProperties("FLGCANCELADO") <> "1" Then
                  Call SetCustomEvent(oEvent, mvarSys, pSITATEND:=gSITATEND_Fechado)
                  'Call mvarNGCal.GSetCustomEvent(oEvent, mvarsys, pSITATEND:=gSITATEND_Fechado)
                  Call RefreshEvent(mvarSys, mvarCalControl, oEvent, False)
                   mvarCalControl.Populate
                   mvarCalControl.RedrawControl
               End If
               oEvent.CustomProperties("EXIBEMSG") = True
            End If
            Set oEvent = Nothing
         Next
      End If
   End If
End Function
Private Function InformarSocios(pData As String) As Boolean
   Dim Ng      As NG_UTILITARIO
   Dim bDebug  As Boolean
   Dim bOk     As Boolean
   
   If ExisteArquivo("c:\debug.txt") Then bDebug = True
   Set Ng = New NG_UTILITARIO
   With Ng
      Set .Sys = mvarSys
      If bDebug Then Call MsgBox("Envia Fechamento")
      bOk = .F_EnviarFechamento(pData, False)
      If bOk Then
         If bDebug Then Call MsgBox("Envia RPS")
         Call .F_EnviarRPS(pData, False)
         
         If bDebug Then Call MsgBox("Envia Faturamento")
         Call .F_EnviarFaturamento(pData, False)
         
         If bDebug Then Call MsgBox("Envia ocorrencias")
         Call .F_EnviarOcorrencias(pData, False)
   
         If bDebug Then Call MsgBox("Envia inventario")
         Call .F_EnviarInventario(pData, False)
      End If
   End With
   Set Ng = Nothing
   InformarSocios = bOk
End Function
Public Function GerarBackup() As Boolean
   Dim Ng As New NG_UTILITARIO
   
   Dim Sql        As String
   Dim sBakName   As String
   Dim sPathBk    As String
   Dim dFile      As Date
   Dim dEvento    As Date
   Dim dAgora     As Date
   Dim bOk        As Boolean

Dim bDebug As Boolean
If ExisteArquivo("c:\debug.txt") Then bDebug = True
   
   sBakName = ""
   sBakName = Replace(mvarSys.Propriedades("NMLOJA"), " ", "")
   If InStr(sBakName, " ") <> 0 Then
      sBakName = Trim(Mid(sBakName, 1, InStr(sBakName, " ")))
   End If
   sBakName = mvarSys.Propriedades("IDCOLIGADA") & sBakName
   If sPathBk = "" Then sPathBk = ProcuraArquivo(Environ("programfiles") & "\Microsoft SQL Server\MSSQL\", sBakName & ".bak")
   If sPathBk = "" Then sPathBk = ProcuraArquivo(Environ("programfiles") & "\Microsoft SQL Server\MSSQL.1\", sBakName & ".bak")
   If sPathBk = "" Then sPathBk = ProcuraArquivo(Environ("programfiles") & "\Microsoft SQL Server\MSSQL.2\", sBakName & ".bak")
   If sPathBk = "" Then sPathBk = ProcuraArquivo(Environ("programfiles") & "\Microsoft SQL Server\MSSQL.3\", sBakName & ".bak")
   If ExisteArquivo(ResolvePathName(sPathBk) & sBakName & ".bak") Then
      dFile = FileDateTime(ResolvePathName(sPathBk) & sBakName & ".bak")

      Sql = "Select Max(EndDateTime) As [Evento]" & vbNewLine
      Sql = Sql & " From OEVENTOAGENDA " & vbNewLine
      Sql = Sql & " Where Year(StartDateTime)=" & Year(mvarSys.xDb.SysDate()) & vbNewLine
      Sql = Sql & " And Month(StartDateTime)=" & Month(mvarSys.xDb.SysDate()) & vbNewLine
      Sql = Sql & " And Day(StartDateTime)=" & Day(mvarSys.xDb.SysDate()) & vbNewLine
      If mvarSys.xDb.Abretabela(Sql) Then
         If Trim(mvarSys.xDb.RsAux("Evento") & "") = "" Then
            dEvento = mvarSys.xDb.SysDate()
         Else
            dEvento = mvarSys.xDb.RsAux("Evento")
         End If
      Else
         dEvento = mvarSys.xDb.SysDate()
      End If
      bOk = False
      If Not Between(DateDiff("h", dEvento, dFile), 0, 1) Then
         dAgora = mvarSys.xDb.SysDate(3)
         If Not Between(DateDiff("h", dFile, dAgora), 0, 1) Then
            bOk = True
         End If
      End If
   End If
   If bOk Then
      Set Ng.Sys = mvarSys
      GerarBackup = Ng.F_Backup(False, False)
      Set Ng = Nothing
   Else
      GerarBackup = True
   End If
If bDebug Then MsgBox ("7")
End Function
Public Function ArmazenarBackup() As Boolean
   Dim Ng As New NG_UTILITARIO
   Set Ng.Sys = mvarSys
   ArmazenarBackup = Ng.F_ExportarBackup
   Set Ng = Nothing
End Function
Private Sub EscreverFechamento(pDate As String)
   On Error Resume Next
   Dim Ng As NG_UTILITARIO
   
   Set Ng = New NG_UTILITARIO
   Set Ng.Sys = mvarSys
   Call Ng.GetHtmleMailDiario(pDate)
End Sub
Private Sub mvarMe_MvwDTFECHARDateClick(ByVal DateClicked As Date)
   If DateClicked <> mvarMe.CmbDTFECHAR.Value Then
      Screen.MousePointer = vbHourglass
      mvarMe.CmbDTFECHAR.Value = DateClicked
      Call LimparGrid
      Call ExibeValoresDia(Format(DateClicked, "dd/mm/yyyy"))
      Call VerificarDia(Format(DateClicked, "dd/mm/yyyy"))
      Screen.MousePointer = vbDefault
   End If
End Sub

Private Sub mvarMe_TaskDialogButtonClicked(ByVal Id As Long, CloseDialog As Variant)
   Select Case Id
      Case xtpTaskButtonOk, xtpTaskButtonYes '1,2
         If nTotalErros = 0 Then
            'mvarMe.TaskDialog.ProgressBarStyle = xtpTaskProgressBarMarquee
            'mvarMe.TaskDialog.StartProgressBarMarquee True, 10
            Screen.MousePointer = vbHourglass
            bExeFechamento = True
            With mvarMe.TaskFlood
                .Reset
               .WindowTitle = "Fechamento..."
               .MainInstructionText = "Em Andamento..."
               .ContentText = "Finalizando Dia"
               .AllowCancellation = False
               .SetProgressBarRange 1, 100
               .EnableCallbackTimer = True
               .MessageBoxStyle = True
               .ProgressBarStyle = xtpTaskProgressBarNormal 'xtpTaskProgressBarMarquee
               .ShowDialog
            End With
            Screen.MousePointer = vbDefault
            bExeFechamento = False
            CloseDialog = True
            
            'Call FinalizarFechamento
         End If
   End Select
End Sub
Private Sub FinalizarFechamento()
   Dim sDate      As String
   Dim i          As Long
   Dim sText      As String
   Dim sMsg       As String
   Dim nMouse     As Integer
   Dim bProducao  As Boolean
   Dim bGerouBak  As Boolean
      
   On Error Resume Next
Dim bDebug As Boolean
If ExisteArquivo("c:\debug.txt") Then bDebug = True
      
   nMouse = Screen.MousePointer
   Screen.MousePointer = vbHourglass
   
   bProducao = (UCase(mvarSys.xDb.Alias) = "PRODUCAO")
   
   sMsg = ""
   sDate = Format(mvarMe.CmbDTFECHAR.Value, "dd/mm/yyyy")

   'Call FecharVendas(sDate)
   DoEvents
   mvarMe.TaskFlood.ProgressBarPos = 5
   sText = "Fechando atendimentos..."
   mvarMe.TaskFlood.ContentText = sText
   If Not FecharAtendimentos(sDate) Then
      sMsg = sMsg & "- Fechamento de Atendimentos." & vbNewLine
   End If

   mvarMe.TaskFlood.ProgressBarPos = 30
   If xVal(mvarSys.GetParam("BACKUPDIA", Default:=1)) = 1 Then
      sText = "Gerando cópia de segurança..."
      mvarMe.TaskFlood.ContentText = sText
      
      If bProducao Then
         bGerouBak = GerarBackup
         If Not bGerouBak Then
            sMsg = sMsg & "- Cópia de Segurança." & vbNewLine
         Else
         End If
      Else
      End If

   Else
      bGerouBak = False
   End If
   
   If mvarMe.TaskDialog.VerifyCheckState Then
      mvarMe.TaskFlood.ProgressBarPos = 45
      sText = "Enviando fechamento para os sócios..."
      mvarMe.TaskFlood.ContentText = sText
      If Not InformarSocios(sDate) Then
         sMsg = sMsg & "- Envio de Informação." & vbNewLine
         mvarMe.Timer.Enabled = True
      End If
      If bProducao And bGerouBak Then
         sText = "Armazenando cópia de segurança..."
         mvarMe.TaskFlood.ContentText = sText
         mvarMe.TaskFlood.ProgressBarPos = 90
         If Not ArmazenarBackup Then
            If Trim(sMsg) <> "" Then
               sMsg = sMsg & "- Armazenamento da Cópia de Segurança." & vbNewLine
            End If
         End If
      End If
   Else
      sText = "Escrevendo fechamento..."
      mvarMe.TaskFlood.ContentText = sText
      Call EscreverFechamento(sDate)
   End If
   If Trim(sMsg) <> "" And Not mvarMe.Timer.Enabled Then
      sMsg = "Foi detectado problemas nas operações abaixo:" & vbNewLine & vbNewLine & sMsg
      Call ExibirInformacao(sMsg)
   End If
   Screen.MousePointer = nMouse
End Sub
Private Function CalculaValorAtend(pIDLOJA As Integer, pIDATEND As Long, Optional ByRef pQTDSERV As Double = 0, Optional ByRef pQTDPROD As Double = 0) As Double
   Dim Sql     As String
   Dim RsAux   As Object
   Dim nValor  As Double
   Dim i As Integer
   
   '* Serviço
   Sql = "Select S.IDSESSAO, P2.IDPROD, IsNull(P2.VLVENDA,0) As [VLVENDA]"
   Sql = Sql & ", TP.IDTPSERVICO, TP.IDTPTRATAMENTO, TP.IDAREA"
   Sql = Sql & " From OATENDIMENTO A "
   Sql = Sql & " Join OSESSAO          S On A.IDLOJA =S.IDLOJA   And A.IDATENDIMENTO=S.IDATENDIMENTO"
   Sql = Sql & " Join OTRATAMENTO_PROD TP On TP.IDLOJA=S.IDLOJA  And (TP.IDTPSERVICO=S.IDTPSERVICO Or IsNull(TP.IDTPSERVICO, 0)=0) And (TP.IDTPTRATAMENTO=S.IDTPTRATAMENTO Or IsNull(TP.IDTPTRATAMENTO, 0)=0) And (TP.IDAREA=S.IDAREA Or IsNull(TP.IDAREA, 0)=0)"
   Sql = Sql & " Join SPRODUTO         P2 On TP.IDLOJA=P2.IDLOJA And TP.IDPROD=P2.IDPROD And P2.EVENDA=1"
   Sql = Sql & " Where A.IDLOJA=" & SqlNum(pIDLOJA)
   Sql = Sql & " And A.IDATENDIMENTO=" & SqlNum(pIDATEND)
   Sql = Sql & " Order By S.IDSESSAO, TP.IDAREA Desc, TP.IDTPTRATAMENTO Desc, TP.IDTPSERVICO Desc"
   If mvarSys.xDb.Abretabela(Sql, RsAux) Then
      pQTDSERV = 0
      nValor = 0
      i = 0
      RsAux.MoveFirst
      While Not RsAux.EOF
         If i <> RsAux("IDSESSAO") Then
            i = RsAux("IDSESSAO")
            pQTDSERV = pQTDSERV + 1
            nValor = nValor + (xVal(RsAux("VLVENDA") & ""))
         End If
         RsAux.MoveNext
      Wend
   End If
   '**************
   '* PRODUTO
   Sql = "Select P.IDPROD, IsNull(AP.QTDCOMPRA,0) As [QTDCOMPRA], IsNull(P.VLVENDA,0) As [VLVENDA]"
   Sql = Sql & " From OATENDIMENTO  A "
   Sql = Sql & " Left Join OATENDIMENTO_PRODUTO AP On A.IDLOJA =AP.IDLOJA  And A.IDATENDIMENTO=AP.IDATENDIMENTO"
   Sql = Sql & " Left Join SPRODUTO              P On AP.IDLOJA=P.IDLOJA And AP.IDPROD=P.IDPROD   And P.EVENDA=1"
   Sql = Sql & " Where A.IDLOJA=" & SqlNum(pIDLOJA)
   Sql = Sql & " And A.IDATENDIMENTO=" & SqlNum(pIDATEND)
   Sql = Sql & " And Not P.IDPROD is Null"
   If mvarSys.xDb.Abretabela(Sql, RsAux) Then
      RsAux.MoveFirst
      While Not RsAux.EOF
         pQTDPROD = pQTDPROD + xVal(RsAux("QTDCOMPRA") & "")
         nValor = nValor + (xVal(RsAux("VLVENDA") & "") * xVal(RsAux("QTDCOMPRA") & ""))
         RsAux.MoveNext
      Wend
   End If
   
   CalculaValorAtend = nValor
End Function

Private Sub mvarMe_TaskFloodConstructed()
'   mvarMe.TaskFlood.StartProgressBarMarquee True, 80
End Sub
'Private Sub mvarMe_TaskDialogTimer(ByVal MilliSeconds As Long, Reset As Variant)
'End Sub
Private Sub mvarMe_TaskFloodTimer(ByVal MilliSeconds As Long, Reset As Variant)
'    If 5000 >= MilliSeconds Then
'        Dim text As String
'        text = "Sending... " & Int(MilliSeconds / 5000 * 100) & "%"
'        taskDiaogFeedback.MainInstructionText = text
'
'        taskDiaogFeedback.ProgressBarPos = MilliSeconds
'        taskDiaogFeedback.SetProgressBarRange 0, 5000
'    Else
'        taskDiaogFeedback.MainInstructionText = "Thanks for the feedback!"
'        taskDiaogFeedback.ContentText = "Our developers will get right on that..."
'        taskDiaogFeedback.ProgressBarPos = 5000
'        taskDiaogFeedback.EnableCallbackTimer = False
'    End If
   If bExeFechamento Then
      bExeFechamento = False
      On Error Resume Next
      With mvarMe.TaskFlood
         
         Call FinalizarFechamento
         .ContentText = ""
         .MainInstructionText = "Operação realizada com sucesso!"
         If mvarMe.Timer.Enabled Then
            .ContentText = "O Dia foi fechado corretamente, porém o e-mail não foi enviado aos sócios."
         End If
         
         .ProgressBarPos = 100
         .ProgressBarStyle = xtpTaskProgressBarNone
         .EnableCallbackTimer = False
         '.Reset
      End With
   End If
End Sub

Private Sub mvarMe_TimerTimer()
   Dim sDate As String
   Static nMin As Integer
   
   DoEvents
   mvarMe.Caption = "Fechamento do Dia - Tentando enviar aos sócios... " & Now()
   mvarMe.Refresh
   sDate = Format(mvarMe.CmbDTFECHAR.Value, "dd/mm/yyyy")
   If nMin <> Format(Now(), "n") Then
      nMin = Format(Now(), "n")
      If InformarSocios(sDate) Then
        mvarMe.Timer.Enabled = False
        mvarMe.Caption = "Fechamento do Dia - Enviado em " & Now()
        mvarMe.TaskFlood.FooterText = "[" & Format(Now(), "hh:mm:ss") & "] Fechamento enviado!"
      End If
   Else
      On Error Resume Next
      mvarMe.TaskFlood.ContentText = "O Dia foi fechado corretamente, porém o e-mail não foi enviado aos sócios."
      mvarMe.TaskFlood.FooterIcon = xtpTaskIconWarning
      mvarMe.TaskFlood.FooterText = "[" & Format(Now(), "hh:mm:ss") & "] O sistema continuará tentando o envio."
      
   End If
End Sub

Private Sub mvarMe_wndDatePickerSelectionChanged()
   Dim DateBegin As Date
   Dim DateEnd As Date
   
   DateBegin = mvarMe.wndDatePicker.Selection.Blocks(0).DateBegin
   DateEnd = mvarMe.wndDatePicker.Selection.Blocks(0).DateEnd
   If Not IsDate(DateBegin) Then DateBegin = Now()
   If Not IsDate(DateEnd) Then DateEnd = Now()
   
   If DateBegin <> mvarMe.CmbDTFECHAR.Value Or DateBegin <> DateEnd Then
      Screen.MousePointer = vbHourglass
      mvarMe.MousePointer = vbHourglass
      mvarMe.CmbDTFECHAR.Value = DateBegin
      Call LimparGrid
      Call ExibeValoresDia(Format(DateBegin, "dd/mm/yyyy"), Format(DateEnd, "dd/mm/yyyy"))
'      If Format(DateBegin, "dd/mm/yyyy") = Format(mvarSys.xDb.SysDate(), "dd/mm/yyyy") Then
'         Call VerificarDia(Format(DateBegin, "dd/mm/yyyy"))
'      Else
         mvarMe.GrdFecharDia.Enabled = False
'      End If
      Screen.MousePointer = vbDefault
      mvarMe.MousePointer = vbDefault
   End If
End Sub
