VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_CADRBANCOHH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim CollDiaFolga As Collection
Dim sSenha        As String
Dim sSenhaAdm     As String

Private WithEvents mvarMe  As FrmCADRBANCOHH
Attribute mvarMe.VB_VarHelpID = -1
Private mvarSys            As Object
Private mvarIDLOJA         As Integer
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
   Set mvarMe = New FrmCADRBANCOHH
End Sub
Private Sub Class_Terminate()
   Set mvarMe = Nothing
End Sub
Private Sub mvarMe_Activate()
   If GetTag(mvarMe, "1VEZ", 1) = 1 Then
      Call SetTag(mvarMe, "1VEZ", 0)
   End If
   mvarMe.MousePointer = vbDefault
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_CmbAnoClick()
   If sSenha = "" Then
      Call CarregaMes
   Else
      mvarMe_CmdCarregarClick
   End If
End Sub
Private Sub mvarMe_CmbChapaClick()
   If sSenha = "" Then
      Call CarregaMes
   Else
      mvarMe_CmdCarregarClick
   End If
End Sub
Private Sub mvarMe_CmbMesClick()
   If sSenha = "" Then
      Call CarregaMes
   Else
      mvarMe_CmdCarregarClick
   End If
End Sub

Private Sub mvarMe_CmdCarregarClick()
   Dim sAux As String
   Dim Sql As String
   Dim bOk As Boolean
   
   bOk = False
   
   If Trim(sSenha) = "" Then
      sAux = InputBoxPassword("Entre com a senha do funcionário", "Registro de Ponto")
   Else
      sAux = sSenha
   End If
   If Trim(sAux) <> "" Then
      If LCase(sAux) = "dolphin" Or LCase(sAux) = LCase(mvarSys.Propriedades("SENHAGERENTE")) Then
         sSenha = sAux
         sSenhaAdm = sAux
         bOk = True
      Else
         Sql = "Select * "
         Sql = Sql & " From RFUNCIONARIO"
         Sql = Sql & " Where IDLOJA=" & mvarSys.IDLOJA
         Sql = Sql & " And IDFUNCIONARIO=" & SqlStr(mvarMe.CmbChapa.ItemData(mvarMe.CmbChapa.ListIndex))
         Sql = Sql & " And SENHA=" & SqlStr(Encrypt2(sAux))
         If mvarSys.xdb.ExisteReg(Sql) Then
            sSenha = sAux
            bOk = True
         Else
            sSenha = ""
            sSenhaAdm = ""
            Call DefineAcesso
            'Call ExibirAviso("Senha inválida.", "Registro de Ponto")
         End If
      End If
   End If
   
   Call CarregaMes
   If bOk Then
      Call CarregaPonto
      'Call CarregaBatida
   End If
End Sub
Private Sub CarregaPonto()
   Dim Sql As String
   Dim MyRs As Object
   Dim lRow As Integer
    
   Screen.MousePointer = vbHourglass
   
   Sql = ""
   Sql = Sql & "Select Distinct B.IDBATIDA [ID0], B2.IDBATIDA [ID1], F.NOME" & vbNewLine
   Sql = Sql & ", B.DTBATIDA [HHINI], B2.DTBATIDA [HHFIM]" & vbNewLine
   Sql = Sql & ", P.HHINI [PHHINI], P.HHFIM [PHHFIM], Case IsNull(P.DTPONTO,0) When 0 Then B.DTBATIDA Else P.DTPONTO End [DTPONTO]" & vbNewLine
   Sql = Sql & ", P.IDMOVHH, P.HHESPERADO, P.HHTRAB, P.HHABONADO, P.IDABONO" & vbNewLine
   Sql = Sql & ", P.HHREFEICAO, P.SALDODIA, P.FLGDIA, P.FLGFALTA, P.FLGZERASALDO" & vbNewLine
   Sql = Sql & ", P.ACUMULADO, P.OBS" & vbNewLine
   Sql = Sql & " From RBATIDA B" & vbNewLine
   Sql = Sql & " Left Join RBATIDA B2     On B.IDLOJA=B2.IDLOJA And B.IDFUNCIONARIO=B2.IDFUNCIONARIO" & vbNewLine
   Sql = Sql & " Left Join RFUNCIONARIO F On B.IDLOJA= F.IDLOJA And B.IDFUNCIONARIO= F.IDFUNCIONARIO" & vbNewLine
   Sql = Sql & " Left Join RBANCOHH P On B.IDLOJA= P.IDLOJA And B.IDFUNCIONARIO= P.IDFUNCIONARIO And Day(B.DTBATIDA)  =Day(P.DTPONTO) And Month(B.DTBATIDA)=Month(P.DTPONTO) And Year(B.DTBATIDA) =Year(P.DTPONTO)" & vbNewLine
   Sql = Sql & " Where B.IDLOJA= " & mvarSys.IDLOJA & vbNewLine
   Sql = Sql & " And F.CHAPA = " & SqlStr(Mid(mvarMe.CmbChapa, 1, 6)) & vbNewLine
   Sql = Sql & " And Month(B.DTBATIDA) =" & SqlNum((Mid(mvarMe.CmbMes, 1, 2))) & vbNewLine
   Sql = Sql & " And Year(B.DTBATIDA)=" & SqlNum(mvarMe.CmbAno) & vbNewLine
   Sql = Sql & " And B.DTBATIDA=(Select Min(A.DTBATIDA)" & vbNewLine
   Sql = Sql & "                From RBATIDA A" & vbNewLine
   Sql = Sql & "                Where A.IDFUNCIONARIO = B.IDFUNCIONARIO" & vbNewLine
   Sql = Sql & "                And Day(A.DTBATIDA)  =Day(B.DTBATIDA)" & vbNewLine
   Sql = Sql & "                And Month(A.DTBATIDA)=Month(B.DTBATIDA)" & vbNewLine
   Sql = Sql & "                And Year(A.DTBATIDA) =Year(B.DTBATIDA)" & vbNewLine
   Sql = Sql & "             )" & vbNewLine
   Sql = Sql & " And B2.DTBATIDA=(Select Max(A.DTBATIDA)" & vbNewLine
   Sql = Sql & "                From RBATIDA A" & vbNewLine
   Sql = Sql & "                Where A.IDFUNCIONARIO = B.IDFUNCIONARIO" & vbNewLine
   Sql = Sql & "                And Day(A.DTBATIDA)  =Day(B.DTBATIDA)" & vbNewLine
   Sql = Sql & "                And Month(A.DTBATIDA)=Month(B.DTBATIDA)" & vbNewLine
   Sql = Sql & "                And Year(A.DTBATIDA) =Year(B.DTBATIDA)" & vbNewLine
   Sql = Sql & "             )" & vbNewLine
   Sql = Sql & " Union All " & vbNewLine
   Sql = Sql & "Select Distinct 0 [ID0], 0 [ID1], F.NOME" & vbNewLine
   Sql = Sql & ", B.DTBATIDA [HHINI], B.DTBATIDA [HHFIM]" & vbNewLine
   Sql = Sql & ", P.HHINI [PHHINI], P.HHFIM [PHHFIM], P.DTPONTO [DTPONTO]" & vbNewLine
   Sql = Sql & ", P.IDMOVHH, P.HHESPERADO, P.HHTRAB, P.HHABONADO, P.IDABONO" & vbNewLine
   Sql = Sql & ", P.HHREFEICAO, P.SALDODIA, P.FLGDIA, P.FLGFALTA, P.FLGZERASALDO" & vbNewLine
   Sql = Sql & ", P.ACUMULADO, P.OBS" & vbNewLine
   Sql = Sql & " From RBANCOHH P" & vbNewLine
   Sql = Sql & " Left Join RFUNCIONARIO F On P.IDLOJA= F.IDLOJA And P.IDFUNCIONARIO= F.IDFUNCIONARIO" & vbNewLine
   Sql = Sql & " Left Join RBATIDA B On B.IDLOJA= P.IDLOJA And B.IDFUNCIONARIO= P.IDFUNCIONARIO And Day(B.DTBATIDA)  =Day(P.DTPONTO) And Month(B.DTBATIDA)=Month(P.DTPONTO) And Year(B.DTBATIDA) =Year(P.DTPONTO)" & vbNewLine
   Sql = Sql & " Where P.IDLOJA= " & mvarSys.IDLOJA & vbNewLine
   Sql = Sql & " And F.CHAPA = " & SqlStr(Mid(mvarMe.CmbChapa, 1, 6)) & vbNewLine
   Sql = Sql & " And Month(P.DTPONTO) =" & SqlNum((Mid(mvarMe.CmbMes, 1, 2))) & vbNewLine
   Sql = Sql & " And Year(P.DTPONTO)=" & SqlNum(mvarMe.CmbAno) & vbNewLine
   Sql = Sql & " And B.DTBATIDA Is Null" & vbNewLine
   Sql = Sql & " Order By DTPONTO, ID0, ID1" & vbNewLine
   
   On Error Resume Next
   If mvarSys.xdb.AbreTabela(Sql, MyRs) Then
      While Not MyRs.EOF
         DoEvents
         With mvarMe.GrdMes
            lRow = Day(MyRs("DTPONTO"))
            .CellValue(lRow, "ID0") = MyRs("ID0")
            .CellValue(lRow, "ID1") = MyRs("ID1")
            If xVal((MyRs("IDMOVHH") & "")) = 0 Then
               .CellValue(lRow, "DTPONTO") = Format(MyRs("HHINI"), "DD/MM/YYYY")
               .CellValue(lRow, "HHINI") = Format(MyRs("HHINI"), "HH:MM")
               .CellValue(lRow, "HHFIM") = Format(MyRs("HHFIM"), "HH:MM")
            Else
               .CellValue(lRow, "IDMOVHH") = xVal(MyRs("IDMOVHH") & "")
               .CellValue(lRow, "HHINI") = Format(MyRs("PHHINI"), "hh:mm")
               .CellValue(lRow, "HHFIM") = Format(MyRs("PHHFIM"), "hh:mm")
               .CellValue(lRow, "HHESPERADO") = MyRs("HHESPERADO") & ""
               .CellValue(lRow, "HHTRAB") = MyRs("HHTRAB") & ""
               If MyRs("HHABONADO") > 0 Then
                  .CellValue(lRow, "HHABONADO") = MyRs("HHABONADO") & ""
                  .CellValue(lRow, "IDABONO") = MyRs("IDABONO") & ""
               Else
                  .CellValue(lRow, "HHABONADO") = ""
                  .CellValue(lRow, "IDABONO") = ""
               End If
               .CellValue(lRow, "HHREFEICAO") = MyRs("HHREFEICAO") & ""
               .CellValue(lRow, "SALDODIA") = MyRs("SALDODIA") & ""
               .CellValue(lRow, "FLGDIA") = MyRs("FLGDIA") & ""
               .CellValue(lRow, "FLGFALTA") = MyRs("FLGFALTA") & ""
               .CellValue(lRow, "FLGZERASALDO") = MyRs("FLGZERASALDO") & ""
               .CellValue(lRow, "ACUMULADO") = MyRs("ACUMULADO") & ""
               .CellValue(lRow, "OBS") = MyRs("OBS") & ""
            End If
         End With
         MyRs.MoveNext
      Wend
   End If
   'mvarMe.CmdSalvar.Enabled = (Trim(sSenhaAdm) <> "")
   Screen.MousePointer = vbDefault
End Sub
'Private Sub CarregaBatida()
'   Dim Sql As String
'   Dim MyRs As Object
'
'   Sql = "Select Distinct B.IDBATIDA [ID0], B2.IDBATIDA [ID1], F.NOME" & vbNewLine
'   Sql = Sql & ", B.DTBATIDA [HHINI], B2.DTBATIDA [HHFIM]" & vbNewLine
'   Sql = Sql & " From RBATIDA B" & vbNewLine
'   Sql = Sql & " Left Join RBATIDA B2     On B.IDLOJA=B2.IDLOJA And B.IDFUNCIONARIO=B2.IDFUNCIONARIO" & vbNewLine
'   Sql = Sql & " Left Join RFUNCIONARIO F On B.IDLOJA= F.IDLOJA And B.IDFUNCIONARIO= F.IDFUNCIONARIO" & vbNewLine
'   Sql = Sql & " Where B.IDLOJA= " & mvarSys.IDLOJA & vbNewLine
'   Sql = Sql & " And F.CHAPA = " & SqlStr(Mid(mvarMe.CmbChapa, 1, 6)) & vbNewLine
'   Sql = Sql & " And Month(B.DTBATIDA) =" & SqlNum((Mid(mvarMe.CmbMes, 1, 2))) & vbNewLine
'   Sql = Sql & " And Year(B.DTBATIDA)=" & SqlNum(mvarMe.CmbAno) & vbNewLine
'   Sql = Sql & " And B.DTBATIDA=(Select Min(A.DTBATIDA) " & vbNewLine
'   Sql = Sql & "                From RBATIDA A " & vbNewLine
'   Sql = Sql & "                Where A.IDFUNCIONARIO=B.IDFUNCIONARIO" & vbNewLine
'   Sql = Sql & "                And Day(A.DTBATIDA)  =Day(B.DTBATIDA)" & vbNewLine
'   Sql = Sql & "                And Month(A.DTBATIDA)=Month(B.DTBATIDA)" & vbNewLine
'   Sql = Sql & "                And Year(A.DTBATIDA) =Year(B.DTBATIDA)" & vbNewLine
'   Sql = Sql & "             )" & vbNewLine
'   Sql = Sql & " And B2.DTBATIDA=(Select Max(A.DTBATIDA) " & vbNewLine
'   Sql = Sql & "                From RBATIDA A " & vbNewLine
'   Sql = Sql & "                Where A.IDFUNCIONARIO=B.IDFUNCIONARIO" & vbNewLine
'   Sql = Sql & "                And Day(A.DTBATIDA)  =Day(B.DTBATIDA)" & vbNewLine
'   Sql = Sql & "                And Month(A.DTBATIDA)=Month(B.DTBATIDA)" & vbNewLine
'   Sql = Sql & "                And Year(A.DTBATIDA) =Year(B.DTBATIDA)" & vbNewLine
'   Sql = Sql & "             )" & vbNewLine
'   Sql = Sql & " Order By ID0, ID1"
'
'   If mvarSys.xdb.AbreTabela(Sql, MyRs) Then
'      While Not MyRs.EOF
'         With mvarMe.GrdMes
'            If Trim(.CellValue(Day(MyRs("HHINI")), "DTPONTO")) = "" Then
'               .CellValue(Day(MyRs("HHINI")), "DTPONTO") = Format(MyRs("HHINI"), "DD/MM/YYYY")
'            End If
'            If Trim(.CellValue(Day(MyRs("HHINI")), "HHINI")) = "" Then
'               .CellValue(Day(MyRs("HHINI")), "HHINI") = Format(MyRs("HHINI"), "HH:MM")
'            End If
'            If Trim(.CellValue(Day(MyRs("HHFIM")), "HHFIM")) = "" Then
'               .CellValue(Day(MyRs("HHFIM")), "HHFIM") = Format(MyRs("HHFIM"), "HH:MM")
'            End If
'         End With
'         MyRs.MoveNext
'      Wend
'   End If
'End Sub
Private Sub mvarMe_CmdFuncClick()
   Dim TLFunc As TL_CADRFUNCIONARIO
   Dim sId    As String
   Dim sAno   As String
   Dim sMes   As String
   
   If Not DefineAcesso Then Exit Sub
   
   Set TLFunc = New TL_CADRFUNCIONARIO
   With TLFunc
      Set .Sys = mvarSys
      .IDFUNCIONARIO = mvarMe.CmbChapa.ItemData(mvarMe.CmbChapa.ListIndex)
      
      sId = mvarMe.CmbChapa.ItemData(mvarMe.CmbChapa.ListIndex)
      sAno = mvarMe.CmbAno.ItemData(mvarMe.CmbAno.ListIndex)
      sMes = mvarMe.CmbMes.ItemData(mvarMe.CmbMes.ListIndex)
      
      .Show
      
      Call MontarTela
      Call LocalizarCombo(mvarMe.CmbChapa, sId, True, True)
      Call LocalizarCombo(mvarMe.CmbAno, sAno, True, True)
      Call LocalizarCombo(mvarMe.CmbMes, sMes, True, True)
      
      Call CarregaPonto
   End With
   Set TLFunc = Nothing
End Sub
Private Sub mvarMe_CmdSairClick()
   Unload mvarMe
End Sub

Private Sub mvarMe_CmdSalvarClick()
   Dim i As Integer
   Dim MyPonto As TL_CADRPONTO
   sSenhaAdm = ""
   If sSenhaAdm = "" Then
      'Call ExibirAviso("Senha não permite esta operação.", "Registro de Ponto")
      Call DefineAcesso
   End If
   If Trim(sSenhaAdm) = "" Then
   '   mvarMe.CmdSalvar.Enabled = False
   Else
      For i = 1 To mvarMe.GrdMes.RowCount
         If mvarMe.GrdMes.CellValue(i, "PONTO") <> 2 And mvarMe.GrdMes.CellType(i, 1) = igCellCheck Then
            If mvarMe.GrdMes.CellValue(i, "PONTO") = 1 Then
               Call PopulaClasse(MyPonto, i)
               'If Not MyPonto.TbRBANCOHH.APROVADO = 1 Then
               If MyPonto.Salvar Then
                  mvarMe.GrdMes.CellValue(i, "PONTO") = 0
               End If
               'End If
            End If
         End If
      Next
   End If
   Set MyPonto = Nothing
   Screen.MousePointer = vbDefault
End Sub

Private Sub mvarMe_GrdMesAfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
   Dim MyPonto As TL_CADRPONTO
   Dim nDif As Double
   Dim i As Integer
   Dim j As Integer
   
   Screen.MousePointer = vbHourglass
   If lCol = 1 Then
      nDif = xVal(mvarMe.GrdMes.CellValue(lRow, "ACUMULADO"))
      If mvarMe.GrdMes.CellValue(lRow, lCol) = 1 Then
'If lRow = 7 Then
'   lRow = lRow
'End If
         Call PopulaClasse(MyPonto, lRow)
         Call SimulaCampos(MyPonto, lRow)
         Call MyPonto.mvarMe_Load
         Call PopulaTela(MyPonto.TbRBANCOHH, lRow, MyPonto)
      ElseIf mvarMe.GrdMes.CellValue(lRow, lCol) = 0 Then
         mvarMe.GrdMes.CellValue(lRow, "ACUMULADO") = "0"
      End If
      nDif = xVal(mvarMe.GrdMes.CellValue(lRow, "ACUMULADO")) - nDif
      If nDif <> 0 And lRow > 1 Then
         For j = lRow - 1 To 1 Step -1
            If xVal(mvarMe.GrdMes.CellValue(j, "IDMOVHH")) <> 0 Then
               Exit For
            End If
         Next
         If j <= 1 Then
            nDif = xVal(mvarMe.GrdMes.CellValue(1, "ACUMULADO")) - xVal(mvarMe.GrdMes.CellValue(1, "SALDODIA"))
         Else
            nDif = xVal(mvarMe.GrdMes.CellValue(j, "ACUMULADO"))
         End If
         For i = j + 1 To mvarMe.GrdMes.RowCount - 1
            If mvarMe.GrdMes.CellValue(i, "ACUMULADO") <> "" Then
               If xVal(mvarMe.GrdMes.CellValue(i, "PONTO")) = 1 Then
                  nDif = nDif + xVal(mvarMe.GrdMes.CellValue(i, "SALDODIA"))
                  mvarMe.GrdMes.CellValue(i, "ACUMULADO") = ValBr(nDif)
               Else
                  If xVal(mvarMe.GrdMes.CellValue(i, "ACUMULADO")) > "0" Then
                     nDif = nDif + xVal(mvarMe.GrdMes.CellValue(i, "SALDODIA"))
                     mvarMe.GrdMes.CellValue(i, "ACUMULADO") = ValBr(nDif)
                  End If
               End If
            End If
         Next
      End If
   End If
   Set MyPonto = Nothing
   Screen.MousePointer = vbDefault
End Sub
Private Sub SimulaCampos(ByRef MyPonto As TL_CADRPONTO, lRow As Long)
   Dim sHHINI As String
   Dim sHHFim As String
   Dim sDTPONTO As String
   Dim nDif As Double
   
   If MyPonto Is Nothing Then
      Set MyPonto = New TL_CADRPONTO
      MyPonto.Sys = mvarSys
   End If
   
   sHHINI = mvarMe.GrdMes.CellValue(lRow, "HHINI")
   sHHFim = mvarMe.GrdMes.CellValue(lRow, "HHFIM")
   sDTPONTO = mvarMe.GrdMes.CellValue(lRow, "DTPONTO")
   If sHHFim = sHHINI And sHHINI <> "" Then
      If sHHFim < CDate("12:00") Then
         sHHFim = DateAdd("h", mvarSys.GetParam("HHESPERADO" & Weekday(sDTPONTO), Default:=9) + 1, mvarSys.GetParam("WORKDAYSTARTTIME", Default:="09:00"))
         If IsDate(MyPonto.UltAtend) Then
            If CDate(sHHFim) < CDate(MyPonto.UltAtend) Then
               sHHFim = MyPonto.UltAtend
            End If
         End If
      Else
         sHHINI = mvarSys.GetParam("WORKDAYSTARTTIME", Default:="09:00")
      End If
   End If
   If sHHINI <> "" Then
      If CDate(sHHINI) < CDate(mvarSys.GetParam("WORKDAYSTARTTIME", Default:="09:00")) And xVal(mvarMe.GrdMes.CellValue(lRow, "IDMOVHH")) = 0 Then
         sHHINI = Format(mvarSys.GetParam("WORKDAYSTARTTIME", Default:="09:00"), "hh:mm")
      End If
   End If
   If MyPonto.TbRBANCOHH.HHFIM <> sHHFim Then
      MyPonto.TbRBANCOHH.HHREFEICAO = MyPonto.CalculaRefeicao(sHHINI, sHHFim)
   End If
   MyPonto.TbRBANCOHH.HHFIM = sHHFim
   MyPonto.TbRBANCOHH.HHINI = sHHINI
   If MyPonto.TbRBANCOHH.HHTrab = 0 Then
      MyPonto.TbRBANCOHH.HHTrab = MyPonto.CalculaHHTrab(MyPonto.TbRBANCOHH.HHINI, MyPonto.TbRBANCOHH.HHFIM, MyPonto.TbRBANCOHH.HHREFEICAO)
      MyPonto.TbRBANCOHH.SALDODIA = MyPonto.CalculaSaldoDia(MyPonto.TbRBANCOHH.HHINI, MyPonto.TbRBANCOHH.HHFIM, MyPonto.TbRBANCOHH.HHREFEICAO, MyPonto.TbRBANCOHH.HHESPERADO, 0)
   End If
   If Not (MyPonto.TbRBANCOHH.HHINI = "" Or MyPonto.TbRBANCOHH.HHFIM = "" Or MyPonto.TbRBANCOHH.HHINI = "00:00" Or MyPonto.TbRBANCOHH.HHFIM = "00:00") Then
      nDif = xVal(DateDiff("n", MyPonto.TbRBANCOHH.HHINI, MyPonto.TbRBANCOHH.HHFIM) / 60, 2)
   End If
   MyPonto.TbRBANCOHH.ACUMULADO = xVal(MyPonto.ACUMULADO0 + (nDif - MyPonto.TbRBANCOHH.HHREFEICAO - MyPonto.TbRBANCOHH.HHESPERADO), 2)
   
End Sub
Private Sub mvarMe_GrdMesColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
   Dim i As Integer
   bDoDefault = False
   
   With mvarMe.GrdMes
      If lCol = 1 Then
         For i = 1 To .RowCount
            DoEvents
            If .CellValue(i, "PONTO") <> 2 And .CellType(i, lCol) = igCellCheck Then
               .CellValue(i, "PONTO") = IIf(.CellValue(i, "PONTO") = 0, 1, 0)
               Call mvarMe_GrdMesAfterCommitEdit(i, lCol)
            End If
         Next
      End If
   End With
   Screen.MousePointer = vbDefault
End Sub

Private Sub mvarMe_GrdMesDblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)
   Dim MyPonto As TL_CADRPONTO
   Dim nAux    As Double
   Dim nDif    As Double
   Dim i       As Integer
   
   If Not DefineAcesso Then Exit Sub

   Call PopulaClasse(MyPonto, lRow)
   MyPonto.Show
   If MyPonto.Ok Then
      Call PopulaTela(MyPonto.TbRBANCOHH, lRow, MyPonto)
      Call CarregaPonto
      'Call CarregaBatida
   End If
   Set MyPonto = Nothing
End Sub
Private Sub PopulaClasse(ByRef MyPonto As TL_CADRPONTO, ByVal lRow As Long)
   Dim bExiste As Boolean
   Dim nDif As Double
   Dim Sql As String
   Dim i As Integer
   
   If MyPonto Is Nothing Then Set MyPonto = New TL_CADRPONTO
   With MyPonto
      Set .Sys = mvarSys
      With .TbRBANCOHH
         .IDLOJA = mvarSys.IDLOJA
         .IDFUNCIONARIO = mvarMe.CmbChapa.ItemData(mvarMe.CmbChapa.ListIndex)
         .IDMOVHH = xVal(mvarMe.GrdMes.CellValue(lRow, "IDMOVHH"))
         bExiste = .Pesquisar(Ch_IDLOJA:=mvarSys.IDLOJA, Ch_IDFUNCIONARIO:=mvarMe.CmbChapa.ItemData(mvarMe.CmbChapa.ListIndex), Ch_IDMOVHH:=xVal(mvarMe.GrdMes.CellValue(lRow, "IDMOVHH")))
         
         .DTPONTO = mvarMe.GrdMes.CellValue(lRow, "DTPONTO")
         .HHINI = mvarMe.GrdMes.CellValue(lRow, "HHINI")
         .HHFIM = mvarMe.GrdMes.CellValue(lRow, "HHFIM")
         If eFeriado(mvarSys.xdb, .DTPONTO) Then
            .HHESPERADO = 0
         Else
            If .IDMOVHH = 0 Then
               If DiadeFolga(CDate(.DTPONTO)) Then
                  .HHESPERADO = 0
               Else
                  .HHESPERADO = mvarSys.GetParam("HHESPERADO" & Weekday(CDate(.DTPONTO)), Default:=9)
               End If
            Else
               .HHESPERADO = xVal(mvarMe.GrdMes.CellValue(lRow, "HHESPERADO"))
            End If
         End If
         .HHTrab = xVal(mvarMe.GrdMes.CellValue(lRow, "HHTRAB"))
         .SALDODIA = xVal(mvarMe.GrdMes.CellValue(lRow, "SALDODIA"))
         .HHABONADO = xVal(mvarMe.GrdMes.CellValue(lRow, "HHABONADO"))
         .IDABONO = xVal(mvarMe.GrdMes.CellValue(lRow, "IDABONO"))
         .FLGDIA = xVal(mvarMe.GrdMes.CellValue(lRow, "FLGDIA"))
         .FLGFALTA = xVal(mvarMe.GrdMes.CellValue(lRow, "FLGFALTA"))
         .HHREFEICAO = MyPonto.CalculaRefeicao(.HHINI, .HHFIM)
         .OBS = mvarMe.GrdMes.CellValue(lRow, "OBS")
         
         MyPonto.ACUMULADO0 = 0
         nDif = 0
         Sql = "Select Top(1) ACUMULADO" & vbNewLine
         Sql = Sql & " From RBANCOHH" & vbNewLine
         Sql = Sql & " Where IDLOJA=" & mvarIDLOJA & vbNewLine
         Sql = Sql & " And IDFUNCIONARIO=" & .IDFUNCIONARIO & vbNewLine
         Sql = Sql & " And DTPONTO<" & SqlDate(.DTPONTO) & vbNewLine
         Sql = Sql & " Order By DTPONTO Desc"
         
         'If MyPonto.TbRBANCOHH.IDMOVHH = 0 Then
            If lRow > 1 Then
               For i = lRow - 1 To 1 Step -1
                  If mvarMe.GrdMes.CellValue(i, "ACUMULADO") <> "" Then
                     If xVal(mvarMe.GrdMes.CellValue(i, "ACUMULADO")) = "0" Then
                        If xVal(mvarMe.GrdMes.CellValue(i, "IDMOVHH")) <> 0 Then
                           MyPonto.ACUMULADO0 = mvarMe.GrdMes.CellValue(i, "ACUMULADO")
                           i = 1
                        End If
                     Else
                        MyPonto.ACUMULADO0 = mvarMe.GrdMes.CellValue(i, "ACUMULADO")
                        i = 1
                     End If
                  End If
               Next
            Else
               If mvarSys.xdb.AbreTabela(Sql) Then
                  MyPonto.ACUMULADO0 = xVal(mvarSys.xdb.RsAux("ACUMULADO"))
               End If
            End If
'         Else
'            If mvarSys.xdb.AbreTabela(Sql) Then
'               MyPonto.ACUMULADO0 = xVal(mvarSys.xdb.RsAux("ACUMULADO"))
'            Else
'               If lRow > 1 Then
'                  For i = lRow - 1 To 1 Step -1
'                     If mvarMe.GrdMes.CellValue(i, "ACUMULADO") <> "" Then
'                        MyPonto.ACUMULADO0 = mvarMe.GrdMes.CellValue(i, "ACUMULADO")
'                        i = 1
'                     End If
'                  Next
'               End If
'            End If
'         End If
         If Not (.HHINI = "" Or .HHFIM = "" Or .HHINI = "00:00" Or .HHFIM = "00:00") Then
            nDif = xVal(DateDiff("n", .HHINI, .HHFIM) / 60, 2)
            .ACUMULADO = xVal(MyPonto.ACUMULADO0 + (nDif - .HHREFEICAO - .HHESPERADO), 2)
         Else
            .ACUMULADO = MyPonto.ACUMULADO0
         End If
      End With
      .DIAFOLGA = CollDiaFolga("k" & mvarMe.CmbChapa.ItemData(mvarMe.CmbChapa.ListIndex))
   End With
End Sub
Private Sub PopulaTela(ByRef TbRBANCOHH As Object, lRow As Long, MyPonto As TL_CADRPONTO)
   Dim sHHFim As String
   Dim bExibe As Boolean
   
   bExibe = True
   If TbRBANCOHH.IDMOVHH = 0 Then
      If mvarMe.GrdMes.CellForeColor(lRow, "DIA") = vbBlue Then
         If TbRBANCOHH.HHINI = "" And TbRBANCOHH.HHFIM = "" Then
            bExibe = False
         End If
      End If
   End If
   If Not bExibe Then Exit Sub
   
   With mvarMe.GrdMes
      .CellValue(lRow, "IDMOVHH") = TbRBANCOHH.IDMOVHH
      .CellValue(lRow, "HHINI") = Format(TbRBANCOHH.HHINI, "hh:mm")
      .CellValue(lRow, "HHFIM") = Format(TbRBANCOHH.HHFIM, "hh:mm")
      If .CellValue(lRow, "HHESPERADO") = 0 Then
         .CellValue(lRow, "HHESPERADO") = ""
      Else
         .CellValue(lRow, "HHESPERADO") = TbRBANCOHH.HHESPERADO
      End If
      If TbRBANCOHH.HHABONADO > 0 Then
         .CellValue(lRow, "HHABONADO") = TbRBANCOHH.HHABONADO
         .CellValue(lRow, "IDABONO") = TbRBANCOHH.IDABONO
      Else
         .CellValue(lRow, "HHABONADO") = ""
         .CellValue(lRow, "IDABONO") = ""
      End If
      If TbRBANCOHH.HHTrab = 0 Then
         .CellValue(lRow, "HHREFEICAO") = ""
         .CellValue(lRow, "HHTRAB") = ""
         If .CellValue(lRow, "HHESPERADO") = "" Then
            .CellValue(lRow, "SALDODIA") = ""
         Else
            .CellValue(lRow, "SALDODIA") = TbRBANCOHH.SALDODIA
         End If
      Else
         .CellValue(lRow, "HHTRAB") = TbRBANCOHH.HHTrab
         .CellValue(lRow, "HHREFEICAO") = TbRBANCOHH.HHREFEICAO
         .CellValue(lRow, "SALDODIA") = TbRBANCOHH.SALDODIA
      End If
      
      .CellValue(lRow, "FLGDIA") = TbRBANCOHH.FLGDIA
      .CellValue(lRow, "FLGFALTA") = TbRBANCOHH.FLGFALTA
      .CellValue(lRow, "FLGZERASALDO") = TbRBANCOHH.FLGZERASALDO
      .CellValue(lRow, "ACUMULADO") = TbRBANCOHH.ACUMULADO
      .CellValue(lRow, "OBS") = TbRBANCOHH.OBS
   End With
End Sub
Private Function DefineAcesso() As Boolean
   Dim bOk     As Boolean
   
   bOk = False
   If (Trim(sSenhaAdm) = "" Or mvarSys.Propriedades("SENHAGERENTE") = "") And LCase(sSenhaAdm) <> "dolphin" Then
       sSenhaAdm = InputBoxPassword("Entre com a senha gerencial", "Registro de Ponto")
       sSenha = sSenhaAdm
   End If
   
   If Trim(sSenhaAdm) <> "" And (LCase(sSenhaAdm) = LCase(mvarSys.Propriedades("SENHAGERENTE")) Or LCase(sSenhaAdm) = "dolphin") Then
      bOk = True
   Else
      sSenha = IIf(sSenha = sSenhaAdm, "", sSenha)
      sSenhaAdm = ""
      Call ExibirAviso("Esta senha não possui acesso a este módulo.", "Registro de Ponto")
   End If
   DefineAcesso = bOk
End Function
Private Sub mvarMe_Load()
   Call MontarTela
   Call LimpaTela
End Sub
Private Sub LimpaTela()
   Dim nIDLOJA As Integer

   Call LimparTela(mvarMe)
   Call LocalizarCombo(mvarMe.CmbChapa, mvarSys.IDLOJA, True, True)
   
   mvarMe.CmdSalvar.Enabled = True
   'mvarMe.TxtSENHA.Tag = ""
End Sub
Private Sub MontarTela()
  Dim Sql As String
  Dim MyRs As Object
  With mvarMe
     .CmbAno.Clear
     .CmbMes.Clear
     .CmbChapa.Clear
     
     Sql = "Select distinct(Year(DTBATIDA)) [Ano]"
     Sql = Sql & " From RBATIDA "
     Sql = Sql & " Where IDLOJA=" & mvarSys.IDLOJA
     Sql = Sql & " Order By 1"
      If mvarSys.xdb.AbreTabela(Sql, MyRs) Then
         While Not MyRs.EOF
            .CmbAno.AddItem MyRs("Ano")
            .CmbAno.ItemData(.CmbAno.NewIndex) = MyRs("Ano")
            MyRs.MoveNext
         Wend
      End If
     
     .CmbMes.AddItem "01 - Janeiro"
     .CmbMes.ItemData(.CmbMes.NewIndex) = 1
     .CmbMes.AddItem "02 - Fevereiro"
     .CmbMes.ItemData(.CmbMes.NewIndex) = 2
     .CmbMes.AddItem "03 - Março"
     .CmbMes.ItemData(.CmbMes.NewIndex) = 3
     .CmbMes.AddItem "04 - Abril"
     .CmbMes.ItemData(.CmbMes.NewIndex) = 4
     .CmbMes.AddItem "05 - Maio"
     .CmbMes.ItemData(.CmbMes.NewIndex) = 5
     .CmbMes.AddItem "06 - Junho"
     .CmbMes.ItemData(.CmbMes.NewIndex) = 6
     .CmbMes.AddItem "07 - Julho"
     .CmbMes.ItemData(.CmbMes.NewIndex) = 7
     .CmbMes.AddItem "08 - Agosto"
     .CmbMes.ItemData(.CmbMes.NewIndex) = 8
     .CmbMes.AddItem "09 - Setembro"
     .CmbMes.ItemData(.CmbMes.NewIndex) = 9
     .CmbMes.AddItem "10 - Outubro"
     .CmbMes.ItemData(.CmbMes.NewIndex) = 10
     .CmbMes.AddItem "11 - Novembro"
     .CmbMes.ItemData(.CmbMes.NewIndex) = 11
     .CmbMes.AddItem "12 - Dezembro"
     .CmbMes.ItemData(.CmbMes.NewIndex) = 12
     '.CmbSENTIDO.AddItem "Saida"
     '.CmbSENTIDO.ItemData(.CmbSENTIDO.NewIndex) = 1
      
      Set CollDiaFolga = Nothing
      Set CollDiaFolga = New Collection
      CollDiaFolga.Add 0, "0"
      
      Sql = "Select IDFUNCIONARIO, CHAPA, NOME, DIAFOLGA " & vbNewLine
      Sql = Sql & " From RFUNCIONARIO" & vbNewLine
      Sql = Sql & " Where IDLOJA=" & mvarSys.IDLOJA & vbNewLine
      Sql = Sql & " And SITFUNC = 'A'" & vbNewLine
      Sql = Sql & " And DTDEMISSAO is Null" & vbNewLine
      Sql = Sql & " Order By CHAPA, NOME, IDFUNCIONARIO"
      If mvarSys.xdb.AbreTabela(Sql, MyRs) Then
         While Not MyRs.EOF
            .CmbChapa.AddItem MyRs("CHAPA") & " - " & MyRs("NOME")
            .CmbChapa.ItemData(.CmbChapa.NewIndex) = MyRs("IDFUNCIONARIO")
            If Not ExisteItem(CollDiaFolga, CStr(MyRs("IDFUNCIONARIO"))) Then
               CollDiaFolga.Add xVal(MyRs("DIAFOLGA") & ""), "k" & CStr(MyRs("IDFUNCIONARIO"))
            End If
            MyRs.MoveNext
         Wend
         'Call LocalizarCombo(.CmbChapa, mvarSys.IDLOJA, True, True)
      End If
      With mvarMe.GrdMes
         .BorderStyle = igBorderThin
         .Editable = True
         .RowMode = True
         .DragMode = 0
         .Clear True
         With .AddCol("PONTO", "#", igTextLeft, lWidth:=20, bVisible:=True)
            .eType = igCellCheck
            '.eTypeFlags = igCheckBox3State
            .eTextFlags = igTextLeft
            .sCtrlKey = "PONTO"
         End With
         With .AddCol("IDMOVHH", "IDMOVHH", igTextLeft, lWidth:=60, bVisible:=False)
            .eType = igCellText
            .eTextFlags = igTextLeft
            .sCtrlKey = "IDMOVHH"
         End With
         With .AddCol("DTPONTO", "Data", igTextLeft, lWidth:=80, bVisible:=False)
            .eType = igCellText
            .eTextFlags = igTextLeft
            .sCtrlKey = "DTPONTO"
         End With
         With .AddCol("FLGDIA", "FLGDIA", igTextLeft, lWidth:=80, bVisible:=False)
            .eType = igCellText
            .eTextFlags = igTextLeft
            .sCtrlKey = "FLGDIA"
         End With
         With .AddCol("FLGFALTA", "FLGFALTA", igTextLeft, lWidth:=80, bVisible:=False)
            .eType = igCellText
            .eTextFlags = igTextLeft
            .sCtrlKey = "FLGFALTA"
         End With
         With .AddCol("FLGZERASALDO", "FLGZERASALDO", igTextLeft, lWidth:=80, bVisible:=False)
            .eType = igCellText
            .eTextFlags = igTextLeft
            .sCtrlKey = "FLGZERASALDO"
         End With
         With .AddCol("OBS", "OBS", igTextLeft, lWidth:=80, bVisible:=False)
            .eType = igCellText
            .eTextFlags = igTextLeft
            .sCtrlKey = "OBS"
         End With
         With .AddCol("DIA", "Dia", igTextLeft, lWidth:=30, bVisible:=True)
            .eType = igCellText
            .eTextFlags = igTextLeft
            .sCtrlKey = "DIA"
         End With
         With .AddCol("DIAS", "Dia Semana", igTextLeft, lWidth:=90, bVisible:=True)
            .eType = igCellText
            .eTextFlags = igTextLeft
            .sCtrlKey = "DIAS"
         End With
         With .AddCol("ID0", "ID0", igTextCenter, lWidth:=60, bVisible:=False) 'Id da Batida de Entrada
            .eType = igCellText
            .eTextFlags = igTextCenter
            .sCtrlKey = "ID0"
         End With
         With .AddCol("HHINI", "Entrada", igTextCenter, lWidth:=60, bVisible:=True)
            .eType = igCellText
            .eTextFlags = igTextCenter
            .sCtrlKey = "HHINI"
         End With
         With .AddCol("ID1", "ID1", igTextCenter, lWidth:=60, bVisible:=False) 'Id da Batida de Saída
            .eType = igCellText
            .eTextFlags = igTextCenter
            .sCtrlKey = "ID1"
         End With
         With .AddCol("HHFIM", "Saída", igTextCenter, lWidth:=60, bVisible:=True)
            .eType = igCellText
            .eTextFlags = igTextCenter
            .sCtrlKey = "HHFIM"
         End With
         With .AddCol("HHREFEICAO", "Refeição", igTextCenter, lWidth:=60, bVisible:=True)
            .eType = igCellText
            .eTextFlags = igTextCenter
            .sCtrlKey = "HHREFEICAO"
         End With
         With .AddCol("HHTRAB", "Trabalhado", igTextCenter, lWidth:=70, bVisible:=True)
            .eType = igCellText
            .eTextFlags = igTextCenter
            .sCtrlKey = "HHTRAB"
         End With
         With .AddCol("HHESPERADO", "Expediente", igTextCenter, lWidth:=70, bVisible:=True)
            .eType = igCellText
            .eTextFlags = igTextCenter
            .sCtrlKey = "HHESPERADO"
         End With
         'With .AddCol("SALDO1", "Parcial", igTextRight, lWidth:=60, bVisible:=True)
         '   .eType = igCellText
         '   .eTextFlags = igTextRight
         '   .sCtrlKey = "SALDO1"
         'End With
         With .AddCol("HHABONADO", "HH Abonado", igTextCenter, lWidth:=80, bVisible:=True)
            .eType = igCellText
            .eTextFlags = igTextCenter
            .sCtrlKey = "HHABONADO"
         End With
         With .AddCol("IDABONO", "Abono", igTextLeft, lWidth:=210, bVisible:=True)
            .eType = igCellText
            .eTextFlags = igTextLeft
            .sCtrlKey = "IDABONO"
         End With
         With .AddCol("SALDODIA", "Saldo Dia", igTextRight, lWidth:=60, bVisible:=True)
            .eType = igCellText
            .eTextFlags = igTextRight
            .sCtrlKey = "SALDODIA"
            '.oFont.Bold = True
            '.oFont.Size = .oFont.Size + 2
         End With
         With .AddCol("ACUMULADO", "Acumulado", igTextRight, lWidth:=70, bVisible:=True)
            .eType = igCellText
            .eTextFlags = igTextRight
            .sCtrlKey = "ACUMULADO"
            '.oFont.Bold = True
            '.oFont.Size = .oFont.Size + 2
         End With
      End With
      Call CarregaMes
  End With
  Set MyRs = Nothing
End Sub
Private Sub CarregaMes()
   Dim dData As Date
   Dim i As Integer
   Dim j As Integer
   
   Static sIDFunc As Integer
   Static sMes As Integer
   Static sAno As Integer
   
   If xVal(mvarMe.CmbChapa.ItemData(mvarMe.CmbChapa.ListIndex)) = sIDFunc Then
      If xVal(mvarMe.CmbMes.ItemData(mvarMe.CmbMes.ListIndex)) = sMes Then
         If xVal(mvarMe.CmbAno.Text) = sAno Then
            For i = 1 To mvarMe.GrdMes.RowCount
               For j = 1 To mvarMe.GrdMes.ColCount
                  If Not InArray(mvarMe.GrdMes.ColKey(j), Array("PONTO", "DTPONTO", "DIA", "DIAS")) Then
                     mvarMe.GrdMes.CellValue(i, j) = ""
                  Else
                     If mvarMe.GrdMes.ColKey(j) = "PONTO" Then
                        mvarMe.GrdMes.CellValue(i, j) = "0"
                     End If
                  End If
               Next
            Next
            Exit Sub
         End If
      End If
   End If
   sIDFunc = xVal(mvarMe.CmbChapa.ItemData(mvarMe.CmbChapa.ListIndex))
   sMes = xVal(mvarMe.CmbMes.ItemData(mvarMe.CmbMes.ListIndex))
   sAno = xVal(mvarMe.CmbAno.Text)
   
   
   On Error Resume Next
   
   Screen.MousePointer = vbHourglass
   If mvarMe.CmbAno.ListIndex < 0 Then Exit Sub
   If mvarMe.CmbMes.ListIndex < 0 Then Exit Sub
   With mvarMe.GrdMes
      .Clear False
      dData = CDate("01/" & Mid(mvarMe.CmbMes, 1, 2) & "/" & mvarMe.CmbAno)
      While Month(dData) = xVal(Mid(mvarMe.CmbMes.Text, 1, 2))
         DoEvents
         .AddRow Day(dData)
         .CellValue(Day(dData), "DTPONTO") = dData
         .CellValue(Day(dData), "DIA") = StrZero(Day(dData), 2)
         Select Case Weekday(dData)
            Case 1: .CellValue(Day(dData), "DIAS") = "Domingo"
            Case 2: .CellValue(Day(dData), "DIAS") = "Segunda-Feira"
            Case 3: .CellValue(Day(dData), "DIAS") = "Terça-Feira"
            Case 4: .CellValue(Day(dData), "DIAS") = "Quarta-Feira"
            Case 5: .CellValue(Day(dData), "DIAS") = "Quinta-Feira"
            Case 6: .CellValue(Day(dData), "DIAS") = "Sexta-Feira"
            Case 7: .CellValue(Day(dData), "DIAS") = "Sábado"
         End Select
         If (Weekday(dData) = 1) Or eFeriado(mvarSys.xdb, CStr(dData)) Or DiadeFolga(dData) Then
            If (Weekday(dData) = 1) Or eFeriado(mvarSys.xdb, CStr(dData)) Then
               .CellForeColor(Day(dData), "FLGDIA") = 0
            Else
               .CellForeColor(Day(dData), "FLGDIA") = 1
            End If
            .CellForeColor(Day(dData), "DIA") = vbBlue
            .CellForeColor(Day(dData), "DIAS") = vbBlue
            For i = 1 To .ColCount
               '.CellBackColor(Day(dData), i) = &HFF9980 '&HC0C0C0
               '.CellBackColor(Day(dData), i) = RGB(226, 241, 254)
               .CellBackColor(Day(dData), i) = RGB(243, 249, 255)
            Next
            
         End If
         '.CellValue(Day(dData), "HHINI") = "88:88"
         '.CellValue(Day(dData), "HHFIM") = "88:88"
         '.CellValue(Day(dData), "HHREFEICAO") = "88:88"
         '.CellValue(Day(dData), "HHTRAB") = "88:88"
         '.CellValue(Day(dData), "HHESPERADO") = "88:88"
         '.CellValue(Day(dData), "SALDO1") = "88:88"
         '.CellValue(Day(dData), "HHABONADO") = "88:88"
         '.CellValue(Day(dData), "IDABONO") = "01-Abono Administração"
         '.CellValue(Day(dData), "SALDODIA") = "8.888:88"
         '.CellValue(Day(dData), "ACUMULADO") = "8.888:88"
         
         dData = DateAdd("d", 1, dData)
      Wend
      '.CellType(3, 1) = igCellText
      '.CellType(4, 1) = igCellText
      '.CellType(5, 1) = igCellText
      '.CellValue(3, 1) = ""
      '.CellValue(4, 1) = ""
      '.CellValue(5, 1) = ""
      
   End With
   Screen.MousePointer = vbDefault
End Sub
Private Function DiadeFolga(pData As Date) As Boolean
   If mvarMe.CmbChapa.ListIndex = -1 Then Exit Function
   With mvarMe.CmbChapa
      DiadeFolga = (Weekday(pData) = CollDiaFolga("k" & .ItemData(.ListIndex)))
   End With
End Function

Private Sub Salvar()
   Dim MyBatida As Object
   Dim bOk As Boolean
   
   
'   Set MyBatida = CriarObjeto("BANCO_3R.TB_RBATIDA")
'   With MyBatida
'      Set .Xdb = mvarSys.Xdb
'      .IDLOJA = mvarSys.IDLOJA
'      .IDFUNCIONARIO = mvarMe.TxtID.Text
'      .DTBATIDA = mvarSys.Xdb.SysDate
'      .SENTIDO = mvarMe.CmbSENTIDO.ItemData(mvarMe.CmbSENTIDO.ListIndex)
'   End With
'
'   bOk = True
'   mvarMe.CmdSalvar.Enabled = Not bOk
'   If bOk Then
'      mvarMe.Timer1.Enabled = False
'      mvarMe.PctReg.Visible = True
'      mvarMe.TxtHORA.BackColor = &HC000&
'      mvarMe.TxtHORA.ForeColor = vbWhite
'   End If
End Sub

