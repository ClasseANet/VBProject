VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NG_Tarefa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'1-Boas Vindas
'2-Confirmar Agenda
'3-eMail de Orientação
'4-Verificar Resultado de Avaliação
'5-Cancelamento/Falta de Sessão
'6-Sessão Não Marcada
Dim gDebug As Boolean
Private mvarsys      As Object
Private mvarIDLOJA   As String
Private mvarBarItem  As Object
Private mvarTLPane   As Object
Private mvarPane     As Object
Private mvarTLCal    As Object
Public Property Set TLCal(ByVal vData As Object)
   Set mvarTLCal = vData
   Set mvarTLCal.Sys = mvarsys
End Property
Public Property Get TLCal() As Object
   Set TLCal = mvarTLCal
End Property
Public Property Set Sys(ByVal vData As Object)
   Set mvarsys = vData
   
   If Val(mvarsys.Propriedades("IDCONTA")) = 0 Then
      Dim Tb_Conta As Object
      If DefineContaLAN(Tb_Conta, 1, mvarsys) Then
         mvarsys.Propriedades("IDCONTA") = Tb_Conta.IDCONTA
      End If
      Set Tb_Conta = Nothing
   End If
   
   Set mvarTLPane.Sys = mvarsys
   mvarIDLOJA = mvarsys.Propriedades("IDLOJA")
   
End Property
Public Property Get Sys() As Object
   Set Sys = mvarsys
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
   Set mvarTLPane.TLCal = mvarTLCal
   mvarTLPane.Show Reload
End Sub
Private Sub Class_Initialize()
   gDebug = False
   Set mvarTLPane = New TL_PaneTarefa
   Set mvarPane = mvarTLPane.Pane
End Sub
Public Sub SalvarTarefas(ByRef pTbEvento As Object, Optional ByRef pQueries As Collection, Optional pTeste As Boolean = False, Optional pTipoTarefa As Integer)
   Dim TbServico As Object
   Dim MyRs As Object
   Dim TbSessao  As Object
   Dim bTeste    As Boolean
   
   If pTbEvento.isMeeting = 0 Then Exit Sub
   If pTbEvento.IDEVENTO = 0 Then Exit Sub
   
   Set TbServico = CreateObject("Banco_3R.TB_OSERVICOEVT")
   Set TbServico.xDb = mvarsys.xDb
   If TbServico.Pesquisar(Ch_IDLOJA:=pTbEvento.IDLOJA, Ch_IDEVENTO:=pTbEvento.IDEVENTO, Ch_IDSERVICOEVT:=1) Then
      bTeste = (TbServico.IDTPSERVICO = 1)
   Else
      bTeste = pTeste
   End If
   
   If CDate(Format(pTbEvento.STARTDATETIME, "dd/mm/yyyy")) >= CDate(mvarsys.xDb.SysDate()) Then
      If pTipoTarefa = 0 Then
         Call QrySaveTarefa(eTpTarefa.TarBoasVindas, pTbEvento, pQueries)  'Boas Vindas
         If bTeste Then
            Call QrySaveTarefa(eTpTarefa.TarVerifAvali, pTbEvento, pQueries)  'Verificar Resultado de Avaliação
         Else
            Call QrySaveTarefa(eTpTarefa.TarConfimaAge, pTbEvento, pQueries) 'Confirmar Agenda
            Call QrySaveTarefa(eTpTarefa.TarOrientacao, pTbEvento, pQueries) 'eMail de Orientação
            Call QrySaveTarefa(eTpTarefa.TarECancelado, pTbEvento, pQueries) 'Cancelamento/Falta de Sessão
            
            Call CancelarTarefa(pTbEvento, pQueries)
         End If
      Else
         Call QrySaveTarefa(pTipoTarefa, pTbEvento, pQueries)
      End If
   Else
      If pTipoTarefa = 0 Then
         Call QrySaveTarefa(eTpTarefa.TarNaoMarcado, pTbEvento, pQueries) 'Sessão Não Marcada
         Call QrySaveTarefa(eTpTarefa.TarECancelado, pTbEvento, pQueries) 'Cancelamento/Falta de Sessão
      Else
         Call QrySaveTarefa(pTipoTarefa, pTbEvento, pQueries)
      End If
   End If
End Sub
Public Function CancelarTarefa(ByRef pTbEvento As Object, Optional pQueries As Collection) As String
   Dim Sql As String
   Dim sIn As String
   Dim MyRs As Object
   
   Sql = "Select T.IDLOJA, T.IDTAREFA, T.IDEVENTO" & vbNewLine
   Sql = Sql & " From OTAREFAEVT T" & vbNewLine
   Sql = Sql & " Where T.IDLOJA=" & pTbEvento.IDLOJA & vbNewLine
   Sql = Sql & " And T.IDCLIENTE=" & pTbEvento.IDCLIENTE & vbNewLine
   Sql = Sql & " And T.IDTPTAREFA=" & eTpTarefa.TarNaoMarcado & vbNewLine
   Sql = Sql & " And T.SITTAREFA<>10" & vbNewLine
   Sql = Sql & " And T.FLGDELETE=0"
   sIn = ""
   If mvarsys.xDb.AbreTabela(Sql, MyRs) Then
      While Not MyRs.EOF
         sIn = sIn & IIf(sIn = "", "", ", ") & MyRs("IDTAREFA") & ""
         MyRs.MoveNext
      Wend
      If Trim(sIn) <> "" Then
         Sql = "Update OTAREFAEVT" & vbNewLine
         Sql = Sql & " Set FLGDELETE=1" & vbNewLine
         Sql = Sql & ", ALTERSTAMP=1" & vbNewLine
         Sql = Sql & ", TIMESTAMP=GetDate()" & vbNewLine
         Sql = Sql & " Where IDLOJA=" & pTbEvento.IDLOJA & vbNewLine
         Sql = Sql & " And IDTAREFA in (" & sIn & ")"
         
         pQueries.Add Sql
      End If
   End If
   
End Function
Public Function QrySaveTarefa(pTPTarefa As Integer, ByRef pTbEvento As Object, Optional pQueries As Collection) As String
   Dim TbTarefa   As Object
   Dim TBAtend    As Object
   Dim bExiste    As Boolean
   Dim Sql        As String
   Dim sDTTarefa  As String
   Dim sDscIni    As String
   Dim sMail      As String
   Dim nDifHora   As Long
   Dim nFLGDELETE As Integer
   
   If pTbEvento.IDEVENTO = 0 Then Exit Function
   
   If pTPTarefa = eTpTarefa.TarBoasVindas Or pTPTarefa = eTpTarefa.TarOrientacao Then
      Sql = "Select EMAIL "
      Sql = Sql & " From OCLIENTE "
      Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
      Sql = Sql & " And IDCLIENTE = " & pTbEvento.IDCLIENTE
      If mvarsys.xDb.AbreTabela(Sql) Then
         sMail = Trim(mvarsys.xDb.RsAux("EMAIL") & "")
      End If
      If sMail = "" Then Exit Function
   End If
   
   Set TbTarefa = CreateObject("Banco_3R.TB_OTAREFAEVT")
   Set TbTarefa.xDb = mvarsys.xDb
   With pTbEvento
      
      sDTTarefa = DefineDTTarefa(pTPTarefa, .STARTDATETIME)
      
      If Trim(sDTTarefa) = "" Then Exit Function
      
      If CDate(Format(sDTTarefa, "dd/mm/yyyy")) < CDate(mvarsys.xDb.SysDate()) Then
         If pTbEvento.STARTDATETIME >= CDate(mvarsys.xDb.SysDate()) Then
            Sql = " IDEVENTO =" & .IDEVENTO
            Sql = Sql & " And IDTPTAREFA =" & pTPTarefa
            Sql = Sql & " And FLGDELETE =0"
            If TbTarefa.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_Where:=Sql) Then
               TbTarefa.FLGDELETE = 1
               pQueries.Add TbTarefa.QrySave
               QrySaveTarefa = TbTarefa.QrySave
            End If
         End If
         Exit Function
      End If
      
      Sql = " IDEVENTO=" & .IDEVENTO
      Sql = Sql & " And IDTPTAREFA=" & pTPTarefa
      If TbTarefa.Pesquisar(Ch_IDLOJA:=.IDLOJA, Ch_Where:=Sql) Then
         bExiste = True
         If Day(TbTarefa.DTTAREFA) <> Day(sDTTarefa) Then
            If pTbEvento.FLGCONFIRMADO = 0 And TbTarefa.SITTAREFA = 10 And pTPTarefa = eTpTarefa.TarConfimaAge Then
               If vbYes = ExibirPergunta("Sessão não confirmada. Deseja reabrir tarefa?", "Tarefas") Then
                  TbTarefa.SITTAREFA = 0
               End If
            End If
         End If
         
         sDscIni = DefineDscTarefa(pTPTarefa) & vbNewLine
         sDscIni = sDscIni & "Evento: "
         sDscIni = sDscIni & Format(.STARTDATETIME, "dd/mm/yy hh:mm") & "h" & vbNewLine
         If Len(TbTarefa.DSCTAREFA) = Len(sDscIni) Then
            TbTarefa.DSCTAREFA = sDscIni
         End If
         
         TbTarefa.DTTAREFA = sDTTarefa
         TbTarefa.DTLEMBRETE = sDTTarefa
         If pTPTarefa = 4 Then 'Verificar Cancelameto/Falta
            nFLGDELETE = IIf(.FLGCANCELADO = 1, 0, 1)
         Else
            nFLGDELETE = .FLGCANCELADO
         End If
         
      Else
         bExiste = False
         sDscIni = DefineDscTarefa(pTPTarefa) & vbNewLine
         sDscIni = sDscIni & "Evento: "
         sDscIni = sDscIni & Format(.STARTDATETIME, "dd/mm/yy hh:mm") & "h" & vbNewLine
      
         TbTarefa.DSCTAREFA = sDscIni
         TbTarefa.SITTAREFA = 0
         TbTarefa.Prioridade = 0
         TbTarefa.FLGLEMBRETE = 1
         TbTarefa.DTLEMBRETE = sDTTarefa
         
         nFLGDELETE = .FLGCANCELADO
      End If
      
      If pTPTarefa = eTpTarefa.TarBoasVindas Then  'Verificar Avaliação
         Set TBAtend = CreateObject("Banco_3R.TB_OATENDIMENTO")
         Set TBAtend.xDb = mvarsys.xDb
         If TBAtend.Pesquisar(Ch_IDLOJA:=pTbEvento.IDLOJA, Ch_Where:="IDEVENTO=" & pTbEvento.IDEVENTO) Then
            nFLGDELETE = .FLGCANCELADO
            Sql = "IDCLIENTE=" & pTbEvento.IDCLIENTE
            Sql = Sql & " And IDEVENTO<>" & pTbEvento.IDEVENTO
            If TBAtend.Pesquisar(Ch_IDLOJA:=pTbEvento.IDLOJA, Ch_Where:=Sql) Then
               nFLGDELETE = 1
            End If
         Else
            nFLGDELETE = 1
         End If
      End If
      
      If pTPTarefa = eTpTarefa.TarVerifAvali Then  'Verificar Avaliação
         Set TBAtend = CreateObject("Banco_3R.TB_OATENDIMENTO")
         Set TBAtend.xDb = mvarsys.xDb
         If TBAtend.Pesquisar(Ch_IDLOJA:=pTbEvento.IDLOJA, Ch_Where:="IDEVENTO=" & pTbEvento.IDEVENTO) Then
            nFLGDELETE = .FLGCANCELADO
         Else
            nFLGDELETE = 1
         End If
      End If
      
      If pTPTarefa = eTpTarefa.TarECancelado Then  'Verificar Cancelamento/Falta de Sessão
         nFLGDELETE = IIf(.FLGCANCELADO = 1, 0, 1)
      End If
      
      If pTPTarefa = eTpTarefa.TarNaoMarcado Then   'Verificar Remarcação
         Sql = "Select * " & vbNewLine
         Sql = Sql & " From OTIPOTAREFA" & vbNewLine
         Sql = Sql & " Where IDLOJA=" & mvarIDLOJA & vbNewLine
         Sql = Sql & " And IDTPTAREFA=" & pTPTarefa
         If mvarsys.xDb.AbreTabela(Sql) Then
            nDifHora = xVal(mvarsys.xDb.RsAux("DIFHORA") & "")
            If DateDiff("d", CDate(.STARTDATETIME), CDate(mvarsys.xDb.SysDate())) < (nDifHora / 24) Then
               nFLGDELETE = 1
            End If
         End If
      End If
      
      If nFLGDELETE = 1 And Not bExiste Then Exit Function
      
      TbTarefa.FLGDELETE = nFLGDELETE
      TbTarefa.IDTPTAREFA = pTPTarefa
      TbTarefa.IDLOJA = .IDLOJA
      TbTarefa.IDEVENTO = .IDEVENTO
      TbTarefa.IDCLIENTE = .IDCLIENTE
      TbTarefa.TITULO = .Subject
      TbTarefa.DTTAREFA = sDTTarefa
      TbTarefa.FLGAUTO = 1
      
      If TbTarefa.IsDirt Then
        pQueries.Add TbTarefa.QrySave
      End If
      QrySaveTarefa = TbTarefa.QrySave
   End With
   Set TbTarefa = Nothing
End Function
Private Function DefineDscTarefa(pTipo As Integer) As String
   Dim sAux As String
   Dim TbTipo As Object
   
   Set TbTipo = CreateObject("Banco_3R.TB_OTIPOTAREFA")
   Set TbTipo.xDb = mvarsys.xDb
   With TbTipo
      If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, CH_IDTPTAREFA:=pTipo) Then
         sAux = .DSCTPTAREFA
      End If
   End With
   DefineDscTarefa = sAux
   Set TbTipo = Nothing
End Function
'1-Confirmar Agenda
'2-eMail de Orientação
'3-Verificar Resultado de Avaliação
'4-Cancelamento/Falta de Sessão
'5-Sessão Não Marcada
Private Function DefineDTTarefa(pTipo As Integer, pDataEvento As String) As String
   Dim sDTTarefa As String
   Dim TbTipo As Object
   
   Set TbTipo = CreateObject("Banco_3R.TB_OTIPOTAREFA")
   Set TbTipo.xDb = mvarsys.xDb
   With TbTipo
      If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, CH_IDTPTAREFA:=pTipo) Then
         sDTTarefa = Format(CDate(pDataEvento) + (.DIFHORA / 24), "DD/MM/YYYY hh:mm")
         
         If pTipo = eTpTarefa.TarNaoMarcado Then
            If CDate(sDTTarefa) <= mvarsys.xDb.SysDate() Then
               sDTTarefa = DateAdd("d", 1, CDate(mvarsys.xDb.SysDate()))
            End If
         End If

         If .DIFHORA > 0 Then
            If Weekday(sDTTarefa) = 1 Then
               sDTTarefa = Format(CDate(sDTTarefa) + 1, "DD/MM/YYYY hh:mm")
            ElseIf Weekday(sDTTarefa) = 7 Then
               sDTTarefa = Format(CDate(sDTTarefa) + 2, "DD/MM/YYYY hh:mm")
            End If
            While eFeriado(mvarsys.xDb, sDTTarefa)
               sDTTarefa = Format(CDate(sDTTarefa) + 1, "DD/MM/YYYY hh:mm")
            Wend
         Else
            If Weekday(sDTTarefa) = 1 Then
               sDTTarefa = Format(CDate(sDTTarefa) - 2, "DD/MM/YYYY hh:mm")
            ElseIf Weekday(sDTTarefa) = 7 Then
               sDTTarefa = Format(CDate(sDTTarefa) - 1, "DD/MM/YYYY hh:mm")
            End If
            While eFeriado(mvarsys.xDb, sDTTarefa)
               sDTTarefa = Format(CDate(sDTTarefa) - 1, "DD/MM/YYYY hh:mm")
            Wend
          End If
      End If
    End With
    Set TbTipo = Nothing
    DefineDTTarefa = Format(sDTTarefa, "DD/MM/YYYY hh:mm:ss")
End Function
Public Sub GravaPrevisaoMeta()
   Dim Sql     As String
   Dim TbItem  As Object
   Dim dHoje   As Date
   Dim nVLPREV As Double
   
   dHoje = mvarsys.xDb.SysDate(1)
     
   Sql = "Select Min(DTITEM) As DTITEM "
   Sql = Sql & " From PMETAITEM"
   Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
   Sql = Sql & " And IDMETA =1"
   Sql = Sql & " And Month(DTITEM)=" & Month(dHoje)
   Sql = Sql & " And Year(DTITEM)=" & Year(dHoje)
      
   If gDebug Then MsgBox "GravaPrevisaoMeta 01"
   If mvarsys.xDb.AbreTabela(Sql) Then
      
      
      
      If IsDate(mvarsys.xDb.RsAux("DTITEM") & "") Then
         Set TbItem = CriarObjeto("Banco_3R.TB_PMETAITEM")
         With TbItem
            Set .xDb = mvarsys.xDb
            If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDMETA:=1, Ch_DTITEM:=mvarsys.xDb.RsAux("DTITEM") & "") Then
               mvarsys.Propriedades("VLPREV") = .VLPREV
            End If
         End With
      Else
         Sql = " Select Count(*) [VLPREV]"
         Sql = Sql & " From OEVENTOAGENDA E"
         Sql = Sql & " join OSERVICOEVT S On E.IDLOJA=S.IDLOJA and E.IDEVENTO=S.IDEVENTO"
         Sql = Sql & " Join OCLIENTE C On E.IDLOJA=C.IDLOJA AND E.IDCLIENTE=C.IDCLIENTE And C.ISENTO=0"
         Sql = Sql & " Where E.IDLOJA=" & mvarIDLOJA
         Sql = Sql & " And day(E.STARTDATETIME) >= " & Day(dHoje)
         Sql = Sql & " And Month(E.STARTDATETIME) =" & Month(dHoje)
         Sql = Sql & " And Year(E.STARTDATETIME) = " & Year(dHoje)
         Sql = Sql & " And E.FLGCANCELADO<>1"
         Sql = Sql & " And S.IDTPSERVICO <>1"
         Sql = Sql & " Group By Year(E.STARTDATETIME), Month(E.STARTDATETIME)"
         Sql = Sql & " Order by Year(E.STARTDATETIME), Month(E.STARTDATETIME)"
         If mvarsys.xDb.AbreTabela(Sql) Then
            nVLPREV = xVal(mvarsys.xDb.RsAux("VLPREV") & "")
         End If
      
         Dim TbMeta  As Object
         Set TbMeta = CriarObjeto("Banco_3R.TB_PMETA")
         Set TbMeta.xDb = mvarsys.xDb
         If Not TbMeta.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDMETA:=1) Then
            TbMeta.IDLOJA = mvarIDLOJA
            TbMeta.IDMETA = 1
            Call TbMeta.Salvar
         End If
         Set TbMeta = Nothing
      
         If TbItem Is Nothing Then
            Set TbItem = CriarObjeto("Banco_3R.TB_PMETAITEM")
         End If
         With TbItem
            Set .xDb = mvarsys.xDb
            .IDLOJA = mvarIDLOJA
            .IDMETA = 1
            .DTITEM = CStr(dHoje)
            .VLPREV = nVLPREV
            
            If gDebug Then MsgBox "Salvar Item"
            On Error Resume Next
            .Salvar
         End With
         mvarsys.Propriedades("VLPREV") = nVLPREV
      End If
   End If
End Sub
Public Sub ExibirTarefa(pCod As Long, Optional pIDCLIENTE As Integer)
   Dim TbTarefa As Object
   Dim TBAtend As Object
   Dim TlTarefa As TL_CadTarefa
   
   Set TbTarefa = CriarObjeto("BANCO_3R.TB_OTAREFAEVT")
   Set TbTarefa.xDb = mvarsys.xDb
   If TbTarefa.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDTAREFA:=pCod) Then
      pCod = TbTarefa.IDTAREFA
      pIDCLIENTE = TbTarefa.IDCLIENTE
   End If
   Set TlTarefa = New TL_CadTarefa
   With TlTarefa
      Set .Sys = mvarsys
      .IDLOJA = mvarIDLOJA
      .IDTAREFA = TbTarefa.IDTAREFA
      .IDCLIENTE = pIDCLIENTE
      
      .Show vbModal
   End With
   Set TlTarefa = Nothing
   Set TbTarefa = Nothing
End Sub

