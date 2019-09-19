VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NG_Consulta3R"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mvarSys      As Object
Private mvarIDLOJA   As String
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
   
   If Val(mvarSys.Propriedades("IDCONTA")) = 0 Then
      Dim Tb_Conta As Object
      If DefineContaLAN(Tb_Conta, 1, mvarSys) Then
         mvarSys.Propriedades("IDCONTA") = Tb_Conta.IDCONTA
      End If
      Set Tb_Conta = Nothing
   End If
   
   Set mvarTLPane.Sys = mvarSys
   mvarIDLOJA = mvarSys.Propriedades("IDLOJA")
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
   Set mvarTLPane.TLCal = mvarTLCal
   mvarTLPane.Show Reload
End Sub
Private Sub Class_Initialize()
   Set mvarTLPane = New TL_PaneDiario
   Set mvarPane = mvarTLPane.Pane
End Sub
Public Sub SalvarTarefas(ByRef pTbEvento As Object, Optional pQueries As Collection, Optional pTeste As Boolean = False, Optional pTipoTarefa As Integer)
   Dim TbServico As Object
   Dim TbSessao  As Object
   Dim bTeste    As Boolean
   
   If pTbEvento.isMeeting = 0 Then Exit Sub
   If pTbEvento.IDEVENTO = 0 Then Exit Sub
   
   Set TbServico = CreateObject("Banco_3R.TB_OSERVICOEVT")
   Set TbServico.xDb = mvarSys.xDb
   If TbServico.Pesquisar(Ch_IDLOJA:=pTbEvento.IDLOJA, Ch_IDEVENTO:=pTbEvento.IDEVENTO, Ch_IDSERVICOEVT:=1) Then
      bTeste = (TbServico.IDTPSERVICO = 1)
   Else
      bTeste = pTeste
   End If
   
   If CDate(Format(pTbEvento.STARTDATETIME, "dd/mm/yyyy")) >= CDate(mvarSys.xDb.SysDate()) Then
      If pTipoTarefa = 0 Then
         If bTeste Then
            Call QrySaveTarefa(3, pTbEvento, pQueries)
         Else
            Call QrySaveTarefa(1, pTbEvento, pQueries) 'Confirmar Agenda
            Call QrySaveTarefa(2, pTbEvento, pQueries) 'eMail de Orientação
            Call QrySaveTarefa(4, pTbEvento, pQueries) 'Cancelamento Falta de Sessão
         End If
      Else
         Call QrySaveTarefa(pTipoTarefa, pTbEvento, pQueries)
      End If
   Else
      If pTipoTarefa = 0 Then
         Call QrySaveTarefa(5, pTbEvento, pQueries) 'Cancelamento Falta de Sessão
      Else
         Call QrySaveTarefa(pTipoTarefa, pTbEvento, pQueries)
      End If
   End If
End Sub
Public Function QrySaveTarefa(pTPTarefa As Integer, ByRef pTbEvento As Object, Optional pQueries As Collection) As String
   Dim TbTarefa   As Object
   Dim TbAtend    As Object
   Dim bExiste    As Boolean
   Dim Sql        As String
   Dim sDTTarefa  As String
   Dim sDscIni    As String
   
   If pTbEvento.IDEVENTO = 0 Then Exit Function
   
   Set TbTarefa = CreateObject("Banco_3R.TB_OTAREFAEVT")
   Set TbTarefa.xDb = mvarSys.xDb
   With pTbEvento
      sDTTarefa = DefineDTTarefa(pTPTarefa, .STARTDATETIME)
      If CDate(Format(sDTTarefa, "dd/mm/yyyy")) < CDate(mvarSys.xDb.SysDate()) Then
         Sql = " IDLOJA =" & .IDLOJA
         Sql = Sql & " And IDEVENTO =" & .IDEVENTO
         Sql = Sql & " And IDTPTAREFA =" & pTPTarefa
         Sql = Sql & " And FLGDELETE =0"
         If TbTarefa.Pesquisar(Ch_Where:=Sql) Then
            TbTarefa.FLGDELETE = 1
            pQueries.Add TbTarefa.QrySave
            QrySaveTarefa = TbTarefa.QrySave
         End If
         
         Exit Function
      End If
      
      Sql = " IDEVENTO=" & .IDEVENTO
      Sql = Sql & " And IDTPTAREFA=" & pTPTarefa
      If TbTarefa.Pesquisar(Ch_IDLOJA:=.IDLOJA, Ch_Where:=Sql) Then
         bExiste = True
         If Day(TbTarefa.DTTAREFA) <> Day(sDTTarefa) Then
            If pTbEvento.FLGCONFIRMADO = 0 And TbTarefa.SITTAREFA = 10 And pTPTarefa = 1 Then
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
         If pTPTarefa = 4 Then
            TbTarefa.FLGDELETE = IIf(.FLGCANCELADO = 1, 0, 1)
         Else
            TbTarefa.FLGDELETE = .FLGCANCELADO
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
      End If
               
      If pTPTarefa = 3 Then 'Verificar Avaliação
         Set TbAtend = CreateObject("Banco_3R.TB_OATENDIMENTO")
         Set TbAtend.xDb = mvarSys.xDb
         If TbAtend.Pesquisar(Ch_IDLOJA:=pTbEvento.IDLOJA, Ch_Where:="IDEVENTO=" & pTbEvento.IDEVENTO) Then
            TbTarefa.FLGDELETE = .FLGCANCELADO
         Else
            TbTarefa.FLGDELETE = 1
         End If
            
      ElseIf pTPTarefa = 4 Then 'Verificar Cancelamento/Falta de Sessão
         TbTarefa.FLGDELETE = IIf(.FLGCANCELADO = 1, 0, 1)
      Else
         TbTarefa.FLGDELETE = .FLGCANCELADO
      End If
      If TbTarefa.FLGDELETE = 1 And Not bExiste Then Exit Function
      
      TbTarefa.IDTPTAREFA = pTPTarefa
      TbTarefa.IDLOJA = .IDLOJA
      TbTarefa.IDEVENTO = .IDEVENTO
      TbTarefa.IDCLIENTE = .IDCLIENTE
      TbTarefa.TITULO = .Subject
      TbTarefa.DTTAREFA = sDTTarefa
      TbTarefa.FLGAUTO = 1
      
      pQueries.Add TbTarefa.QrySave
      QrySaveTarefa = TbTarefa.QrySave
   End With
   Set TbTarefa = Nothing
End Function
Private Function DefineDscTarefa(pTipo As Integer) As String
   Dim sAux As String
   Dim TbTipo As Object
   
   Set TbTipo = CreateObject("Banco_3R.TB_OTIPOTAREFA")
   Set TbTipo.xDb = mvarSys.xDb
   With TbTipo
      If .Pesquisar(CH_IDTPTAREFA:=pTipo) Then
         sAux = .DSCTPTAREFA
      End If
   End With
   DefineDscTarefa = sAux
   Set TbTipo = Nothing
End Function
Private Function DefineDTTarefa(pTipo As Integer, pDataEvento As String) As String
   Dim sDTTarefa As String
   Dim TbTipo As Object
   
   Set TbTipo = CreateObject("Banco_3R.TB_OTIPOTAREFA")
   Set TbTipo.xDb = mvarSys.xDb
   With TbTipo
      If .Pesquisar(CH_IDTPTAREFA:=pTipo) Then
         sDTTarefa = Format(CDate(pDataEvento) + (.DIFHORA / 24), "DD/MM/YYYY hh:mm")
         If pTipo = 3 Or pTipo = 4 Or pTipo = 5 Then
            If Weekday(sDTTarefa) = 1 Then
               sDTTarefa = Format(CDate(sDTTarefa) + 1, "DD/MM/YYYY hh:mm")
            ElseIf Weekday(sDTTarefa) = 7 Then
               sDTTarefa = Format(CDate(sDTTarefa) + 2, "DD/MM/YYYY hh:mm")
            End If
            While eFeriado(mvarSys.xDb, sDTTarefa)
               sDTTarefa = Format(CDate(sDTTarefa) + 1, "DD/MM/YYYY hh:mm")
            Wend
         Else
            If Weekday(sDTTarefa) = 1 Then
               sDTTarefa = Format(CDate(sDTTarefa) - 2, "DD/MM/YYYY hh:mm")
            ElseIf Weekday(sDTTarefa) = 7 Then
               sDTTarefa = Format(CDate(sDTTarefa) - 1, "DD/MM/YYYY hh:mm")
            End If
            While eFeriado(mvarSys.xDb, sDTTarefa)
               sDTTarefa = Format(CDate(sDTTarefa) - 1, "DD/MM/YYYY hh:mm")
            Wend
          End If
      End If
    End With
    Set TbTipo = Nothing
    DefineDTTarefa = sDTTarefa
End Function
Public Sub GravaPrevisaoMeta()
   Dim Sql     As String
   Dim TbItem  As Object
   Dim dHoje   As Date
   Dim nVLPREV As Double
   
   dHoje = mvarSys.xDb.SysDate(1)
     
   Sql = "Select Min(DTITEM) As DTITEM "
   Sql = Sql & " From PMETAITEM"
   Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
   Sql = Sql & " And IDMETA =1"
   Sql = Sql & " And Month(DTITEM)=" & Month(dHoje)
   Sql = Sql & " And Year(DTITEM)=" & Year(dHoje)
   If mvarSys.xDb.Abretabela(Sql) Then
      
      Set TbItem = CreateObject("Banco_3R.TB_PMETAITEM")
      
      If IsDate(mvarSys.xDb.RsAux("DTITEM")) Then
         With TbItem
            Set .xDb = mvarSys.xDb
            If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDMETA:=1) Then
               mvarSys.Propriedades("VLPREV") = .VLPREV
            End If
         End With
      Else
         Sql = " Select Count(*) [VLPREV]"
         Sql = Sql & " From OEVENTOAGENDA E"
         Sql = Sql & " join OSERVICOEVT S On E.IDLOJA=S.IDLOJA and E.IDEVENTO=S.IDEVENTO"
         Sql = Sql & " Join OCLIENTE C On E.IDCLIENTE=C.IDCLIENTE And C.ISENTO=0"
         Sql = Sql & " Where E.IDLOJA=" & mvarIDLOJA
         Sql = Sql & " And day(E.STARTDATETIME) >= " & Day(dHoje)
         Sql = Sql & " And Month(E.STARTDATETIME) =" & Month(dHoje)
         Sql = Sql & " And Year(E.STARTDATETIME) = " & Year(dHoje)
         Sql = Sql & " And E.FLGCANCELADO<>1"
         Sql = Sql & " And S.IDTPSERVICO <>1"
         Sql = Sql & " Group By Year(E.STARTDATETIME), Month(E.STARTDATETIME)"
         Sql = Sql & " Order by Year(E.STARTDATETIME), Month(E.STARTDATETIME)"
         If mvarSys.xDb.Abretabela(Sql) Then
            nVLPREV = xVal(mvarSys.xDb.RsAux("VLPREV") & "")
         End If
      
         With TbItem
            Set .xDb = mvarSys.xDb
            .IDLOJA = mvarIDLOJA
            .IDMETA = 1
            .DTITEM = CStr(dHoje)
            .VLPREV = nVLPREV
            .Salvar
         End With
         mvarSys.Propriedades("VLPREV") = nVLPREV
      End If
   End If
End Sub
