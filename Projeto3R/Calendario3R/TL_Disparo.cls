VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_Disparo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe    As FrmDisparos
Attribute mvarMe.VB_VarHelpID = -1

Private m_lEditRow As Long
Private m_lEditCol As Long
Private m_Crescente As Boolean
Private m_TpManipulo As Boolean
Private CollNUMTRAT As Collection
Private CollTPMANIPULO As Collection

Private mvarSys            As Object
Private mvarIDLOJA         As Integer
Private mvarIDMAQUINA      As Integer
Private mvarIDATENDIMENTO  As Long
Public Property Let IDATENDIMENTO(ByVal vData As Long)
    mvarIDATENDIMENTO = vData
End Property
Public Property Get IDATENDIMENTO() As Long
   IDATENDIMENTO = mvarIDATENDIMENTO
End Property
Public Property Let IDMAQUINA(ByVal vData As Integer)
    mvarIDMAQUINA = vData
End Property
Public Property Get IDMAQUINA() As Integer
   IDMAQUINA = mvarIDMAQUINA
End Property
Public Property Let IDLOJA(ByVal vData As Integer)
    mvarIDLOJA = vData
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
   Set mvarMe = New FrmDisparos
End Sub
Private Sub Class_Terminate()
   Set mvarMe = Nothing
End Sub
Public Sub Show(Optional pMode As FormShowConstants = vbModal)
   DoEvents
   If mvarIDLOJA = 0 Then
      Call ExibirAviso("Código da Loja Inválido! (""" & mvarIDLOJA & """)")
      Exit Sub
   End If
   mvarMe.Show pMode
End Sub
Private Sub mvarMe_Activate()
   'mvarMe.GrdDisparos.Draw
'   MontaTela
  
   If GetTag(mvarMe, "1VEZ", "1") = "1" Then
      Screen.MousePointer = vbDefault
      mvarMe.MousePointer = vbDefault
      Call SetTag(mvarMe, "1VEZ", "0")
      
      Dim i As Integer
      Dim iRow As Integer
      With mvarMe.GrdDisparos
         .Redraw = False
         For i = 1 To .RowCount
            If xVal(.CellValue(i, "IDDISPARO") & "") = 0 Then
               Exit For
            End If
         Next
         iRow = IIf(i < 25, 1, i - 2)
         .RowVisible(iRow) = True
         iRow = IIf(.RowCount > i + 10, i + 10, i)
         .RowVisible(iRow) = True
         'Call .SetCurCell(1, "QTDDISPARO")
         Call .SetCurCell(i, "QTDDISPARO")
         .Redraw = True
      
         .SetFocus
      End With
      mvarMe.Refresh
   End If
End Sub
Private Sub mvarMe_CmbDTDisparoChange()
   If m_lEditRow > 0 Then
      mvarMe.GrdDisparos.CellValue(m_lEditRow, m_lEditCol) = Format(mvarMe.CmbDTDisparo.Value, "dd/mm/yyyy")
   End If
End Sub

Private Sub mvarMe_CmbDTDisparoLostFocus()
   If m_lEditRow > 0 Then
      mvarMe.GrdDisparos.CellValue(m_lEditRow, m_lEditCol) = Format(mvarMe.CmbDTDisparo.Value, "dd/mm/yyyy")
   End If
   mvarMe.CmbDTDisparo.Visible = False
End Sub

Private Sub mvarMe_CmbDTINICloseUp()
   DoEvents
   'Call RefreshGrid
   Call SetTag(mvarMe.GrdDisparos, "GRUPO", "")
   Call PopulaTela
End Sub
Private Sub mvarMe_CmbDTINILostFocus()
   'Call RefreshGrid
End Sub
Private Sub mvarMe_CmbDTINIValidate(Cancel As Boolean)
'   Call RefreshGrid
End Sub
Private Sub mvarMe_CmbIDMAQUINAClick()
   With mvarMe.CmbIDMAQUINA
      m_Crescente = (GetTag(.Tag, .ItemData(.ListIndex) & "_TPNUM", 0) = 1)
      m_TpManipulo = (GetTag(.Tag, .ItemData(.ListIndex) & "_TPMANIPULO", 0) = 1)
   
      Call mvarMe_CmbIDTPTRATAMENTOClick
   
      Call SetTag(mvarMe.CmbIDMAQUINA, "TEXT", .Text)
      mvarIDMAQUINA = .ItemData(.ListIndex)
   End With
   
End Sub
Private Sub mvarMe_CmbIDTPTRATAMENTOClick()
   Dim i As Integer
   
   If mvarMe.CmbIDTPTRATAMENTO.Visible Then
      For i = 1 To mvarMe.GrdDisparos.RowCount - 1
         If mvarMe.GrdDisparos.CellValue(i, "IDDISPARO") = 0 And xVal(mvarMe.GrdDisparos.CellValue(i, "NUMDISPARO") & "") = 0 Then
            Exit For
         End If
         If mvarMe.GrdDisparos.CellValue(i, "IDDISPARO0") = 0 Then
            If vbYes = ExibirPergunta("Existem lançamentos alterados." & vbNewLine & vbNewLine & "Desejar Salvá-los?", "Disaparos") Then
               Call SalvarDisparos
            End If
            Exit For
         End If
      Next
      Call PopulaTela
   End If
End Sub
Private Sub mvarMe_CmdCancelClick()
   Unload mvarMe
End Sub
Private Sub mvarMe_CmdDownClick()
   Dim lRow As Long
   Dim lCol As Long
   Dim lColKey As String
   Dim sVal As String
   Dim sTo  As String
   
   With mvarMe.GrdDisparos
      lRow = .CurRow
      lColKey = .ColKey(.CurCol)
      
      For lCol = 1 To .ColCount
         sVal = ""
         sTo = ""
         If Not IsNull(.CellValue(lRow, lCol)) Then sVal = .CellValue(lRow, lCol)
         If Not IsNull(.CellValue(lRow + 1, lCol)) Then sTo = .CellValue(lRow + 1, lCol)
         
         .CellValue(lRow, lCol) = sTo
         .CellValue(lRow + 1, lCol) = sVal
         If IsNumeric(sTo) Then .CellValue(lRow, lCol) = xVal(sTo)
         If IsNumeric(sVal) Then .CellValue(lRow + 1, lCol) = xVal(sVal)
         
         DoEvents
         .SetFocus
         Call .SetCurCell(lRow + 1, lColKey)
      Next
   End With
End Sub
Private Sub mvarMe_CmdExcluirClick()
   Call ExcluirDisparo(mvarMe.GrdDisparos.CurRow)
End Sub
Private Sub mvarMe_CmdOkClick()
   If Not VerificaDisparos() Then Exit Sub
   If SalvarDisparos Then
      Unload mvarMe
   End If
End Sub
Private Function VerificaDisparos() As Boolean
   Dim i As Long
   
   With mvarMe.GrdDisparos
      If Val(.CellValue(1, "IDAREA")) = 0 And .CellText(1, "IDAREA") = "" Then
         Call ExibirInformacao("Área inválida.")
         Call .SetCurCell(1, "IDAREA")
         .SetFocus
         Exit Function
      End If
   End With
   If Trim(mvarMe.CmbIDMAQUINA.Text) = "" Or mvarMe.CmbIDMAQUINA.ListIndex = -1 Then
      Call ExibirInformacao("Informe a máquina de operação.")
      mvarMe.CmbIDMAQUINA.Enabled = True
      mvarMe.CmbIDMAQUINA.SetFocus
      Exit Function
   End If

   VerificaDisparos = True
End Function
Private Function SalvarDisparos() As Boolean
   Dim i          As Integer
   Dim cQueries   As Collection
   Dim sTag       As String
   Dim TbSessao   As Object 'TB_OSESSAO
   Dim TbDisparo  As Object 'TB_OMAQDISPAROS
  
   Set cQueries = New Collection
   '* Itens
      
      With mvarMe.GrdDisparos
         For i = 1 To .RowCount - 1
         'If i = 1153 Then
         '   i = i
         'End If
         
            sTag = ""
            If Not IsMissing(.RowTag(i)) Then
               sTag = .RowTag(i)
            End If
            If .CellValue(i, "IDDISPARO") = 0 And xVal(.CellValue(i, "NUMDISPARO") & "") = 0 Then
               Exit For
            End If
            If xVal(GetTag(sTag, "ISDIRT", 0)) <> 0 Then
               '****************
               '* Salvar Sessão
               If Val(.CellValue(i, "IDLOJA")) <> 0 _
                  And Val(.CellValue(i, "IDATENDIMENTO") & "") <> 0 _
                  And Val(.CellValue(i, "IDSESSAO") & "") <> 0 Then
                                             
                  Set TbSessao = CriarObjeto("BANCO_3R.TB_OSESSAO")
                  Set TbSessao.xDb = mvarSys.xDb
                  If TbSessao.Pesquisar(Ch_IDLOJA:=Val(.CellValue(i, "IDLOJA")), Ch_IDATENDIMENTO:=Val(.CellValue(i, "IDATENDIMENTO")), Ch_IDSESSAO:=Val(.CellValue(i, "IDSESSAO"))) Then
                     TbSessao.DISPAROS = xVal(.CellValue(i, "QTDDISPARO") & "")
                     cQueries.Add TbSessao.QryUpDate
                  End If
               End If
               
               '****************
               '* Salvar Disparo
               Set TbDisparo = CriarObjeto("BANCO_3R.TB_OMAQDISPAROS")
               TbDisparo.xDb = mvarSys.xDb
               If Not IsDate(.CellValue(i, "DTDISPARO")) Then
                  Call ExibirStop("Data Inválida")
                  Call .SetCurCell(i, "DTDISPARO")
                  Exit Function
               End If
               TbDisparo.DTDISPARO = .CellValue(i, "DTDISPARO")
               TbDisparo.FLGDELETE = 0
               TbDisparo.IDAREA = xVal(.CellValue(i, "IDAREA") & "")
               TbDisparo.IDATENDIMENTO = xVal(.CellValue(i, "IDATENDIMENTO") & "")
               TbDisparo.IDDISPARO = xVal(.CellValue(i, "IDDISPARO") & "")
               TbDisparo.IDLOJA = mvarIDLOJA
               TbDisparo.IDMAQUINA = mvarIDMAQUINA 'mvarMe.CmbIDMAQUINA.ItemData(mvarMe.CmbIDMAQUINA.ListIndex)
               TbDisparo.IDSESSAO = xVal(.CellValue(i, "IDSESSAO") & "")
               TbDisparo.IDTPTRATAMENTO = xVal(.CellValue(i, "IDTPTRATAMENTO") & "")
               TbDisparo.IDMANIPULO = xVal(.CellValue(i, "IDMANIPULO") & "")
               TbDisparo.IDTPMANIPULO = xVal(CollTPMANIPULO("k" & TbDisparo.IDMANIPULO))
               
               TbDisparo.IDLAMP = xVal(.CellValue(i, "IDLAMP") & "")
               
               TbDisparo.NUMDISPARO = xVal(.CellValue(i, "NUMDISPARO"))
               TbDisparo.QTDDISPARO = xVal(.CellValue(i, "QTDDISPARO"))
               
               If TbDisparo.IDLAMP = 0 Then TbDisparo.IDLAMP = 1
               If TbDisparo.IDMANIPULO = 0 Then TbDisparo.IDMANIPULO = 1
               If TbDisparo.IDTPMANIPULO = 0 Then TbDisparo.IDTPMANIPULO = 1
               If TbDisparo.IDTPTRATAMENTO = 0 Then TbDisparo.IDTPTRATAMENTO = 1
                                             
               If TbDisparo.isDirt Then
                  cQueries.Add TbDisparo.QrySave
               End If
            End If
         Next
      End With
               
   If cQueries.Count = 0 Then
      SalvarDisparos = True
   Else
      If mvarSys.xDb.Executa(cQueries) Then
         'i = mvarMe.CmbIDTPTRATAMENTO.ListIndex
         'mvarMe.CmbIDTPTRATAMENTO.ListIndex = -1
         'mvarMe.CmbIDTPTRATAMENTO.ListIndex = i
         Call PopulaTela
         SalvarDisparos = True
      Else
         Call ExibirInformacao("Erro ao gravar itens e produtos.")
      End If
   End If
        
Saida:
   Set TbSessao = Nothing
   Set TbDisparo = Nothing
   Set cQueries = Nothing
End Function
Private Sub mvarMe_CmdSalvarClick()
   If Not VerificaDisparos() Then Exit Sub
   Call ExibirResultado(mvarSys, SalvarDisparos)
End Sub
Private Sub mvarMe_CmdUpClick()
   Dim lRow As Long
   Dim lCol As Long
   Dim lColKey As String
   Dim sVal As String
   Dim sTo  As String
   
   With mvarMe.GrdDisparos
      lRow = .CurRow
      lColKey = .ColKey(.CurCol)
   
      For lCol = 1 To .ColCount
         sVal = ""
         sTo = ""
         If Not IsNull(.CellValue(lRow, lCol)) Then sVal = .CellValue(lRow, lCol)
         If Not IsNull(.CellValue(lRow - 1, lCol)) Then sTo = .CellValue(lRow - 1, lCol)
         
         .CellValue(lRow, lCol) = sTo
         .CellValue(lRow - 1, lCol) = sVal
         If IsNumeric(sTo) Then .CellValue(lRow, lCol) = xVal(sTo)
         If IsNumeric(sVal) Then .CellValue(lRow - 1, lCol) = xVal(sVal)
         
         DoEvents
         .SetFocus
         Call .SetCurCell(lRow - 1, lColKey)
      Next
   End With
End Sub
Private Sub mvarMe_GrdDisparosAfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
   Dim nValor  As Integer
   Dim sMsg    As String
   Dim i       As Integer
   Dim Sql     As String
   Dim nIDLAMP As Integer
   
   Screen.MousePointer = vbHourglass
   With mvarMe.GrdDisparos
      If .ColKey(lCol) = "IDMANIPULO" Then
         If xVal(.CellValue(lRow, "IDDISPARO0")) = 0 And .RowCount > lRow Then
            If .CellValue(lRow, "IDMANIPULO") <> .CellValue(lRow + 1, "IDMANIPULO") Then
               .CellValue(lRow, "IDTPMANIPULO") = CollTPMANIPULO("k" & .CellValue(lRow, "IDMANIPULO"))
               Sql = "Select isNull(Max(IDLAMP),0) [IDLAMP] "
               Sql = Sql & " From OLAMPADA "
               Sql = Sql & " Where IDMANIPULO=" & .CellValue(lRow, "IDMANIPULO")
               Sql = Sql & " And IDLOJA=" & mvarIDLOJA
               nIDLAMP = 0
               If mvarSys.xDb.AbreTabela(Sql) Then
                  nIDLAMP = xVal(mvarSys.xDb.RsAux("IDLAMP"))
               End If
               If nIDLAMP <= 0 Then
                  Sql = "Select isNull(Max(IDLAMP),0) [IDLAMP] "
                  Sql = Sql & " From OLAMPADA "
                  Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
                  Call mvarSys.xDb.AbreTabela(Sql)
                  
                  Dim TbLamp As Object
                  Set TbLamp = CriarObjeto("BANCO_3R.TB_OLAMPADA")
                  Set TbLamp.xDb = mvarSys.xDb
                  TbLamp.IDMANIPULO = .CellValue(lRow, "IDMANIPULO")
                  TbLamp.IDLAMP = xVal(mvarSys.xDb.RsAux("IDLAMP")) + 1
                  TbLamp.IDLOJA = mvarIDLOJA
                  If TbLamp.Salvar Then
                     nIDLAMP = TbLamp.IDLAMP
                  End If
               End If
               For i = lRow To mvarMe.GrdDisparos.RowCount - 1
                  .CellValue(i, "IDMANIPULO") = .CellValue(lRow, "IDMANIPULO")
                  .CellValue(i, "IDTPMANIPULO") = CollTPMANIPULO("k" & .CellValue(i, "IDMANIPULO"))
                  .CellValue(i, "IDLAMP") = nIDLAMP
               Next
            End If
         End If
      End If
   
      If InArray(.ColKey(lCol), Array("NUMDISPARO", "QTDDISPARO")) Then
         If IsNumeric(.CellValue(lRow, lCol)) And (xVal(.CellValue(lRow, lCol) & "") <> 0 Or .CellValue(lRow, "IDTPSERVICO") = 1) Then
            Call RenumerarDisparo(lRow, lCol)
            Call CalculaQtdDisparo(lRow, lCol)
            Call DefineCelCorrente(lRow, lCol)

         If InArray(.ColKey(lCol), Array("QTDDISPARO", "NUMDISPARO")) Then
            If m_Crescente Then
               Call RenumerarDisparo(lRow, lCol)
               Call CalculaQtdDisparo(lRow, lCol)
               Call DefineCelCorrente(lRow, lCol)
                     
            Else
               If xVal(.CellValue(lRow, "NUMDISPARO") & "") <= 0 Then
                  sMsg = "Número (" & ValBr(xVal(.CellValue(lRow, "NUMDISPARO") & ""), 0) & ") do registro ínvalido." & vbNewLine & vbNewLine
                  sMsg = sMsg & "Favor entrar em contato com administrador para adicionar disparos na máquina" & vbNewLine & vbNewLine
                  sMsg = sMsg & "Entre abaixo, com o número de disparos adicionados na máquina."
                  
                  nValor = xVal(InputBox(sMsg, "Disparos", 0))
                  If nValor > 0 Then
                     Call IncluiGrdLinha(True, lRow)
                     
                     With mvarMe.GrdDisparos
                        For i = 1 To .ColCount
                           .CellValue(lRow, i) = .CellValue(lRow - IIf(lRow = 1, 0, 1), i)
                        Next
                        .CellValue(lRow, "IDAREA") = 0
                        .CellValue(lRow, "NUMDISPARO") = ValBr(nValor, 0) + .CellValue(lRow - 1, "NUMDISPARO")
                        .CellValue(lRow, "QTDDISPARO") = 0
                     End With
                     Call RenumerarDisparo(lRow, lCol)
                     Call CalculaQtdDisparo(lRow, lCol)
                     Call DefineCelCorrente(lRow + IIf(lRow = .RowCount, 0, 1), lCol)
                  End If
               End If
            End If
            If xVal(.CellValue(lRow, "QTDDISPARO")) = 0 Then
               If xVal(.CellValue(lRow, "IDTPSERVICO") & "") > 1 Then '* Serviço = Teste
                  DoEvents
                  ExibirAviso "Valor inválido." & vbNewLine & vbNewLine & "Somente testes podem ter zero disparos."
                  .CellValue(lRow, "IDDISPARO") = ""
                  .CellValue(lRow, "NUMDISPARO") = ""
                  .CellValue(lRow, "QTDDISPARO") = ""
                  .SetCurCell lRow, lCol
                  
               End If
            End If
            
         End If
            
         Else
            If .CellValue(lRow, "IDCLIENTE") = 0 And .CellValue(lRow, "IDAREA") <> 0 Then
               Call ExibirAviso("Valor inválido.")
            Else
               Call RenumerarDisparo(lRow, lCol)
            End If
            .CellValue(lRow, lCol) = 0
            Call .SetCurCell(lRow, lCol)
         End If
      End If
   End With
   Screen.MousePointer = vbDefault
End Sub
Private Sub RenumerarDisparo(lRow As Long, lCol As Long)
   Dim bReNum  As Boolean
   Dim i As Long
   
   With mvarMe.GrdDisparos
      If xVal(.CellValue(lRow, "IDDISPARO")) = 0 Then
         bReNum = True
      Else
         If lRow < .RowCount - 1 Then
            If xVal(.CellValue(lRow, "IDDISPARO")) <> xVal(.CellValue(lRow, "IDDISPARO")) + 1 Then
               bReNum = True
            End If
         End If
      End If
      If bReNum Then
         For i = lRow To .RowCount - 1
            If xVal(.CellValue(i, "NUMDISPARO") & "") <> 0 Or xVal(.CellValue(i, "QTDDISPARO") & "") <> 0 Or .CellValue(i, "IDTPSERVICO") = 1 Then
               If i = 1 Then
                  If xVal(.CellValue(i, "IDDISPARO")) = 0 Then
                     If xVal(.CellValue(i, "ITEM") & "") = 0 Then
                        .CellValue(i, "IDDISPARO") = 1
                     Else
                        .CellValue(i, "IDDISPARO") = .CellValue(i, "ITEM")
                     End If
                  End If
               Else
                  If xVal(.CellValue(i - 1, "IDDISPARO0")) = 0 And xVal(.CellValue(i - 1, "IDDISPARO")) = 0 Then
                     Exit Sub
                  End If
                  If xVal(.CellValue(i, "IDDISPARO0")) = 0 Then
                     If xVal(.CellValue(i, "IDDISPARO")) = 0 And .CellValue(i, "ITEM") > .CellValue(i - 1, "IDDISPARO") Then
                        .CellValue(i, "IDDISPARO") = .CellValue(i, "ITEM")
                     ElseIf .CellValue(i, "IDDISPARO") <> .CellValue(i, "ITEM") Then
                        .CellValue(i, "IDDISPARO") = .CellValue(i - 1, "IDDISPARO") + 1
                     End If
                     
                  End If
'                  If xVal(.CellValue(i - 1, "IDDISPARO")) <= 0 Then
'                     .CellValue(i, "IDDISPARO") = "0"
'                  Else
'                     .CellValue(i, "IDDISPARO") = .CellValue(i - 1, "IDDISPARO") + 1
'                  End If
               End If
               Call CalculaQtdDisparo(i, lCol)
               If xVal(.CellValue(i, "NUMDISPARO") & "") = 0 And xVal(.CellValue(i, "QTDDISPARO") & "") = 0 And xVal(.CellValue(i, "IDCLIENTE")) <> 0 Then
                  .CellValue(i, "IDDISPARO") = "0"
               End If
            Else
               If xVal(.CellValue(i, "IDCLIENTE")) = 0 Or (xVal(.CellValue(i, "NUMDISPARO") & "") = 0 And xVal(.CellValue(i, "QTDDISPARO") & "") = 0) Then
                  Exit For
               Else
                  .CellValue(i, "IDDISPARO") = "0"
               End If
            End If
         Next
      End If
   End With
End Sub
Private Sub CalculaQtdDisparo(lRow As Long, lCol As Long)
   Dim i As Integer
   Dim bAchou As Boolean
   Dim Sql As String
   Dim xNumDisparo As Long
   
   With mvarMe.GrdDisparos
      If .ColKey(lCol) = "NUMDISPARO" Then
         If lRow = 1 Then
            If m_Crescente Then
               .CellValue(lRow, "QTDDISPARO") = .CellValue(lRow, "NUMDISPARO")
            Else
               .CellValue(lRow, "QTDDISPARO") = ValBr(0, 0)
            End If
         Else
            If xVal(.CellValue(lRow, "NUMDISPARO") & "") = 0 Or xVal(.CellValue(lRow - 1, "NUMDISPARO") & "") = 0 Or (xVal(.CellValue(lRow - 1, "QTDDISPARO") & "") = 0 And Not m_Crescente) Then
               If xVal(.CellValue(lRow, "ITEM") & "") = xVal(.CellValue(lRow, "IDDISPARO") & "") Then
                  .CellValue(lRow, "QTDDISPARO") = ValBr(0, 0)
               Else
                  .CellValue(lRow, "QTDDISPARO") = ""
               End If
            Else
               If lRow > 2 Then
                  xNumDisparo = xVal(.CellValue(lRow - 1, "NUMDISPARO"))
               Else '*Recuperar último registro deste manípulo
               End If
               If m_TpManipulo Then
                  For i = lRow - 1 To 1 Step -1
                     If .CellValue(lRow, "IDTPMANIPULO") = .CellValue(i, "IDTPMANIPULO") Then
                        xNumDisparo = xVal(.CellValue(i, "NUMDISPARO") & "")
                        bAchou = True
                        Exit For
                     End If
                  Next
                  If Not bAchou Then
                     xNumDisparo = 0
                     Sql = "Select IDDISPARO, QTDDISPARO, NUMDISPARO "
                     Sql = Sql & " FROM OMAQDISPARO"
                     Sql = Sql & " Where FLGDELETE=0"
                     Sql = Sql & " And IDLOJA=" & mvarIDLOJA
                     Sql = Sql & " And IDTPMANIPULO=" & .CellValue(lRow, "IDTPMANIPULO")
                     Sql = Sql & " And IDDISPARO = (Select Max(IDDISPARO)"
                     Sql = Sql & "     From OMAQDISPARO"
                     Sql = Sql & "     Where IDLOJA=" & mvarIDLOJA
                     Sql = Sql & "     And FLGDELETE=0"
                     Sql = Sql & "     And IDTPMANIPULO=" & .CellValue(lRow, "IDTPMANIPULO")
                     Sql = Sql & "     And  IDDISPARO<" & .CellValue(lRow - 1, "IDDISPARO")
                     Sql = Sql & "     )"
                     If mvarSys.xDb.AbreTabela(Sql) Then
                        xNumDisparo = xVal(mvarSys.xDb.RsAux("NUMDISPARO"))
                     End If
                  End If
               End If
               If m_Crescente Then
                  .CellValue(lRow, "QTDDISPARO") = ValBr(xVal(.CellValue(lRow, "NUMDISPARO")), 0) - xNumDisparo
               Else
                  .CellValue(lRow, "QTDDISPARO") = ValBr(xNumDisparo - xVal(.CellValue(lRow, "NUMDISPARO")), 0)
               End If
            End If
            If xVal(.CellValue(lRow, "QTDDISPARO")) < 0 Then
               If xVal(.CellValue(lRow, "IDAREA")) <> 0 Then
                  .CellValue(lRow, "QTDDISPARO") = ValBr(0, 0)
               End If
            End If
         End If
      End If
      If .ColKey(lCol) = "QTDDISPARO" Then
         If lRow = 1 Then
            If m_Crescente Then
               .CellValue(lRow, "QTDDISPARO") = ValBr(.CellValue(lRow, "QTDDISPARO"), 0)
               .CellValue(lRow, "NUMDISPARO") = ValBr(xVal(.CellValue(lRow, "QTDDISPARO")), 0)
            Else
               .CellValue(lRow, "QTDDISPARO") = ValBr(0, 0)
            End If
            
         Else
            If xVal(.CellValue(lRow - 1, "NUMDISPARO") & "") = 0 Or (xVal(.CellValue(lRow - 1, "QTDDISPARO") & "") = 0 And xVal(.CellValue(lRow - 1, "IDTPSERVICO") & "") > 1) Or IsNull(.CellValue(lRow, "QTDDISPARO")) Then
               .CellValue(lRow, "NUMDISPARO") = ValBr(0, 0)
            Else
               If xVal(.CellValue(lRow, "QTDDISPARO") & "") <> 0 Or IsNull(.CellValue(lRow, "QTDDISPARO")) Then
                  xNumDisparo = xVal(.CellValue(lRow - 1, "NUMDISPARO"))
                  If m_TpManipulo Then
                     For i = lRow - 1 To 1 Step -1
                        If .CellValue(lRow, "IDTPMANIPULO") = .CellValue(i, "IDTPMANIPULO") Then
                           xNumDisparo = xVal(.CellValue(i, "NUMDISPARO") & "")
                           bAchou = True
                           Exit For
                        End If
                     Next
                     If Not bAchou Then
                        xNumDisparo = 0
                        Sql = "Select IDDISPARO, QTDDISPARO, NUMDISPARO "
                        Sql = Sql & " FROM OMAQDISPARO"
                        Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
                        Sql = Sql & " And FLGDELETE=0"
                        Sql = Sql & " And IDTPMANIPULO=" & .CellValue(lRow, "IDTPMANIPULO")
                        Sql = Sql & " And IDDISPARO = (Select Max(IDDISPARO)"
                        Sql = Sql & "     From OMAQDISPARO"
                        Sql = Sql & "     Where IDLOJA=" & mvarIDLOJA
                        Sql = Sql & "     And FLGDELETE=0"
                        Sql = Sql & "     And IDTPMANIPULO=" & .CellValue(lRow, "IDTPMANIPULO")
                        Sql = Sql & "     And IDDISPARO<=" & .CellValue(lRow - 1, "IDDISPARO")
                        Sql = Sql & "     )"
                        If mvarSys.xDb.AbreTabela(Sql) Then
                           xNumDisparo = xVal(mvarSys.xDb.RsAux("NUMDISPARO"))
                        End If
                     End If
                  End If
                  If m_Crescente Then
                     .CellValue(lRow, "NUMDISPARO") = ValBr(xNumDisparo + xVal(.CellValue(lRow, "QTDDISPARO")), 0)
                  Else
                     .CellValue(lRow, "NUMDISPARO") = ValBr(xNumDisparo - xVal(.CellValue(lRow, "QTDDISPARO")), 0)
                  End If
               End If
            End If
            'If xVal(.CellValue(lRow, "NUMDISPARO")) < 0 Then
            '   .CellValue(lRow, "NUMDISPARO") = ValBr(0, 0)
            'End If
         End If
      End If
   End With
End Sub
Private Sub DefineCelCorrente(lRow As Long, lCol As Long)
   With mvarMe.GrdDisparos
      If .ColKey(lCol) = "QTDDISPARO" Or .ColKey(lCol) = "NUMDISPARO" Then
         If lRow < .RowCount - 1 Then
            If .CellValue(lRow + 1, "IDAREA") & "" = "" Then
               Call .SetCurCell(lRow + 1, .ColIndex("IDAREA"))
            ElseIf .CellValue(lRow + 1, "DTDISPARO") & "" = "" Then
               Call .SetCurCell(lRow + 1, .ColIndex("DTDISPARO"))
            Else
               If .ColKey(lCol) = "QTDDISPARO" Then
                  Call .SetCurCell(lRow + 1, .ColIndex("QTDDISPARO"))
               ElseIf .ColKey(lCol) = "NUMDISPARO" Then
                  Call .SetCurCell(lRow + 1, .ColIndex("NUMDISPARO"))
               Else
                  Call .SetCurCell(lRow + 1, .ColIndex("NUMDISPARO"))
               End If
            End If
         End If
      Else
         If .CellValue(lRow, "IDAREA") = "" Then
            Call .SetCurCell(lRow, .ColIndex("IDAREA"))
         ElseIf .CellValue(lRow, "DTDISPARO") = "" Then
            Call .SetCurCell(lRow, .ColIndex("DTDISPARO"))
         Else
            Call .SetCurCell(lRow, .ColIndex("QTDDISPARO"))
         End If
      End If
   End With
End Sub
Private Sub mvarMe_GrdDisparosBeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
   With mvarMe.GrdDisparos
      If lConvErr <> 0 Then
         If .CellValue(lRow, "IDCLIENTE") = 0 And .CellValue(lRow, "IDAREA") <> 0 Then
            ExibirAviso "Valor inválido."
            eResult = igEditResCancel
         End If
      Else
         If xVal(.CellValue(lRow, "IDSESSAO") & "") <> 0 Then
            If .ColKey(lCol) = "IDAREA" Then
               Call ExibirStop("Item relacionado ao atendimento e não pode ser alterado.")
               eResult = igEditResCancel
            ElseIf InArray(.ColKey(lCol), Array("QTDDISPARO", "NUMDISPARO")) Then
               If vNewValue = 0 Then
                  If xVal(.CellValue(lRow, "IDTPSERVICO") & "") > 1 Then '* Serviço = Teste
                     DoEvents
                     ExibirAviso "Valor inválido." & vbNewLine & vbNewLine & "Somente testes podem ter zero disparos."
                     eResult = igEditResCancel
                  End If
               End If
            End If
         ElseIf InArray(.ColKey(lCol), Array("QTDDISPARO", "NUMDISPARO")) Then
            If xVal(.CellValue(lRow, "IDTPTRATAMENTO") & "") = 0 Then
               ExibirAviso "Tratamento inválido."
               eResult = igEditResCancel
            End If
            If xVal(.CellValue(lRow, "IDMANIPULO") & "") = 0 Then
               ExibirAviso "Manípulo inválido."
               eResult = igEditResCancel
            End If
          ElseIf InArray(.ColKey(lCol), Array("QTDDISPARO", "NUMDISPARO")) Then
            If xVal(.CellValue(lRow, "IDTPMANIPULO") & "") = 0 Then
               ExibirAviso "Manípulo inválido."
               eResult = igEditResCancel
            End If
         End If
      End If
   End With
   Call DirtRow(lRow, eResult)
End Sub
Private Sub DirtRow(lRow As Long, Optional eResult As iGrid251_75B4A91C.EEditResults = igEditResCommit)
   Dim sTag As String
   
   If eResult = igEditResCommit Then
      sTag = ""
      If Not IsMissing(mvarMe.GrdDisparos.RowTag(lRow)) Then
         sTag = mvarMe.GrdDisparos.RowTag(lRow)
      End If
      mvarMe.GrdDisparos.RowTag(lRow) = SetTag(sTag, "ISDIRT", 1)
   End If
End Sub

Private Sub mvarMe_GrdDisparosCellSelectionChange(ByVal lRow As Long, ByVal lCol As Long, ByVal bSelected As Boolean)
   If xVal(mvarMe.GrdDisparos.CellValue(lRow, "IDDISPARO")) = 0 Then
      If lRow > 1 Then
         mvarMe.CmdUp.Enabled = (xVal(mvarMe.GrdDisparos.CellValue(lRow, "IDCLIENTE")) = xVal(mvarMe.GrdDisparos.CellValue(lRow - 1, "IDCLIENTE")))
      End If
      If lRow < mvarMe.GrdDisparos.RowCount Then
         mvarMe.CmdDown.Enabled = (xVal(mvarMe.GrdDisparos.CellValue(lRow, "IDCLIENTE")) = xVal(mvarMe.GrdDisparos.CellValue(lRow + 1, "IDCLIENTE")))
      End If

      If lRow = 1 Then
         mvarMe.CmdUp.Enabled = False
         mvarMe.CmdDown.Enabled = True
      ElseIf lRow = mvarMe.GrdDisparos.RowCount - 1 Then
         mvarMe.CmdUp.Enabled = True
         mvarMe.CmdDown.Enabled = False
      ElseIf lRow = mvarMe.GrdDisparos.RowCount Then
         mvarMe.CmdUp.Enabled = False
         mvarMe.CmdDown.Enabled = False
      End If
   Else
      mvarMe.CmdUp.Enabled = False
      mvarMe.CmdDown.Enabled = False
   End If
   mvarMe.CmdUp.Picture = mvarMe.ImageList1.ListImages(IIf(mvarMe.CmdUp.Enabled, "UP", "UP_D")).Picture
   mvarMe.CmdDown.Picture = mvarMe.ImageList1.ListImages(IIf(mvarMe.CmdDown.Enabled, "DOWN", "DOWN_D")).Picture
   'DoEvents
   'mvarMe.GrdDisparos.SetFocus
   
End Sub

Private Sub mvarMe_GrdDisparosColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
   bDoDefault = False
End Sub
Private Sub mvarMe_GrdDisparosColHeaderDblClick(ByVal lCol As Long)
   Dim i As Integer
   With mvarMe.GrdDisparos
      'If .ColKey(lCol) = "FOTOTIPO" Then
      'End If
   End With
End Sub

Private Sub mvarMe_GrdDisparosDblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)
   If (lRow = mvarMe.GrdDisparos.RowCount) Then bRequestEdit = False
   If mvarMe.GrdDisparos.ColKey(lCol) = "IDDISPARO" Then
      If xVal(mvarMe.GrdDisparos.CellValue(lRow, "IDATENDIMENTO") & "") <> 0 Then
         Dim TlAtend As Object
         Set TlAtend = CriarObjeto("Calendario3R.TL_ATENDIMENTO")
         With TlAtend
            Set .Sys = mvarSys
            .IDLOJA = mvarIDLOJA
            .IDATENDIMENTO = xVal(mvarMe.GrdDisparos.CellValue(lRow, "IDATENDIMENTO"))
            If .IDEVENTO = 0 Then
               Call ExibirInformacao("Atendimento '" & .IDATENDIMENTO & "' não existe.", "Atendimento")
            Else
               .Show
            End If
         End With
         Set TlAtend = Nothing
      End If
   End If
End Sub

Private Sub mvarMe_GrdDisparosKeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
   If Shift = 1 Or Shift = 2 Or Shift = 4 Then
      mvarMe.GrdDisparos.Redraw = False
      mvarMe.GrdDisparos.Editable = False
      mvarMe.CmbDTINI.SetFocus
      
      If KeyCode = vbKeyDown Then
         'MsgBox "1"
         If mvarMe.CmdDown.Enabled Then Call mvarMe_CmdDownClick
      ElseIf KeyCode = vbKeyUp Then
         'MsgBox "2"
          If mvarMe.CmdUp.Enabled Then Call mvarMe_CmdUpClick
      End If
      KeyCode = 0
      mvarMe.GrdDisparos.SetFocus
      mvarMe.GrdDisparos.Editable = True
      mvarMe.GrdDisparos.Redraw = True
   End If
End Sub

Private Sub mvarMe_GrdDisparosMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
   With mvarMe.GrdDisparos
      If lRow = .RowCount Then
         If lRow > 1 Then
            If Val(.CellValue(lRow - 1, "IDAREA")) = 0 Then Exit Sub
         End If
         Call IncluiGrdLinha
      End If
   End With
End Sub
Private Sub mvarMe_GrdDisparosQuitCustomEdit()
   mvarMe.CmbDTDisparo.Visible = False
End Sub
Private Sub mvarMe_GrdDisparosRequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
   Dim nLin As Long
   Dim i As Integer
   
   With mvarMe.GrdDisparos
      bCancel = (GetTag(.ColTag(lCol), "EDIT", True) = False)
      'If .CellFmtString(lRow, lCol) <> "" Then sText = .CellValue(lRow, lCol)
      eTextEditOpt = GetTag(.ColTag(lCol), "EDITOPT", 0)
            
      If .ColKey(lCol) = "DTDISPARO" Then
         nLin = IIf(lRow = 1, 2, lRow - 1)
         If IsDate(.CellValue(nLin, "DTDISPARO")) Then
            mvarMe.CmbDTDisparo.Value = .CellValue(nLin, "DTDISPARO")
         End If
         Call MoveDTDisparo(lRow, lCol, bCancel)
      End If
      If InArray(.ColKey(lCol), Array("DTDISPARO", "IDTPTRATAMENTO", "IDAREA")) Then
         If xVal(.CellValue(lRow, "IDSESSAO") & "") <> 0 Then
            mvarMe.CmbDTDisparo.Visible = False
            Call ExibirStop("Item relacionado ao atendimento e não pode ser alterado.")
            bCancel = True
         End If
      End If
   End With
End Sub
Private Sub MoveDTDisparo(ByVal lRow As Long, ByVal lCol As Long, ByRef bCancel As Boolean)
   Dim lLeft As Long
   Dim lTop As Long
   Dim lWidth As Long
   Dim lHeight As Long
   
   m_lEditRow = lRow
   m_lEditCol = lCol
   
   ' disabling built-in editor
   bCancel = True
   
   ' calculating the position of the current cell
   With mvarMe.GrdDisparos
      .CellBoundary lRow, lCol, lLeft, lTop, lWidth, lHeight
   
      mvarMe.CmbDTDisparo.Left = .Left + (lLeft) * Screen.TwipsPerPixelX
      mvarMe.CmbDTDisparo.Top = .Top + (lTop) * Screen.TwipsPerPixelY
      mvarMe.CmbDTDisparo.Width = lWidth * Screen.TwipsPerPixelX
      mvarMe.CmbDTDisparo.Height = lHeight * Screen.TwipsPerPixelY

      If Not IsDate(.CellValue(lRow, lCol)) Then
         .CellValue(lRow, lCol) = Format(mvarMe.CmbDTDisparo.Value, "dd/mm/yyyy")
      End If
      mvarMe.CmbDTDisparo.Value = .CellValue(lRow, lCol)
      mvarMe.CmbDTDisparo.ZOrder 0
   End With
   mvarMe.CmbDTDisparo.Visible = True
   mvarMe.CmbDTDisparo.SetFocus
End Sub

Private Sub mvarMe_GrdDisparosValidate(Cancel As Boolean)
   mvarMe.GrdDisparos.CommitEdit
End Sub

Private Sub mvarMe_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim Sql As String
   Dim MyTb As Object
   
'   If Shift = 1 Or Shift = 2 Then
'      If KeyCode = vbKeyF1 Then
'         Sql = "Delete From OTPMANIPULO"
'         Call mvarsys.xdb.Executa(Sql)
'
'         Sql = "Alter Table OTPMANIPULO Alter Column DSCMANIPULO Varchar(20) Null" & vbNewLine
'         Call mvarsys.xdb.Executa(Sql)
'
'         Sql = "SET IDENTITY_INSERT OTPMANIPULO on;" & vbNewLine
'         Sql = Sql & "Insert Into OTPMANIPULO (IDTPMANIPULO, DSCMANIPULO) Values (1, 'Pequeno');" & vbNewLine
'         Sql = Sql & "Insert Into OTPMANIPULO (IDTPMANIPULO, DSCMANIPULO) Values (2, 'Grande');" & vbNewLine
'         Sql = Sql & "SET IDENTITY_INSERT OTPMANIPULO off;" & vbNewLine
'         Call mvarsys.xdb.Executa(Sql)
'      End If
'   End If
End Sub

Private Sub mvarMe_KeyUp(KeyCode As Integer, Shift As Integer)
   Dim lRow As Long
   Dim nID  As Long
   Dim i    As Long
   
   If mvarMe.ActiveControl Is mvarMe.GrdDisparos Then
      With mvarMe.GrdDisparos
         If .CurRow <= 0 Then Exit Sub
         lRow = .CurRow
         If KeyCode = vbKeyDelete Then
            Call ExcluirDisparo(lRow)

         ElseIf KeyCode = vbKeyInsert Then
            IncluiGrdLinha vRowBefore:=lRow
            Call DirtRow(lRow)
            
         ElseIf KeyCode = vbKeyF5 Then
            Call RefreshGrid
         End If
      End With
   End If
End Sub
Public Function ExcluirDisparo(lRow As Long) As Boolean
   Dim Sql As String
   Dim bMax As Boolean
   Dim bAtend As Boolean
   With mvarMe.GrdDisparos
      Dim TBAtend As Object
      Set TBAtend = CriarObjeto("BANCO_3R.TB_OATENDIMENTO")
      Set TBAtend.xDb = mvarSys.xDb
      If Not TBAtend.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDATENDIMENTO:=xVal(.CellValue(lRow, "IDATENDIMENTO") & "")) Then
         .CellValue(lRow, "IDSESSAO") = ""
      End If
      
      bMax = False
      Sql = "Select Max(IDDISPARO) [MX]"
      Sql = Sql & " From OMAQDISPAROS"
      Sql = Sql & " Where FLGDELETE=0"
      Sql = Sql & " And IDLOJA=" & mvarIDLOJA
      Sql = Sql & " And IDMAQUINA=" & mvarIDMAQUINA
      If mvarSys.xDb.AbreTabela(Sql) Then
         bMax = (xVal(mvarSys.xDb.RsAux("MX") & "") = xVal(.CellValue(lRow, "IDDISPARO") & ""))  'And xVal(.CellValue(lRow, "IDSESSAO")) <> 0)
      End If
      
      bAtend = False
      bAtend = (xVal(.CellValue(lRow, "IDSESSAO") & "") <> 0)
      If bAtend And Not bMax Then
         Call ExibirStop("Item relacionado ao atendimento e não pode ser removido.")
      Else
         If lRow <> .RowCount And lRow > 0 Then
            If xVal(.CellValue(lRow, "IDDISPARO") & "") <> 0 Then
               If vbYes = ExibirPergunta("Esta ação irá excluir o Disparo Nº" & ValBr(xVal(.CellValue(lRow, "IDDISPARO") & ""), 0), pDefaultYes:=False) Then
                  If xVal(.CellValue(lRow, "ITEM") & "") <> 0 Then
                     If DeleteDisparo(xVal(.CellValue(lRow, "IDDISPARO") & "")) Then
                        Call PopulaGrdDisparos(True)
                     Else
                        Call ExibirAviso("Erro ao excluir disparo.")
                        Exit Function
                     End If
                  End If
               Else
                  Exit Function
               End If
            End If
            
            If Not bAtend Then
               .RemoveRow (lRow)
               Call RenumerarDisparo(lRow, 1)
            End If
            ExcluirDisparo = True
            
            .SetFocus
            If lRow < .RowCount Then
               .SetCurCell lRow, 1
            ElseIf lRow > 1 Then
               .SetCurCell lRow - 1, 1
            End If
         End If
      End If
   End With
End Function
Private Sub mvarMe_Load()
    mvarMe.Icon = mvarSys.MDI.Icon
   
   '* Monta Tela
   Call MontaTela
   
   Screen.MousePointer = vbDefault
End Sub
Private Sub MontaTela()
   Call MontaDTINI
   Call MontaCmbMaquina
   Call MontaCmbManipulo
   Call MontaCmbTratamento
   Call MontaGrdDisparos
   Call PopulaTela
End Sub
Private Function DeleteDisparo(pIDDISPARO As Long) As Boolean
   Dim MyTb As Object 'TB_OMAQDISPAROS
   
   Set MyTb = CriarObjeto("BANCO_3R.TB_OMAQDISPAROS")
   With MyTb
      Set .xDb = mvarSys.xDb
      If .Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDMAQUINA:=mvarIDMAQUINA, Ch_IDDISPARO:=pIDDISPARO) Then
         .FLGDELETE = 1
         DeleteDisparo = .Salvar
      Else
         DeleteDisparo = True
      End If
   End With
End Function
Private Sub PopulaTela()
   Screen.MousePointer = vbHourglass
   If mvarMe.GrdDisparos.ColCount > 0 Then
      Call LimparTela
      Call PopulaGrdDisparos
   End If
   Screen.MousePointer = vbDefault
End Sub
Private Sub PopulaGrdDisparos(Optional bRefresh As Boolean)
   Dim MyRs As Object
   Dim sAux As String
'   Dim MyTb As Object 'TB_OMAQDISPAROS
'   Dim sTag As String
   Dim i    As Integer
   Dim j    As Integer
   Dim Sql  As String
   Dim sTPTRAT As String
   Dim sGr As String
   Dim sDtIni As String
   
   If mvarMe.CmbIDMAQUINA.ListIndex < 0 Then
      Exit Sub
   Else
      If Not bRefresh Then
         bRefresh = (mvarMe.CmbIDMAQUINA.Text <> GetTag(mvarMe.CmbIDMAQUINA, "TEXT"))
      End If
   End If
   
   mvarMe.GrdDisparos.Redraw = False
   
   sTPTRAT = ""
   sAux = CollNUMTRAT("k" & mvarMe.CmbIDMAQUINA.ItemData(mvarMe.CmbIDMAQUINA.ListIndex))
   If bRefresh Then sGr = "-1"
   
   For i = 1 To GetTag(sAux, "GR", "0")
      sGr = GetTag(sAux, "GR" & i, "")
      If InStr(sGr, mvarMe.CmbIDTPTRATAMENTO.ItemData(mvarMe.CmbIDTPTRATAMENTO.ListIndex)) <> 0 Then
         sTPTRAT = sGr
         Exit For
      End If
   Next
   If GetTag(mvarMe.GrdDisparos, "GRUPO", "") = sGr And Not bRefresh Then
      GoTo Saida
   Else
      Call SetTag(mvarMe.GrdDisparos, "GRUPO", sGr)
      mvarMe.GrdDisparos.Clear False
      Call IncluiGrdLinha
      
   End If

   Sql = ""
   Sql = "Select Max(D.DTDISPARO) As [DTDISPARO] "
   Sql = Sql & " From OMAQDISPAROS D"
   Sql = Sql & " Where D.FLGDELETE = 0"
   Sql = Sql & " And D.IDMAQUINA = " & mvarIDMAQUINA
   Sql = Sql & " And D.IDLOJA = " & mvarIDLOJA
   
   If Trim(sTPTRAT) <> "" Then
      Sql = Sql & " And (D.IDTPTRATAMENTO in (" & sTPTRAT & ") Or D.IDTPTRATAMENTO is Null)"
   End If

   If mvarSys.xDb.AbreTabela(Sql) Then
      sDtIni = Format(mvarSys.xDb.RsAux("DTDISPARO"), "dd/mm/yyyy")
      If sDtIni = "" Then sDtIni = "01/01/2010"
      If CDate(Format(mvarMe.CmbDTINI.Value, "dd/mm/yyyy")) <= CDate(sDtIni) Then
         sDtIni = Format(mvarMe.CmbDTINI.Value, "dd/mm/yyyy")
      End If
   End If

   Sql = ""
   'Sql = "Select Row_Number() Over (Order By IDDISPARO, DTDISPARO) [#], *"
   'Sql = Sql & " FROM ("
   Sql = Sql & " Select D.IDMAQUINA, D.IDDISPARO, D.IDDISPARO [IDDISPARO0], 0 [IDCLIENTE]"
   Sql = Sql & ", D.IDTPMANIPULO"
   Sql = Sql & ", Convert(varchar, D.DTDISPARO, 103) [DTDISPARO]"
   Sql = Sql & ", IsNull(D.IDAREA,0) [IDAREA], D.NUMDISPARO, D.QTDDISPARO"
   Sql = Sql & ", D.IDLOJA, D.IDATENDIMENTO, D.IDSESSAO"
   Sql = Sql & ", D.DTDISPARO as DATA, D.IDDISPARO [ITEM], S.IDTPSERVICO, D.IDTPTRATAMENTO"
   Sql = Sql & ", D.IDMANIPULO, D.IDLAMP"
   Sql = Sql & " From OMAQDISPAROS D"
   Sql = Sql & " Left Join OSESSAO S On S.IDLOJA=D.IDLOJA And S.IDATENDIMENTO=D.IDATENDIMENTO And S.IDSESSAO=D.IDSESSAO "
   Sql = Sql & " Where D.IDLOJA =" & mvarIDLOJA
   Sql = Sql & " And D.FLGDELETE = 0"
   Sql = Sql & " And D.IDMAQUINA = " & mvarMe.CmbIDMAQUINA.ItemData(mvarMe.CmbIDMAQUINA.ListIndex) 'mvarIDMAQUINA
   If Not IsNull(mvarMe.CmbDTINI.Value) Then
      Sql = Sql & " And D.DTDISPARO >=" & SqlDate(sDtIni)
   End If
   If Trim(sTPTRAT) <> "" Then
      Sql = Sql & " And (D.IDTPTRATAMENTO in (" & sTPTRAT & ") Or D.IDTPTRATAMENTO is Null)"
   End If
   Sql = Sql & " Union All"
   'Sql = Sql & " Select A.IDMAQUINA, (Select Count(*)+1 From OMAQDISPAROS) [IDDISPARO]"
   Sql = Sql & " Select A.IDMAQUINA, 0 [IDDISPARO], 0 [IDDISPARO0], A.IDCLIENTE"
   Sql = Sql & ", M.IDTPMANIPULO"
   Sql = Sql & ", Convert(varchar, A.DTATEND, 103) [DTDISPARO]"
   Sql = Sql & ", IsNull(S.IDAREA,0) [IDAREA], Null [NUMDISPARO], Null [QTDDISPARO]"
   Sql = Sql & ", S.IDLOJA, S.IDATENDIMENTO, S.IDSESSAO"
   Sql = Sql & ", A.DTATEND as DATA"
   Sql = Sql & ", (Select Max(IDDISPARO)+1"
   Sql = Sql & "   From OMAQDISPAROS"
   Sql = Sql & "   Where IDLOJA = " & mvarIDLOJA
   Sql = Sql & "   And IDMAQUINA = " & mvarMe.CmbIDMAQUINA.ItemData(mvarMe.CmbIDMAQUINA.ListIndex)
   'Sql = Sql & "   AND IDSESSAO IS NULL "
   Sql = Sql & "   AND FLGDELETE = 0) [ITEM]"
   Sql = Sql & ", S.IDTPSERVICO, S.IDTPTRATAMENTO"
   Sql = Sql & ", 1 AS [IDMANIPULO], 1 [IDLAMP]"
   Sql = Sql & " From OSESSAO S "
   Sql = Sql & " Join OATENDIMENTO A      On S.IDLOJA=A.IDLOJA And S.IDATENDIMENTO=A.IDATENDIMENTO"
   Sql = Sql & " Left Join OMAQDISPAROS D On S.IDLOJA=D.IDLOJA And S.IDATENDIMENTO=D.IDATENDIMENTO And S.IDSESSAO=D.IDSESSAO And D.FLGDELETE = 0"
   Sql = Sql & " Left Join OMANIPULO M    On S.IDLOJA=M.IDLOJA And S.IDMANIPULO=M.IDMANIPULO"
   Sql = Sql & " Where A.IDLOJA = " & mvarIDLOJA
   Sql = Sql & " And A.IDMAQUINA = " & mvarMe.CmbIDMAQUINA.ItemData(mvarMe.CmbIDMAQUINA.ListIndex) 'mvarIDMAQUINA
   Sql = Sql & " And D.IDSESSAO Is Null"
   If Trim(sTPTRAT) <> "" Then
      Sql = Sql & " And (S.IDTPTRATAMENTO in (" & sTPTRAT & ") Or S.IDTPTRATAMENTO is Null)"
   End If
   Sql = Sql & " Order By ITEM, DATA"
   'Sql = Sql & " ) Temp"
   
   'D.IDDISPARO , D.DTDISPARO, D.IDAREA, D.NUMDISPARO, D.QTDDISPARO, D.IDLOJA, D.IDATENDIMENTO, D.IDSESSAO
   If mvarSys.xDb.AbreTabela(Sql, MyRs) Then
      With mvarMe.GrdDisparos
         .Clear
         .FillFromRS MyRs
         Call IncluiGrdLinha(False)
         '****************
         '* Pintar clientes
         sAux = ""
         For i = 1 To .RowCount - 1
            For j = 1 To .ColCount
               If xVal(.CellValue(i, "IDCLIENTE") & "") = 0 Then
                  .CellBackColor(i, j) = vbWhite
               ElseIf sAux = .CellValue(i, "IDCLIENTE") Then
                  .CellBackColor(i, j) = .CellBackColor(i - IIf(i = 1, 0, 1), j)
               Else
                  If i = 1 Then
                     .CellBackColor(i, j) = vbWhite
                  Else
                     If .CellBackColor(i - 1, j) = &HE0E0E0 Then
                        .CellBackColor(i, j) = vbWhite
                     Else
                        .CellBackColor(i, j) = &HE0E0E0   ' &H8000000F&  &H00E0E0E0&
                     End If
                  End If
               End If
            Next
            sAux = .CellValue(i, "IDCLIENTE")
         Next
         
         '************
         '* Exibir últ. Lançamento
         DoEvents
         For i = 1 To .RowCount
            If xVal(.CellValue(i, "IDDISPARO") & "") = 0 Then
               Exit For
            End If
         Next
         Call .SetCurCell(i, "QTDDISPARO")
         i = IIf(i < 25, 1, i - 1)
         .Redraw = True
         .RowVisible(i) = True
         Call .EnsureVisibleRow(i, igEnsVisRowTop)
         .Redraw = True
         
      End With
   End If
Saida:
   mvarMe.GrdDisparos.Redraw = True
End Sub
Private Sub LimparTela()
   With mvarMe
'      If .CmbIDMAQUINA.ListCount > 1 Then .CmbIDMAQUINA.ListIndex = -1
'      .GrdDisparos.Clear False
'      Call IncluiGrdLinha
   End With
End Sub

Private Sub MontaDTINI()
   Dim Sql As String
   mvarMe.CmbDTINI.Visible = True
   mvarMe.CmbDTINI.Enabled = True
   
   Sql = "Select DTOPERACAO "
   Sql = Sql & " From OLOJA"
   Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
   If mvarSys.xDb.AbreTabela(Sql) Then
      mvarMe.CmbDTINI.Value = mvarSys.xDb.RsAux("DTOPERACAO")
   Else
      mvarMe.CmbDTINI.Value = CDate("01/01/" & Year(mvarSys.xDb.Sysdate()))
   End If
   
   Sql = "Select Max(DTMANUT) [DATA]"
   Sql = Sql & " From OMAQMANUT"
   Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
   Sql = Sql & " And IDTPMANUT=2"
   If mvarSys.xDb.AbreTabela(Sql) Then
      If IsNull(mvarSys.xDb.RsAux("DATA")) Then
         
         Sql = "Select Max(DTDISPARO) [DATA]"
         Sql = Sql & " From OMAQDISPAROS"
         Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
         Sql = Sql & " And FLGDELETE=0"
         Sql = Sql & " And QTDDISPARO=0"
         If mvarSys.xDb.AbreTabela(Sql) Then
            If IsNull(mvarSys.xDb.RsAux("DATA")) Then
               Sql = "Select Min(DTATEND) [DATA]"
               Sql = Sql & " From OATENDIMENTO"
               Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
               If mvarSys.xDb.AbreTabela(Sql) Then
                  If Not IsNull(mvarSys.xDb.RsAux("DATA")) Then
                     mvarMe.CmbDTINI.Value = mvarSys.xDb.RsAux("DATA")
                  End If
               End If
            Else
               mvarMe.CmbDTINI.Value = mvarSys.xDb.RsAux("DATA")
            End If
         End If
      Else
         mvarMe.CmbDTINI.Value = mvarSys.xDb.RsAux("DATA")
      End If
   End If
End Sub

Private Sub MontaCmbMaquina()
   Dim TbCmb   As Object
   Dim oRs     As Object
   Dim oCombo  As XtremeSuiteControls.ComboBox
   Dim Sql     As String
   Dim sTag    As String
   Dim sAux    As String
   Dim i       As Integer
      
   Set oCombo = mvarMe.CmbIDMAQUINA
   
   Set TbCmb = CriarObjeto("BANCO_3R.TB_OMAQUINA")
   With TbCmb
      Set .xDb = mvarSys.xDb
      oCombo.Clear
      Set CollNUMTRAT = New Collection
                  
      Sql = "Select Distinct M.*, T.DSCMAQ, T.TPNUM, T.TPMANIPULO, T.NUMTRATAMENTO "
      Sql = Sql & " From OMAQUINA M "
      Sql = Sql & " Join OSALA_MAQUINA S On M.IDLOJA=S.IDLOJA And M.IDMAQUINA=S.IDMAQUINA"
      Sql = Sql & " Join OTPMAQ T On M.IDLOJA=T.IDLOJA And M.IDTPMAQ=T.IDTPMAQ"
      Sql = Sql & " Where M.SITMAQUINA=1"
      Sql = Sql & " And S.IDLOJA=" & mvarIDLOJA
      Sql = Sql & " And S.DTINICIO <=" & SqlDate(Format(Now(), "dd/mm/yyyy"))
      Sql = Sql & " And (S.DTFIM is Null Or S.DTFIM >=" & SqlDate(Format(Now(), "dd/mm/yyyy")) & ")"
      sTag = ""
      If mvarSys.xDb.AbreTabela(Sql, oRs) Then
         While Not oRs.EOF
            oCombo.AddItem oRs("CODMAQUINA")
            oCombo.ItemData(oCombo.NewIndex) = oRs("IDMAQUINA")
            sTag = SetTag(sTag, oRs("IDMAQUINA") & "_TPNUM", oRs("TPNUM"))
            sTag = SetTag(sTag, oRs("IDMAQUINA") & "_TPMANIPULO", oRs("TPMANIPULO"))
            
            CollNUMTRAT.Add oRs("NUMTRATAMENTO") & "", "k" & oRs("IDMAQUINA")
            oRs.MoveNext
         Wend
         oCombo.Tag = sTag
      End If
   End With
   Set TbCmb = Nothing
   With oCombo
      .ListIndex = IIf(.ListCount = 1, 0, -1)
      .Enabled = (.ListCount > 1)
   End With
End Sub
Private Sub MontaCmbManipulo()
   Dim TbCmb  As Object
   Dim oCombo As XtremeSuiteControls.ComboBox
   Dim Sql    As String
      
   Set oCombo = mvarMe.CmbIDMANIPULO
   
   Set TbCmb = CriarObjeto("BANCO_3R.TB_OMANIPULO")
   With TbCmb
      Set .xDb = mvarSys.xDb
      oCombo.Clear
      Set CollTPMANIPULO = New Collection
      
      
'      Sql = "Select M.* "
'      Sql = Sql & " From OMAQUINA M Join OSALA_MAQUINA S On M.IDMAQUINA=S.IDMAQUINA"
'      Sql = Sql & " Where S.IDLOJA=" & mvarIDLOJA
'      Sql = Sql & " And S.DTINICIO <='" & Format(Now(), "dd/mm/yyyy") & "'"
'      Sql = Sql & " And (S.DTFIM is Null Or S.DTFIM >='" & Format(Now(), "dd/mm/yyyy") & "')"
      Sql = "Select * From OMANIPULO Where IDLOJA=" & mvarIDLOJA
      If mvarSys.xDb.AbreTabela(Sql) Then
         While Not mvarSys.xDb.RsAux.EOF
            .Popula mvarSys.xDb.RsAux
            
            oCombo.AddItem .CODMANIPULO
            oCombo.ItemData(oCombo.NewIndex) = .IDMANIPULO
            CollTPMANIPULO.Add .IDTPMANIPULO, "k" & .IDMANIPULO
            mvarSys.xDb.RsAux.MoveNext
         Wend
      End If
   End With
   Set TbCmb = Nothing
   With oCombo
      .ListIndex = IIf(.ListCount >= 1, 0, -1)
      .Enabled = (.ListCount > 1)
   End With
End Sub
Private Sub MontaCmbTratamento()
   Dim TbCmb  As Object
   Dim oCombo As XtremeSuiteControls.ComboBox
   Dim Sql    As String
      
   Set oCombo = mvarMe.CmbIDTPTRATAMENTO
   
   Set TbCmb = CriarObjeto("BANCO_3R.TB_OTPTRATAMENTO")
   With TbCmb
      Set .xDb = mvarSys.xDb
      oCombo.Clear
      
      Sql = "Select *"
      Sql = Sql & " From OTPTRATAMENTO"
      Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
      Sql = Sql & " And FLGDISPARO=1"
      If mvarSys.xDb.AbreTabela(Sql) Then
         While Not mvarSys.xDb.RsAux.EOF
            .Popula mvarSys.xDb.RsAux
            
            oCombo.AddItem .DSCTRATAMENTO
            oCombo.ItemData(oCombo.NewIndex) = .IDTPTRATAMENTO
            mvarSys.xDb.RsAux.MoveNext
         Wend
      End If
   End With
   Set TbCmb = Nothing
   With oCombo
      .ListIndex = IIf(.ListCount >= 1, 0, -1)
      .Enabled = (.ListCount > 1)
   End With
End Sub
Private Sub MontaGrdDisparos()
   Dim sAux       As String
   Dim i          As Integer
     
   On Error Resume Next
   
   With mvarMe.GrdDisparos
      .ShowControlsInAllCells = True
      .Header.DragCols = False
      .SilentValidation = True
      '.Header.ImageList = mvarMe.IlsIcons
      '.ImageList = mvarMe.IlsIcons
      .Header.Font.Size = 9
      .Header.Font.Name = mvarMe.Font.Name
      .Font.Size = .Header.Font.Size
      .Font.Name = .Header.Font.Name
      .DefaultRowHeight = 20
            
      With .Combos
         Dim TbArea     As Object 'TB_OAREA
         Set TbArea = CriarObjeto("BANCO_3R.TB_OAREA")
         Set TbArea.xDb = mvarSys.xDb
         With .Add("IDAREA")
            .AddItem "_Outros", 0
            If TbArea.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_ORDERBY:="DSCAREA") Then
               While Not TbArea.RS.EOF
                  TbArea.Popula
                  .AddItem TbArea.DSCAREA, TbArea.IDAREA
                  TbArea.RS.MoveNext
               Wend
            End If
            Set TbArea = Nothing
            '.AutoAdjustWidth
            .Width = 120
         End With
         
         Dim TbTrat  As Object 'TB_OTRATAMENTO
         Set TbTrat = CriarObjeto("BANCO_3R.TB_OTPTRATAMENTO")
         Set TbTrat.xDb = mvarSys.xDb
         With .Add("IDTPTRATAMENTO")
            If TbTrat.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_ORDERBY:="IDTPTRATAMENTO") Then
               While Not TbTrat.RS.EOF
                  TbTrat.Popula
                  .AddItem TbTrat.DSCTRATAMENTO, TbTrat.IDTPTRATAMENTO
                  TbTrat.RS.MoveNext
               Wend
            End If
            Set TbTrat = Nothing
            '.AutoAdjustWidth
            .Width = 120
         End With
         
         Dim TbManip  As Object 'TB_OMANIPULO
         Set TbManip = CriarObjeto("BANCO_3R.TB_OMANIPULO")
         Set TbManip.xDb = mvarSys.xDb
         With .Add("IDMANIPULO")
            If TbManip.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_ORDERBY:="IDMANIPULO") Then
               While Not TbManip.RS.EOF
                  TbManip.Popula
                  .AddItem TbManip.CODMANIPULO, TbManip.IDMANIPULO
                  TbManip.RS.MoveNext
               Wend
            End If
            Set TbManip = Nothing
            '.AutoAdjustWidth
            .Width = 80
         End With
      End With
            
'      With .AddCol(sKey:="#", sHeader:="#", lWidth:=30, bVisible:=True)
'         .eType = igCellText
'         .eTextFlags = igTextLeft
'         .sCtrlKey = "#"
'         .bSelected = False
'      End With

      With .AddCol(sKey:="IDMAQUINA", sHeader:="#", lWidth:=35, bVisible:=False)
         .eType = igCellText
         .eTextFlags = igTextLeft
         .sCtrlKey = "IDMAQUINA"
         .bSelected = False
      End With
      With .AddCol(sKey:="IDDISPARO", sHeader:="#", lWidth:=35, bVisible:=True)
         .eType = igCellText
         .eTextFlags = igTextLeft
         .sCtrlKey = "IDDISPARO"
         .sFmtString = "#,###"
         .bSelected = False
      End With
      With .AddCol(sKey:="IDDISPARO0", sHeader:="IDDISPARO0", lWidth:=35, bVisible:=False)
         .eType = igCellText
         .eTextFlags = igTextLeft
         .sCtrlKey = "IDDISPARO0"
         .sFmtString = "#,###"
         .bSelected = False
      End With
      
      With .AddCol(sKey:="IDCLIENTE", sHeader:="#", lWidth:=30, bVisible:=False)
         .eType = igCellText
         .eTextFlags = igTextLeft
         .sCtrlKey = "IDCLIENTE"
         .sFmtString = "#,###"
         .bSelected = False
      End With
      With .AddCol(sKey:="IDTPMANIPULO", sHeader:="IDTPMANIPULO", lWidth:=35, bVisible:=False)
         .eType = igCellText
         .eTextFlags = igTextLeft
         .sCtrlKey = "IDTPMANIPULO"
         .bSelected = False
      End With
      
      With .AddCol(sKey:="DTDISPARO", sHeader:="Data", lWidth:=85, bVisible:=True)
         .eType = igCellText
         .eTextFlags = igTextCenter
         .sCtrlKey = "DTDISPARO"
      End With
      With .AddCol(sKey:="IDTPTRATAMENTO", sHeader:="Tratamento", lWidth:=120, bVisible:=True)
         .eType = igCellText
         .eTypeFlags = igTextCenter
         .sCtrlKey = "IDTPTRATAMENTO"
      End With
      With .AddCol(sKey:="IDAREA", sHeader:="Área", lWidth:=120, bVisible:=True)
         .eType = igCellCombo
         .eTypeFlags = igComboBtnFlat
         .sCtrlKey = "IDAREA"
      End With
      With .AddCol(sKey:="IDMANIPULO", sHeader:="Manípulo", lWidth:=80, bVisible:=True)
         .eType = igCellCombo
         .eTypeFlags = igComboBtnFlat
         .sCtrlKey = "IDMANIPULO"
         .bSelected = True
      End With
      With .AddCol(sKey:="IDLAMP", sHeader:="Lâmpada", lWidth:=35, bVisible:=False)
         .eType = igCellText
         .eTextFlags = igTextLeft
         .sCtrlKey = "IDLAMP"
         .bSelected = False
      End With
      With .AddCol(sKey:="NUMDISPARO", sHeader:="Nº Regsitro", lWidth:=85, bVisible:=True)
         .eType = igCellText
         .eTextFlags = igTextRight
         .sCtrlKey = "NUMDISPARO"
         .sFmtString = "#,###"
      End With
      With .AddCol(sKey:="QTDDISPARO", sHeader:="Qtd. Disparos", lWidth:=85, bVisible:=True)
         .eType = igCellText
         .eTextFlags = igTextRight
         .sCtrlKey = "QTDDISPARO"
         .sFmtString = "0"
      End With
      With .AddCol(sKey:="IDLOJA", sHeader:="IDLOJA", lWidth:=30, bVisible:=False)
         .eType = igCellText
         .eTextFlags = igTextCenter
         .sCtrlKey = "IDLOJA"
         .eTextFlags = igTextCenter
      End With
      With .AddCol(sKey:="IDATENDIMENTO", sHeader:="IDATENDIMENTO", lWidth:=30, bVisible:=False)
         .eType = igCellText
         .eTypeFlags = igTextCenter
         .sCtrlKey = "IDATENDIMENTO"
      End With
      With .AddCol(sKey:="IDSESSAO", sHeader:="IDSESSAO", lWidth:=30, bVisible:=False)
         .eType = igCellText
         .eTypeFlags = igTextCenter
         .sCtrlKey = "IDSESSAO"
      End With
      With .AddCol(sKey:="DATA", sHeader:="DATA", lWidth:=30, bVisible:=False)
         .eType = igCellText
         .eTypeFlags = igTextCenter
         .sCtrlKey = "DATA"
      End With
      With .AddCol(sKey:="ITEM", sHeader:="ITEM", lWidth:=30, bVisible:=False)
         .eType = igCellText
         .eTypeFlags = igTextCenter
         .sCtrlKey = "ITEM"
      End With
      With .AddCol(sKey:="IDTPSERVICO", sHeader:="IDTPSERVICO", lWidth:=30, bVisible:=False)
         .eType = igCellText
         .eTypeFlags = igTextCenter
         .sCtrlKey = "IDTPSERVICO"
      End With
                       
      .ColHeaderTextFlags("IDDISPARO") = igTextCenter
      .ColHeaderTextFlags("DTDISPARO") = igTextCenter
      .ColHeaderTextFlags("NUMDISPARO") = igTextRight
      .ColHeaderTextFlags("QTDDISPARO") = igTextRight
                       
      For i = 1 To .ColCount
         .ColTag(i) = ""
      Next
                                              
      '.ColTag("#") = SetTag(.ColTag("#"), "EDIT", False)
      .ColTag("IDDISPARO") = SetTag(.ColTag("IDDISPARO"), "EDIT", False)
      .ColTag("IDLOJA") = SetTag(.ColTag("IDLOJA"), "EDIT", False)
      .ColTag("IDATENDIMENTO") = SetTag(.ColTag("IDATENDIMENTO"), "EDIT", False)
      .ColTag("IDSESSAO") = SetTag(.ColTag("IDSESSAO"), "EDIT", False)
      .ColTag("IDTPMANIPULO") = SetTag(.ColTag("IDTPMANIPULO"), "EDIT", False)
      
            
      '.ColTag("DTDISPARO") = SetTag(.ColTag("DTDISPARO"), "EDITOPT", igText)
      .ColTag("NUMDISPARO") = SetTag(.ColTag("NUMDISPARO"), "EDITOPT", igTextEditNumberOnly)
      .ColTag("QTDDISPARO") = SetTag(.ColTag("QTDDISPARO"), "EDITOPT", igTextEditNumberOnly)
                                  
      IncluiGrdLinha
   End With
End Sub
Private Sub IncluiGrdLinha(Optional bLinNova As Boolean = True, Optional vRowBefore)
   Dim lRow As Long
   Dim i As Integer
   Dim nColDefault As Integer
   Dim bRedraw As Boolean
   Dim bAdd As Boolean
   
'   On Error Resume Next
   nColDefault = 4
   With mvarMe.GrdDisparos
      If .ColCount <= 0 Then Exit Sub
      bRedraw = .Redraw
      .Redraw = False
      lRow = .RowCount
      If lRow > 1 Then .CellForeColor(lRow - 1, "IDDISPARO") = vbGrayed
      
      If bLinNova Then
         
         If Not IsMissing(vRowBefore) Then
            .AddRow vRowBefore:=vRowBefore
            .CellValue(vRowBefore, "IDTPTRATAMENTO") = 1
            .CellValue(vRowBefore, "IDAREA") = 0
            .CellValue(vRowBefore, "IDMANIPULO") = 1
            .CellValue(vRowBefore, "IDTPMANIPULO") = 1
         Else
            If lRow > 0 Then
               .AddRow vRowBefore:=lRow
               For i = 0 To .Combos("IDTPTRATAMENTO").ListCount
                  If .CellText(lRow, "IDTPTRATAMENTO") = .Combos("IDTPTRATAMENTO").ItemText(i) Then
                     .CellValue(lRow, "IDTPTRATAMENTO") = .Combos("IDTPTRATAMENTO").ItemValue(i)
                     Exit For
                  End If
               Next
               bAdd = True
            End If
         End If
         lRow = .RowCount
         For i = 1 To .ColCount
            If GetTag(.ColTag(i), "EDIT", True) = False Then
               If lRow > 0 Then
                  .CellForeColor(lRow, i) = vbGrayText
               End If
            End If
         Next
      End If
      
      If Not bAdd Then
         .AddRow
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
      .RowMode = (.RowCount = 1)
      If Not IsMissing(vRowBefore) Then
         lRow = vRowBefore
      End If
      If lRow > 0 Then
         If .CellValue(lRow, "IDAREA") = "" Then
            nColDefault = .ColIndex("IDAREA")  '.SetCurCell .RowCount, 1
         End If
         If .CellValue(lRow, "DTDISPARO") = "" Then
            nColDefault = .ColIndex("DTDISPARO")  '.SetCurCell .RowCount, 1
         End If
         .SetCurCell lRow, nColDefault
      End If
      If .Visible Then
         .SetFocus
      End If
      .Redraw = bRedraw
   End With
End Sub
Private Sub RefreshGrid()
   Dim lRow  As Integer
   Dim i As Integer
   Dim bRedraw As Boolean
   
   If mvarMe.CmbDTINI.Visible And mvarMe.GrdDisparos.Visible Then
      mvarMe.Refresh
      DoEvents
      Screen.MousePointer = vbHourglass
      With mvarMe.GrdDisparos
         lRow = .CurRow
         i = .CurCol
         bRedraw = .Redraw
         .Redraw = False
         Call PopulaTela
         If .RowCount >= lRow And .ColCount >= i Then
            .SetCurCell lRow, i
         End If
         .Redraw = bRedraw
      End With
      Screen.MousePointer = vbDefault
   End If
End Sub

Private Sub mvarMe_Resize()
   mvarMe.Width = 8235 + 1750
   mvarMe.GrdDisparos.Width = mvarMe.Width - 2 * (mvarMe.GrdDisparos.Left) - mvarMe.CmdDown.Width
   mvarMe.CmdDown.Left = mvarMe.GrdDisparos.Left + mvarMe.GrdDisparos.Width + 30
   mvarMe.CmdUp.Left = mvarMe.CmdDown.Left
End Sub
