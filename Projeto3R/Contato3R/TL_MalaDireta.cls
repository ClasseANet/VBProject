VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_MalaDireta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe     As FrmMalaDireta
Attribute mvarMe.VB_VarHelpID = -1
Private mvarSys As Object
Private mvarRs As Object
Const idxFilterAll = "(All)"
Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Public Property Set Rs(ByVal vData As Object)
   Set mvarRs = vData
End Property
Public Property Get Rs() As Object
    Set Rs = mvarRs
End Property
Public Sub Show()
   mvarMe.Show vbModal
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmMalaDireta
End Sub
Private Sub Class_Terminate()
   On Error Resume Next
   Unload mvarMe
   Set mvarMe = Nothing
End Sub
Private Sub mvarMe_Activate()
'   PosicionaForm
'   Call mvarMe_Resize
End Sub

Private Sub mvarMe_ChkAtivo(Index As Integer)
    Call mvarMe_CmdCarregarClick
End Sub
Private Sub mvarMe_CmdCarregarClick()
   Call CmdCarregar
End Sub
Private Sub CmdCarregar(Optional pSql As String)
   Dim Sql  As String
   Dim i    As Long
    
   Screen.MousePointer = vbHourglass
   'Call mvarMe_Resize
   Call LimparGrid
         
   Sql = "Select '' [#], C.IDCLIENTE [Id], C.NOME [Nome], C.EMAIL [Email], C.CIDADE [Cidade], C.BAIRRO [Bairro]" & vbNewLine
   Sql = Sql & ", C.FOTOTIPO [Fototipo], C.ATIVO [Ativo]" & vbNewLine
   Sql = Sql & ", E1.StartDateTime [Últ.Sessão]" & vbNewLine
   Sql = Sql & ", E2.StartDateTime [Prox.Sessão]" & vbNewLine
   Sql = Sql & " From OCLIENTE C " & vbNewLine
         
   Sql = Sql & " Left Join OEVENTOAGENDA E1 On E1.IDLOJA=C.IDLOJA " & vbNewLine
   Sql = Sql & "    And E1.IDCLIENTE=C.IDCLIENTE " & vbNewLine
   Sql = Sql & "    And E1.FLGCANCELADO=0 " & vbNewLine
   Sql = Sql & "    And E1.StartDateTime<GetDate() " & vbNewLine
   Sql = Sql & "    And E1.StartDateTime=(Select Max(E.StartDateTime) From OEVENTOAGENDA E Where E.IDLOJA=C.IDLOJA AND E.IDCLIENTE=C.IDCLIENTE And E.FLGCANCELADO=0 And E.StartDateTime<GetDate())" & vbNewLine
   
   Sql = Sql & " Left Join OEVENTOAGENDA E2 On E2.IDLOJA=C.IDLOJA " & vbNewLine
   Sql = Sql & "    And E2.IDCLIENTE=C.IDCLIENTE " & vbNewLine
   Sql = Sql & "    And E2.FLGCANCELADO=0 " & vbNewLine
   Sql = Sql & "    And E2.StartDateTime>=GetDate() " & vbNewLine
   Sql = Sql & "    And E2.StartDateTime=(Select Min(E.StartDateTime) From OEVENTOAGENDA E Where E.IDLOJA=C.IDLOJA AND E.IDCLIENTE=C.IDCLIENTE And E.FLGCANCELADO=0 And E.StartDateTime>=GetDate())" & vbNewLine

   
   Sql = Sql & " Where C.IDLOJA=" & mvarSys.Propriedades("IDLOJA") & vbNewLine
   Sql = Sql & " And (C.ATIVO is null"
   If mvarMe.ChkAtivo.Value Then
      Sql = Sql & " Or C.ATIVO = 1"
   End If
   If mvarMe.ChkInativo.Value Then
      Sql = Sql & " Or C.ATIVO = 0"
   End If
   If mvarMe.ChkEmEspera.Value Then
      Sql = Sql & " Or C.ATIVO = 2"
   End If
   Sql = Sql & " )" & vbNewLine
   Sql = Sql & " And (C.FLGMARKETING=1 or C.FLGMARKETING is null)" & vbNewLine
   Sql = Sql & " And C.EMAIL<>''" & vbNewLine
   Sql = Sql & " Order By C.NOME"
   
   If pSql = "" Then pSql = Sql
   
   With mvarMe.GrdContato
      If mvarSys.xDb.Abretabela(pSql, mvarRs) Then
         Call FillRCFromRS(mvarRs, mvarMe.GrdContato)
      End If
      If pSql <> "" Then
         For i = 0 To .Records.Count - 1
            .Records(i).Item(0).HasCheckbox = True
            .Records(i).Item(0).Checked = True
         Next
         .Populate
         .Redraw
         Call PopulaStGrd
         For i = .Columns.Count - 1 To 0 Step -1
            .Columns(i).BestFit
         Next
         If .Columns(0).Width < 50 Then .Columns(0).Width = 50
      Else

         .Columns(0).Width = 30
         .Columns(0).Alignment = xtpAlignmentCenter
         .Columns(1).Visible = False
         For i = 0 To .Records.Count - 1
            .Records(i).Item(0).HasCheckbox = True
            .Records(i).Item(0).Checked = True
         Next
         .Populate
         .Redraw
         Call PopulaStGrd
         For i = .Columns.Count - 1 To 0 Step -1
            .Columns(i).BestFit
         Next
         .Columns(8).Width = 80
         .Columns(9).Width = 80
         
         .Columns(7).BestFit
         .Columns(6).BestFit
         .Columns(5).BestFit
         
         .Columns(5).Width = IIf(.Columns(5).Width < 100, 100, .Columns(3).Width)
         .Columns(4).BestFit
         .Columns(4).Width = IIf(.Columns(4).Width < 100, 100, .Columns(3).Width)
         .Columns(3).BestFit
         .Columns(3).Width = IIf(.Columns(3).Width < 200, 200, .Columns(3).Width)
         .Columns(2).Width = IIf(.Columns(2).Width < 240, 240, .Columns(2).Width)
         .Columns(8).Width = 80
         .Columns(9).Width = 80
         .Columns(0).Width = 50
      End If
      For i = 0 To .Columns.Count - 1
         .Columns(0).EditOptions.AddComboButton
         .Columns(0).EditOptions.GetInplaceButton(0).InsideCellButton = True
      Next
            
      .Populate
      .Redraw
   End With
   mvarMe.CmdEnviar.Enabled = (mvarMe.TxtHtml.Text <> "") And (mvarMe.GrdContato.Records.Count > 0)
   
   Screen.MousePointer = vbDefault
End Sub
Private Sub PopulaStGrd(Optional pSel As Long, Optional pItemSel As Long, Optional pEnviado As Long, Optional pErro As Long)
   Dim nTotal As Long
   Dim sAux As String
   Dim i As Long
   
   If IsEmpty(pSel) Or pSel = 0 Then
      pSel = 0
      If IsEmpty(pItemSel) Or pItemSel = 0 Then
         For i = 0 To mvarMe.GrdContato.Records.Count - 1
            If mvarMe.GrdContato.Records(i).Item(0).Checked Then
               pSel = pSel + 1
            End If
         Next
      Else
         pSel = pItemSel + GetTag(mvarMe.GrdContato, "PSEL", 0)
      End If
   End If
   
   nTotal = mvarMe.GrdContato.Records.Count
   sAux = ""
   sAux = sAux & "Total: ( " & nTotal & " Itens )"
   sAux = sAux & Space(10)
   sAux = sAux & "Selecionados: ( " & pSel & " Itens )"
   
   If Not (IsEmpty(pEnviado) Or pEnviado = 0) Then sAux = sAux & "   Enviados: ( " & pEnviado & " Itens )"
   If Not (IsEmpty(pErro) Or pErro = 0) Then sAux = sAux & "   Erros: ( " & pErro & " Itens )"
   
   Call SetTag(mvarMe.GrdContato, "PSEL", pSel)
   
   mvarMe.LblStGrd.Caption = sAux
End Sub
Private Sub LimparGrid()
   Dim Sql As String
   
   Sql = "SELECT '' [#], '' [Id], '' [Nome], '' [Email], '' [Cidade], '' [Bairro]"
   Sql = Sql & ", '' [Fototipo], '' [Ativo], '' [Últ.Sessão],''[Próx.Sessão]"
   Call mvarSys.xDb.Abretabela(Sql, mvarRs)
   Call FillRCFromRS(mvarRs, mvarMe.GrdContato)
   mvarMe.GrdContato.Records.DeleteAll
   mvarMe.CmdEnviar.Enabled = (mvarMe.TxtHtml.Text <> "") And (mvarMe.GrdContato.Records.Count > 0)
   
End Sub

Private Sub mvarMe_CmdEdit()
   Dim oEdtHtml As Object
   On Error Resume Next
   
   Set oEdtHtml = CriarObjeto("CAHtm.EDITOR")
   If Not oEdtHtml Is Nothing Then
      Set oEdtHtml.Sys = mvarSys
      oEdtHtml.File = mvarMe.TxtHtml.Text
      Call oEdtHtml.Show
      If oEdtHtml.File <> "" Then
         Call mvarMe_CmdOpen(oEdtHtml.File)
      End If
   End If
End Sub

Private Sub mvarMe_CmdEnviar()
   Dim PrgBar   As ProgressBar
   Dim bResult  As Boolean
   Dim xMail    As Object
   Dim sFile    As String
   Dim sCaption As String
   Dim sMessage As String
   Dim sTitulo  As String
   Dim sAtach   As String
   Dim sMail    As String
   Dim sAux     As String
   Dim sPath    As String
   Dim sNome    As String
   Dim i        As Integer
   Dim j        As Integer
   Dim k        As Integer
   Dim iSend    As Integer
   Dim iToSend  As Integer
   Dim nTotal   As Integer
   Dim nErro    As Long
   Dim nEnvio   As Long
   Dim sFileRes As String
   Dim dData    As Date
   Dim sMsg     As String
   Dim sLbl     As String
   Dim nGrupo   As Integer
   Dim cGrpNome As Collection
   Dim cGrpMail As Collection
   Dim cGrpKey  As Collection
   
'mvarMe.Timer1.Interval = 30000
'mvarMe.Timer1.Enabled = True
'Exit Sub
   Screen.MousePointer = vbHourglass
   If Trim(mvarMe.TxtTitulo.Text) = "" Then
      Call ExibirInformacao("Título inválido", "Mala Direta")
      GoTo Saida
   End If
   If Trim(mvarMe.TxtHtml.Text) = "" Or Not ExisteArquivo(mvarMe.TxtHtml.Text) Then
      Call ExibirInformacao("Arquivo inválido", "Mala Direta")
      GoTo Saida
   End If
   If mvarRs Is Nothing Then
      Call ExibirInformacao("Lista de clientes inválida", "Mala Direta")
      GoTo Saida
   End If
   If mvarRs.RecordCount <= 0 Then
      Call ExibirInformacao("Lista de clientes inválida", "Mala Direta")
      GoTo Saida
   End If
   
   sCaption = mvarSys.MDI.Caption

   sFile = mvarMe.TxtHtml.Text
   sTitulo = mvarMe.TxtTitulo.Text
   sMessage = ReadTextFile(sFile)
   
   nErro = 0
   nEnvio = 0
   dData = Now()
   Call GetNameFromPath(mvarMe.TxtHtml.Text, sFileRes)
   sFileRes = sFileRes + "Envio_" + Format(dData, "yymmddhhnnss") + ".txt"
   Call WriteIniFile(sFileRes, "Global", "Data", CStr(dData))
   Call WriteIniFile(sFileRes, "Global", "Arquivo", sFile)
   Call WriteIniFile(sFileRes, "Global", "Titulo", sTitulo)
   Call WriteIniFile(sFileRes, "Global", "Remetente", mvarSys.GetParam("FromDisplayName") + "( " + mvarSys.GetParam("MailUID") + " )")
   
   Set xMail = CriarObjeto("CAMail.SendMail")
   With xMail
      .UseAuthentication = (mvarSys.GetParam("UseAuthentication") = xtpChecked)
      .UsePopAuthentication = (mvarSys.GetParam("UsePopAuthentication") = xtpChecked)
      
      .POP3Host = mvarSys.GetParam("POP3Host") ' "pop3.bol.com.br"
      .SMTPHost = mvarSys.GetParam("SMTPHost") ' "smtps.bol.com.br"
      .SMTPPort = mvarSys.GetParam("SMTPPort") ' 587
      .Username = mvarSys.GetParam("MailUID")  ' "diogenes72@bol.com.br"
      .Password = Decrypt2(mvarSys.GetParam("MailPWD"))
      
      .FromDisplayName = mvarSys.GetParam("FromDisplayName") 'FromDisplayName ' "Diogenes"
                  
      .Receipt = True
      .SMTPHostValidation = 0 'VALIDATE_HOST_NONE
                        
      .Subject = sTitulo
      .Message = sMessage
      
      .RecipientDisplayName = "Cliente"           ' "DiSantos"
      .From = .Username
      .AsHTML = True
      
      sFile = GetNameFromPath(sFile, sPath)
      sAux = Dir(sPath, vbArchive)
      sAtach = ""
      While sAux <> ""
         If UCase(sAux) <> UCase(sFile) And Right(sAux, 4) <> ".htm" And Right(sAux, 5) <> ".html" Then
            If InStr(UCase(sMessage), UCase(sAux)) <> 0 Then
               sAtach = sAtach & IIf(sAtach = "", "", ";") & sPath & sAux
            End If
         End If
         sAux = Dir()
      Wend
      If Trim(sAtach) <> "" Then
         .Attachment = sAtach
      End If
      .Subject = sTitulo
      
      
      'If mvarSys.xDb.AbreTabela(Sql, oRs) Then
      'If Not mvarRs Is Nothing Then
          'mvarRs.MoveFirst
      If mvarMe.GrdContato.Records.Count > 0 Then
         Set PrgBar = mvarSys.MDI.ProgressBar
         mvarSys.MDI.CommandBars.StatusBar.Pane(3).Handle = PrgBar.hwnd
         PrgBar.Visible = True
         PrgBar.Value = 0
         PrgBar.ZOrder 0
         PrgBar.DragMode = 1
         PrgBar.UseVisualStyle = True
         PrgBar.Min = 1
         PrgBar.Max = mvarMe.GrdContato.Records.Count - 1 'mvarRs.recordcount
         iSend = 0
         For i = 0 To mvarMe.GrdContato.Records.Count - 1
            If mvarMe.GrdContato.Records(i).Item(0).Checked Then
               iToSend = iToSend + 1
            End If
         Next
         PrgBar.Max = iToSend
         nTotal = mvarMe.GrdContato.Records.Count
         sAux = ""
         sAux = sAux & "Total: ( " & nTotal & " Itens )"
         sAux = sAux & Space(10)
         sAux = sAux & "Selecionados: ( " & iToSend & " Itens )"
         Call SetTag(mvarMe.GrdContato, "PSEL", iToSend)
         mvarMe.LblStGrd.Caption = sAux
         bResult = True
         
         Dim iNome As Integer
         Dim ieMail As Integer
         iNome = 0
         ieMail = 0
         For i = 0 To mvarMe.GrdContato.Columns.Count - 1
            If UCase(mvarMe.GrdContato.Columns(i).Caption) = "NOME" Then iNome = i
            If UCase(mvarMe.GrdContato.Columns(i).Caption) = "EMAIL" Then ieMail = i
            If iNome <> 0 And ieMail <> 0 Then i = mvarMe.GrdContato.Columns.Count - 1
         Next
         If iNome = 0 Then
            ExibirAviso ("Coluna de Nome do cliente não definida, favor definir coluna com descrição 'Nome'")
            Exit Sub
         End If
         If ieMail = 0 Then
            ExibirAviso ("Coluna de e-Mail do cliente não definida, favor definir coluna com descrição 'eMail'")
            Exit Sub
         End If
         
        
         nGrupo = xVal(mvarMe.txtGrupo.Text)
         nGrupo = IIf(nGrupo <= 0, 1, nGrupo)
         If nGrupo <> 1 Then
            sNome = mvarMe.GrdContato.Records(i).Item(iNome).Value
            .RecipientDisplayName = "Cliente"
            .Recipient = .From
            Set cGrpMail = New Collection
            Set cGrpNome = New Collection
            Set cGrpKey = New Collection
            
            sMail = ""
            For i = 0 To mvarMe.GrdContato.Records.Count - 1
               If mvarMe.GrdContato.Records(i).Item(0).Checked Or (cGrpMail.Count > 0 And i = mvarMe.GrdContato.Records.Count - 1) Then
                  
                  If mvarMe.GrdContato.Records(i).Item(0).Checked Then
                     sMail = sMail & mvarMe.GrdContato.Records(i).Item(ieMail).Value & "; "
                     cGrpMail.Add mvarMe.GrdContato.Records(i).Item(ieMail).Value
                     cGrpNome.Add mvarMe.GrdContato.Records(i).Item(iNome).Value
                     cGrpKey.Add CStr(i)
                  End If
                  If cGrpMail.Count = nGrupo Or (cGrpMail.Count > 0 And i = mvarMe.GrdContato.Records.Count - 1) Then
                     If Mid(StrReverse(Trim(sMail)), 1, 1) = ";" Then
                        sMail = Mid(Trim(sMail), 1, Len(Trim(sMail)) - 1)
                     End If
                     If Trim(sMail) <> "" Then
                        If False Then
                           .BccRecipient = "diogenes72@bol.com.br; disantos@ig.com.br; diogenes72@gmail.com; diogenes1972@hotmail.com"
                           .Connect
                           .Send
                           .Disconnect
                           bResult = .SendSuccesful
                           mvarRs.MoveLast
                        Else
                           bResult = False
                           If nEnvio < 100 And nErro < 100 Then
                              .BccRecipient = sMail
                              .Connect
                              .Send
                              .Disconnect
                              bResult = .SendSuccesful
                           End If
                        End If
                        
                        If bResult Then
                           For k = 1 To cGrpMail.Count
                              nEnvio = nEnvio + 1
                              sNome = cGrpNome(k)
                              sMail = cGrpMail(k)
                              j = xVal(cGrpKey(k))
                              
                              Call WriteIniFile(sFileRes, "Enviados", "Qtd", CStr(nEnvio))
                              Call WriteIniFile(sFileRes, "Enviados", "Envio_" + StrZero(nEnvio, 4), Format(Now(), "dd/mm/yy hh:mm:ss") + " - " + sMail + String(50 - Len(sMail), " ") + sNome)
                              mvarMe.GrdContato.Records(j).Item(0).Checked = False
                           Next
                           'mvarMe.GrdContato.Records(i).Visible = False
                        Else
                           For k = 1 To cGrpMail.Count
                              nErro = nErro + 1
                              sNome = cGrpNome.Item(k)
                              sMail = cGrpMail.Item(k)
                              j = xVal(cGrpKey.Item(k))
                              
                              Call WriteIniFile(sFileRes, "Não Enviados", "Qtd", CStr(nErro))
                              Call WriteIniFile(sFileRes, "Não Enviados", "Erro_" + StrZero(nErro, 4), Format(Now(), "dd/mm/yy hh:mm:ss") + " - " + sMail + String(50 - Len(sMail), " ") + sNome)
                              mvarMe.GrdContato.Records(j).Item(0).Checked = True
                              bResult = False
                           Next
                        End If
                        iSend = iSend + cGrpMail.Count
                        
                        Set cGrpMail = New Collection
                        Set cGrpNome = New Collection
                        Set cGrpKey = New Collection
                        sMail = ""

                     End If
                     
                     mvarMe.GrdContato.Redraw
                     PrgBar.Value = iSend
                     
                     mvarSys.MDI.Caption = iSend & " / " & iToSend & "  " & sMail
                     If nEnvio + nErro <> 0 Then
                        Call PopulaStGrd(pEnviado:=nEnvio, pErro:=nErro)
                     End If
                  End If
               End If
            Next
         Else
            For i = 0 To mvarMe.GrdContato.Records.Count - 1
               If mvarMe.GrdContato.Records(i).Item(0).Checked Then
                  sNome = mvarMe.GrdContato.Records(i).Item(iNome).Value
                  If Trim(sNome) = "" Then
                     .RecipientDisplayName = "Cliente"
                  Else
                     .RecipientDisplayName = sNome
                  End If
                  
                  sMail = mvarMe.GrdContato.Records(i).Item(ieMail).Value
                  If Trim(sMail) <> "" Then
                     If False Then
                        .Recipient = "diogenes72@bol.com.br"
                        .Connect
                        .Send
                        .Disconnect
                        mvarRs.MoveLast
                     Else
                        .Recipient = sMail
                        .Connect
                        .Send
                        .Disconnect
                        If .SendSuccesful Then
                           nEnvio = nEnvio + 1
                           Call WriteIniFile(sFileRes, "Enviados", "Qtd", CStr(nEnvio))
                           Call WriteIniFile(sFileRes, "Enviados", "Envio_" + StrZero(nEnvio, 4), Format(Now(), "dd/mm/yy hh:mm:ss") + " - " + sMail + String(50 - Len(sMail), " ") + sNome)
                           mvarMe.GrdContato.Records(i).Item(0).Checked = False
                           'mvarMe.GrdContato.Records(i).Visible = False
                        Else
                           
                           nErro = nErro + 1
                           Call WriteIniFile(sFileRes, "Não Enviados", "Qtd", CStr(nErro))
                           Call WriteIniFile(sFileRes, "Não Enviados", "Erro_" + StrZero(nErro, 4), Format(Now(), "dd/mm/yy hh:mm:ss") + " - " + sMail + String(50 - Len(sMail), " ") + sNome)
                           mvarMe.GrdContato.Records(i).Item(0).Checked = True
                           'mvarMe.GrdContato.Records(i).Visible = True
                           bResult = False
                           
                        End If
                     End If
                     iSend = iSend + 1
                  End If
               Else
                  'mvarMe.GrdContato.Records(i).Visible = False
               End If
               mvarMe.GrdContato.Redraw
               PrgBar.Value = iSend
               
               mvarSys.MDI.Caption = iSend & " / " & iToSend & "  " & sMail
               If nEnvio + nErro <> 0 Then
                  Call PopulaStGrd(pEnviado:=nEnvio, pErro:=nErro)
               End If
            Next
         End If
'         While Not mvarRs.EOF
'            If Trim(mvarRs("EMAIL") & "") <> "" Then
'               If False Then
'                  .Recipient = "diogenes72@bol.com.br"
'                  .Connect
'                  .Send
'                  .Disconnect
'                  mvarRs.MoveLast
'               Else
'                  .Recipient = mvarRs("EMAIL")
'                  .Connect
'                  .Send
'                  .Disconnect
'               End If
'            End If
'            PrgBar.Value = mvarRs.AbsolutePosition
'            mvarSys.MDI.Caption = mvarRs.AbsolutePosition & " / " & mvarRs.recordcount
'            mvarRs.MoveNext
'         Wend
         PrgBar.Visible = False
      End If
'      bResult = .SendSuccesful
      'If Not bResult Then
      '   sMail = mvarMe.GrdContato.Records(i).Item(2).Value
      'End If
   End With
   Set xMail = Nothing
   mvarSys.MDI.Caption = sCaption
      
   Screen.MousePointer = vbDefault
   On Error Resume Next
   sMsg = ""
   sMsg = sMsg + "Mala direta executada com sucesso!" & vbNewLine & vbNewLine
   sMsg = sMsg + "Total: " & iSend & vbNewLine
   sMsg = sMsg + "Enviados: " & nEnvio & vbNewLine
   sMsg = sMsg + "Erros: " & nErro & vbNewLine
   MsgBox sMsg, vbOKOnly, "Mala direta"
   
Saida:
   Screen.MousePointer = vbDefault
End Sub
Private Sub CarregaErros(pFile As String)
   Dim i          As Integer
   Dim j          As Integer
   Dim sFileRes   As String
   Dim sMail      As String
   Dim sAux       As String
   Dim nErro      As String
   Dim nEnvio     As String
   Dim bAchou     As Boolean
   
   'Call GetNameFromPath(mvarMe.TxtHtml.Text, sFileRes)
   'sFileRes = sFileRes + "Envio_" + "150806170557" + ".txt"
   Call SetTag(mvarMe, "CAPTION", mvarMe.Caption)
   sFileRes = pFile
   
   nErro = xVal(ReadIniFile(sFileRes, "Não Enviados", "Qtd"))
   nEnvio = xVal(ReadIniFile(sFileRes, "Enviados", "Qtd"))
   For i = 0 To mvarMe.GrdContato.Records.Count - 1
      DoEvents
      sMail = mvarMe.GrdContato.Records(i).Item(3).Value
      bAchou = False
      If Not bAchou Then
         For j = 1 To nErro
            DoEvents
            sAux = ReadIniFile(sFileRes, "Não Enviados", "Erro_" + StrZero(j, 4))
            If InStr(UCase(sAux), UCase(sMail)) <> 0 Then
               j = nErro
               mvarMe.GrdContato.Records(i).Item(0).Checked = True
               bAchou = True
            End If
         Next
      End If
      If Not bAchou Then
         DoEvents
         For j = 1 To nEnvio
            sAux = ReadIniFile(sFileRes, "Enviados", "Envio_" + StrZero(j, 4))
            If InStr(UCase(sAux), UCase(sMail)) <> 0 Then
               j = nEnvio
               mvarMe.GrdContato.Records(i).Item(0).Checked = False
               bAchou = True
            End If
         Next
      End If
      If Not bAchou Then
         mvarMe.GrdContato.Records(i).Item(0).Checked = False
      End If
      mvarMe.GrdContato.Redraw
      mvarMe.Caption = GetTag(mvarMe, "CAPTION") & " - " & i & "/" & mvarMe.GrdContato.Records.Count - 1
      Call PopulaStGrd
   Next
   mvarMe.Caption = GetTag(mvarMe, "CAPTION")
End Sub
Private Sub mvarMe_CmdImportarClick()
   Dim ODlg  As Object
   Dim sFile As String
   Dim nPos As Integer
   Dim sMail As String
   Dim sTexto As String
   
   Set ODlg = mvarMe.CommonDialog1  '  CriarObjeto("MSComDlg.CommonDialog", False)
   sFile = SelecionarArquivo(ODlg, "Arquivo", "", "Query (*.Sql)|*.sql|Texto (*.txt)|*.txt|Planilha (*.xls, *.xlsx)|*.xls;*.xlsx")

   Screen.MousePointer = vbHourglass
   If ExisteArquivo(sFile) Then
      If UCase(StrReverse(Mid(StrReverse(sFile), 1, 3))) = "TXT" Then
         If "ENVIO_" = Mid(GetNameFromPath(sFile), 1, 6) Then
            Call CarregaErros(sFile)
         Else
            sTexto = ReadTextFile(sFile)
            While Trim(sTexto) <> ""
               nPos = InStr(sTexto, ";")
               sMail = Mid(sTexto, 1, nPos - 1)
               sTexto = Mid(sTexto, nPos + 1)
            Wend
            Call ExibirInformacao("Arquivo importado")
         End If
      ElseIf UCase(StrReverse(Mid(StrReverse(sFile), 1, 3))) = "SQL" Then
         Call CmdCarregar(ReadTextFile(sFile))
      End If
   Else
      mvarMe.GrdContato.Records.DeleteAll
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub mvarMe_CmdInverterClick()
   Dim i As Integer
   Screen.MousePointer = vbHourglass
   With mvarMe.GrdContato
      If .SelectedRows.Count > 1 Then
         For i = 0 To .SelectedRows.Count - 1
            mvarMe.GrdContato.SelectedRows(i).Record(0).Checked = Not mvarMe.GrdContato.SelectedRows(i).Record(0).Checked
         Next
      Else
         For i = 0 To .Records.Count - 1
            .Records(i).Item(0).Checked = Not .Records(i).Item(0).Checked
         Next
      End If
      .Populate
   End With
   Call PopulaStGrd
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_CmdNenhumClick()
   Dim i As Integer
   Screen.MousePointer = vbHourglass
'   DoEvents
   With mvarMe.GrdContato
      If .SelectedRows.Count > 1 Then
         For i = 0 To .SelectedRows.Count - 1
            mvarMe.GrdContato.SelectedRows(i).Record(0).Checked = False
         Next
      Else
         For i = 0 To .Records.Count - 1
            .Records(i).Item(0).Checked = False
         Next
      End If
      .Populate
   End With
   Call PopulaStGrd
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_CmdOpen(sFile As String)
'   Dim ODlg  As Object
'   If sFile = "" Then
'      Set ODlg = CriarObjeto("MSComDlg.CommonDialog", False)
'      sFile = SelecionarArquivo(ODlg, "Arquivo", "", "*.htm|*.htm")
'   End If
   
   If sFile = "" Then
      sFile = SelecionarArquivo(mvarMe.CommonDialog1, "Arquivo", "", "*.htm;*.html|*.htm;*.html")
   End If
   Call Navigate(sFile)
End Sub
Private Sub Navigate(Optional pFile As String = "")
   Dim bExiste As Boolean
   Dim sTexto As String
   Dim nIni As Integer
   Dim nFim As Integer
   
   If pFile = "" Then
      pFile = mvarMe.TxtHtml.Text
   End If
   mvarMe.TxtHtml.Text = pFile
   
   If ExisteArquivo(pFile) Then
      mvarMe.WebBrowser1.Navigate pFile
      If Trim(mvarMe.TxtTitulo.Text) = "" Then
         sTexto = ReadTextFile(pFile)
         nIni = InStr(UCase(sTexto), "<TITLE>")
         nFim = InStr(UCase(sTexto), "</TITLE>")
         If nIni > 0 And nFim > nIni Then
            mvarMe.TxtTitulo.Text = Trim(Mid(sTexto, nIni + 7, nFim - nIni - 7))
         End If
      End If
      mvarMe.CmdEnviar.Enabled = (pFile <> "") And (mvarMe.GrdContato.Records.Count > 0)
      mvarMe.TabControl1.Item(1).Enabled = (pFile <> "")
   Else
      mvarMe.CmdEnviar.Enabled = False
      mvarMe.TabControl1.Item(1).Enabled = False
   End If
End Sub
Private Sub mvarMe_CmdSairClick()
   Unload mvarMe
End Sub
Private Sub mvarMe_cmdSelecionarClick()
   Dim i
   Screen.MousePointer = vbHourglass
   With mvarMe.GrdContato
      For Each i In .SelectedRows
         '.Records(i).Item(0).Checked = True
         i.Record(0).Checked = True
      Next
      .Populate
   End With
   Call PopulaStGrd
   Screen.MousePointer = vbDefault
End Sub

Private Sub mvarMe_CmdTodosClick()
   Dim i As Integer
   Screen.MousePointer = vbHourglass
   With mvarMe.GrdContato
      If .SelectedRows.Count > 1 Then
         For i = 0 To .SelectedRows.Count - 1
            mvarMe.GrdContato.SelectedRows(i).Record(0).Checked = True
         Next
      Else
         For i = 0 To .Records.Count - 1
            .Records(i).Item(0).Checked = True
         Next
      End If
      .Populate
   End With
   Call PopulaStGrd
   Screen.MousePointer = vbDefault
End Sub

Private Sub mvarMe_GrdContatoItemCheck(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
   Call PopulaStGrd
End Sub

Private Sub mvarMe_GrdContatoKeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeySpace Then
      'mvarme.GrdContato.
   End If
End Sub

Private Sub mvarMe_Load()
   Call PosicionaForm
   mvarMe.TabControl1.Item(1).Enabled = True
   mvarMe.TabControl1.Item(0).Selected = True
   mvarMe.CmdImportar.Enabled = True
   mvarMe.txtGrupo.Text = 10
   mvarMe.txtGrupo.Enabled = True

   Call LimparGrid
End Sub
Private Sub mvarMe_GRdContatoRowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
   Dim TlContato As New TL_CadCliente
   Screen.MousePointer = vbHourglass
   With TlContato
      Set .Sys = mvarSys
      .IDCLIENTE = Val(Row.Record.Item(1).Value)
      If .IDCLIENTE <> 0 Then
         Call .Show
         If .IsDirt Then
            Call mvarMe_CmdCarregarClick
         End If
      End If
   End With
   Set TlContato = Nothing
   
   mvarMe.GrdContato.Redraw
   mvarMe.GrdContato.SetFocus
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_Resize()
   On Error Resume Next
   With mvarMe
      .TabControl1.Move 120, 120, .Width - 300, .Height - 600
      '=====================================================
      .TxtTitulo.Width = .TabControl1.Width - .TxtTitulo.Left - 120
      .WebBrowser1.Width = .TabControl1.Width - .WebBrowser1.Left - 120
      .WebBrowser1.Height = .TabControl1.Height - .WebBrowser1.Top - 800
      .TxtHtml.Top = .WebBrowser1.Height + .WebBrowser1.Top + 60
      .TxtHtml.Width = .TabControl1.Width - .TxtHtml.Left - (.CmdOpen.Width + .CmdEdit.Width + 280)
      .CmdOpen.Top = .TxtHtml.Top
      .CmdOpen.Left = .TxtHtml.Width + 120
      .CmdEdit.Top = .CmdOpen.Top
      .CmdEdit.Left = .CmdOpen.Left + .CmdOpen.Width
      '======================================================
      '.TabControl1.Item(1).Selected = True
      '.TabControl1.Item(1).Enabled = True
      '======================================================
      .CmdCarregar.Move 120, 60
      .CmdImportar.Move .CmdCarregar.Left + .CmdCarregar.Width + 60, .CmdCarregar.Top
      
      .GrdContato.Top = .CmdCarregar.Top + .CmdCarregar.Height + 60
      .GrdContato.Width = .TabControl1.Width - .GrdContato.Left - 180
      .GrdContato.Height = 0.83 * .TabControl1.Height  ' .TabControl1.Height - .GrdContato.Top - 840
      
      .LblStGrd.Top = .GrdContato.Top + .GrdContato.Height + 20
      .LblStGrd.Left = .GrdContato.Left
      .LblStGrd.Width = .GrdContato.Width
      
      .CmdSair.Left = .TabControl1.Width - .CmdSair.Width - 480
      .CmdEnviar.Left = .CmdSair.Left - .CmdEnviar.Width - 600
      '.CmdCarregar.Left = .CmdEnviar.Left - .CmdCarregar.Width - 480
      
      .CmdSair.Top = .LblStGrd.Top + .LblStGrd.Height + 60
      .CmdEnviar.Top = .CmdSair.Top
      .CmdSelecionar.Top = .CmdSair.Top
      .CmdTodos.Top = .CmdSair.Top
      .CmdNenhum.Top = .CmdSair.Top
      .CmdInverter.Top = .CmdSair.Top
      '=========================================================
      mvarMe.ChkAtivo.Top = 120 'mvarMe.GrdContato.Top - mvarMe.ChkAtivo.Height - 60
      mvarMe.ChkInativo.Top = mvarMe.ChkAtivo.Top
      mvarMe.ChkEmEspera.Top = mvarMe.ChkAtivo.Top
      mvarMe.ChkEmEspera.Left = mvarMe.GrdContato.Left + mvarMe.GrdContato.Width - mvarMe.ChkEmEspera.Width - 240
      mvarMe.ChkInativo.Left = mvarMe.ChkEmEspera.Left - mvarMe.ChkInativo.Width - 120
      mvarMe.ChkAtivo.Left = mvarMe.ChkInativo.Left - mvarMe.ChkAtivo.Width - 120
      
      mvarMe.txtGrupo.Visible = mvarMe.ChkInativo.Top
      mvarMe.Label2.Visible = True
      mvarMe.txtGrupo.Top = 120
      mvarMe.Label2.Top = mvarMe.txtGrupo.Top
      mvarMe.txtGrupo.Left = mvarMe.ChkAtivo.Left - mvarMe.txtGrupo.Width - 360
      mvarMe.Label2.Left = mvarMe.txtGrupo.Left - mvarMe.Label2.Width - 60
      
   End With
End Sub
Private Sub PosicionaForm()
   Dim Left As Long
   Dim Top As Long
   Dim Right As Long
   Dim Bottom As Long
   On Error Resume Next
   mvarSys.MDI.CommandBars.GetClientRect Left, Top, Right, Bottom
   mvarSys.MDI.CommandBars(1).GetWindowRect Left, Top, Right, Bottom
      
   mvarMe.Move mvarSys.MDI.Left + 120, Bottom, mvarSys.MDI.Width - mvarMe.Left - 240, mvarSys.MDI.Height - mvarMe.Top - 480
   mvarMe.Move mvarSys.MDI.Left + 120, Bottom, mvarSys.MDI.Width - mvarMe.Left - 240, mvarSys.MDI.Height - mvarMe.Top - 480
   
   'mvarMe.Top = Bottom
   'mvarMe.Left = mvarSys.MDI.Left + 60
   'mvarMe.Width = mvarSys.MDI.Width - mvarMe.Left - 120
   'mvarMe.Height = mvarSys.MDI.Height - mvarMe.Top - 420
End Sub

Private Sub mvarMe_Timer1Timer()
   Static nQtd  As Integer
   Dim sTitulo As String
   nQtd = nQtd + 1
   sTitulo = "[" & nQtd & "] upgrade - ocorrencia 5919412"
   
'   Call EnviarGolden("diogenes72@bol.com.br", sTitulo)
'   Call EnviarGolden("adriane76@bol.com.br", sTitulo)
'   Call EnviarGolden("disantos@ig.com.br", sTitulo)
   'Call EnviarGolden("alteracao.cadastral@goldencross.com.br")
'''   Call EnviarGolden("apoio.mpe@goldencross.com.br", sTitulo)
'''   Call EnviarGolden("yan.costa@goldencross.com.br", sTitulo)
   
   
   mvarSys.MDI.Caption = nQtd
   'mvarMe.Timer1.Interval = 60000
End Sub

Private Sub mvarMe_TxtHtmlLostFocus()
   Call Navigate
End Sub
Private Sub EnviarGolden(eMailTo As String, eTit As String)
   Dim PrgBar   As ProgressBar
   Dim bResult  As Boolean
   Dim xMail    As Object
   Dim sFile    As String
   Dim sCaption As String
   Dim sMessage As String
   Dim sTitulo  As String
   Dim sAtach   As String
   Dim sMail    As String
   Dim sAux     As String
   Dim sPath    As String
   Dim i        As Integer
   Dim iSend    As Integer
   
   Static nQtd As Integer
   
   Screen.MousePointer = vbHourglass
    sTitulo = eTit
   If Trim(sTitulo) = "" Then
      Call ExibirInformacao("Título inválido", "Mala Direta")
      GoTo Saida
   End If
   If Trim(mvarMe.TxtHtml.Text) = "" Or Not ExisteArquivo(mvarMe.TxtHtml.Text) Then
      Call ExibirInformacao("Arquivo inválido", "Mala Direta")
      GoTo Saida
   End If
   If mvarRs Is Nothing Then
      Call ExibirInformacao("Lista de clientes inválida", "Mala Direta")
      GoTo Saida
   End If
   If mvarRs.RecordCount <= 0 Then
      Call ExibirInformacao("Lista de clientes inválida", "Mala Direta")
      GoTo Saida
   End If
   
   sCaption = mvarSys.MDI.Caption

   sFile = mvarMe.TxtHtml.Text
   sTitulo = eTit
   sMessage = ReadTextFile(sFile)
   
   Set xMail = CriarObjeto("CAMail.SendMail")
   With xMail
      .UseAuthentication = (mvarSys.GetParam("UseAuthentication") = xtpChecked)
      .UsePopAuthentication = (mvarSys.GetParam("UsePopAuthentication") = xtpChecked)
      
      If True Then
         .POP3Host = "pop3.bol.com.br"
         .SMTPHost = "smtps.bol.com.br"
         .SMTPPort = 587
         .Username = "diogenes72@bol.com.br"
         .Password = "dolphin72"
      Else
         .POP3Host = "pop.ig.com.br"
         .SMTPHost = "smtp.ig.com.br"
         .SMTPPort = 587
         .Username = "disantos@ig.com.br"
         .Password = "dolphin"
      End If
      .FromDisplayName = "Diogenes Santos Ramos"
                  
      .Receipt = True
      .SMTPHostValidation = 0 'VALIDATE_HOST_NONE
                        
      .Subject = sTitulo
      .Message = sMessage
      
      .RecipientDisplayName = "GoldenCross"           ' "DiSantos"
      .From = .Username
      .AsHTML = True
      
      sFile = GetNameFromPath(sFile, sPath)
      'sAux = Dir(sPath, vbArchive)
      'While sAux <> ""
      '   If UCase(sAux) <> UCase(sFile) And Right(sAux, 4) <> ".htm" And Right(sAux, 5) <> ".html" Then
      '      sAtach = sAtach & IIf(sAtach = "", "", ";") & sPath & sAux
      '   End If
      '   sAux = Dir()
      'Wend
      sAtach = "C:\Tmp\Upgrade_Golden.pdf;C:\Tmp\Resposta_Yan.txt"
      If Trim(sAtach) <> "" Then
         .Attachment = sAtach
      End If
      .Subject = sTitulo
      
      
      .Recipient = eMailTo
      .Connect
      .Send
      .Disconnect
      
      bResult = .SendSuccesful
      
   End With
   Set xMail = Nothing
      
   Debug.Print bResult & " " & eMailTo & " " & Now()
   Screen.MousePointer = vbDefault
   On Error Resume Next
'   MsgBox IIf(bResult, "Mala direta executada com sucesso!" & vbNewLine & "(" & iSend & IIf(iSend <= 1, " mensagem envida.", " mensagens enviadas.)"), "Erro :" & sMail)
Saida:
   Screen.MousePointer = vbDefault
End Sub

