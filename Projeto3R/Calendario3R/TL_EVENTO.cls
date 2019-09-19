VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_Evento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe    As FrmEventoCal
Attribute mvarMe.VB_VarHelpID = -1

Dim bEventoNovo   As Boolean
Dim dToday        As Date
Dim ClFLGAREA     As Collection
Dim ClFLGAVALIACAO As Collection
Dim ClFLGDISPARO  As Collection
Dim ClFLGSESSAO   As Collection

Private mvarCalControl  As CalendarControl
Private mvarCalEvent    As CalendarEvent
Private mvarSys         As Object
Private mvarIDLOJA      As Integer
Private mvarisLoad      As Boolean
Private mvarIDEVENTO    As Long

Private mvarBeginSelection As Date
Private mvarEndSelection As Date
Private mvarAllDay As Boolean
Public Property Let IDLOJA(ByVal vData As Integer)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Integer
    IDLOJA = mvarIDLOJA
End Property
Public Property Let IDEVENTO(ByVal vData As Long)
   mvarIDEVENTO = vData
   If mvarCalEvent Is Nothing Then
      If Not mvarCalControl Is Nothing Then
         Set mvarCalEvent = mvarCalControl.DataProvider.GetEvent(mvarIDEVENTO)
      End If
   End If
End Property
Public Property Get IDEVENTO() As Long
    IDEVENTO = mvarIDEVENTO
End Property
Public Property Let isLoad(ByVal vData As Boolean)
    mvarisLoad = vData
End Property
Public Property Get isLoad() As Boolean
    isLoad = mvarisLoad
End Property
Public Property Get Form() As Object
   If mvarMe Is Nothing Then
      Set mvarMe = New FrmEventoCal
   End If
   Set Form = mvarMe
End Property
Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
   mvarIDLOJA = mvarSys.Propriedades("IDLOJA")
   On Error Resume Next
   dToday = CDate(mvarSys.xDb.Sysdate)
End Property
Public Property Get Sys() As Object
   Set Sys = mvarSys
End Property
Public Property Set CalEvent(ByVal vData As CalendarEvent)
   Set mvarCalEvent = vData
End Property
Public Property Get CalEvent() As CalendarEvent
   If mvarCalEvent Is Nothing Then
      If Not mvarCalControl Is Nothing Then
         Set mvarCalEvent = mvarCalControl.DataProvider.GetEvent(mvarIDEVENTO)
      End If
   End If
   
   Set CalEvent = mvarCalEvent
End Property
Public Property Set CalControl(ByVal vData As Object)
   Set mvarCalControl = vData
'   If mvarCalControl.Visible Then
'      mvarCalControl.DataProvider.ClearCache
'   End If
End Property
Public Property Get CalControl() As Object
    Set CalControl = mvarCalControl
End Property
Public Sub Show()
   bEventoNovo = False
   If mvarMe Is Nothing Then
      Set mvarMe = New FrmEventoCal
   End If
   DoEvents
   mvarMe.Show vbModal
   
End Sub
Public Sub Limpar()
   mvarIDEVENTO = 0
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmEventoCal
   Set mvarCalEvent = Nothing
End Sub
Private Sub Class_Terminate()
   On Error Resume Next
   Unload mvarMe
   Set mvarMe = Nothing
   Set mvarCalEvent = Nothing
   Set mvarCalControl = Nothing
End Sub
Private Sub mvarMe_Activate()
   Dim Record As ReportRecord
   Dim Item As ReportRecordItem
   Dim bChk As Boolean
   
   '* Monta e Popula Tela
'   Call MontaTela
'   Call PopulaTela
   DoEvents
   Screen.MousePointer = vbDefault
   If GetTag(mvarMe, "LOADED", "N") = "N" Then
      'mvarMe.GrdSERVICOEVT.Width = 5600
      'mvarMe.GrdSERVICOEVT.Left = 200
      Call SetTag(mvarMe, "LOADED", "S")
      If mvarMe.ChkAllDayEvent.Value = vbChecked Then
         If mvarMe.ChkMeeting.Value = vbUnchecked Then
            Call mvarMe_ChkMeetingClick
         Else
            mvarMe.ChkMeeting.Value = vbUnchecked
         End If
         If mvarIDEVENTO < 0 Then
            mvarMe.TxtLocation.Enabled = True
            mvarMe.TxtLocation.Text = ""
            mvarMe.TxtSubject.Text = "Tarefas"
            mvarMe.CmbLabel.ListIndex = 1
            mvarMe.CmbLabel.Enabled = False
            mvarMe.txtBody.SetFocus
         End If
      Else
         If mvarMe.ChkMeeting.Value = vbChecked Then
            Call mvarMe_ChkMeetingClick
         Else
            If bEventoNovo Then
               mvarMe.ChkMeeting.Value = vbChecked
            Else
               Call mvarMe_ChkMeetingClick
            End If
         End If
         
         If mvarMe.ChkMeeting.Value = vbUnchecked Then
            mvarMe.TabEvento(0).Selected = True
         End If
      End If
      Call ExibeDuracao
      Call VerificaMensagemI
   Else
     mvarisLoad = True
     mvarMe.isLoad = mvarisLoad
     'Call PopulaTela
   End If

'    Item.Format = "$ %s"

End Sub
Private Sub VerificaMensagemI()
   Dim Sql As String
   Dim sMsg As String
   
   Sql = "Select F.IDFATURA, C.NOME"
   Sql = Sql & " From FFATURA F"
   Sql = Sql & " Join OCLIENTE C On C.IDLOJA=F.IDLOJA And C.IDCLIENTE=F.IDCLIENTE"
   Sql = Sql & " Where F.IDLOJA=" & mvarIDLOJA
   Sql = Sql & " And F.SITFATURA = 0"
   Sql = Sql & " And F.IDCLIENTE = " & xVal(mvarMe.TxtIDCLIENTE.Text)
   
   With mvarSys.xDb
      If .AbreTabela(Sql) Then
         With .RsAux
            If .RecordCount > 0 Then
               sMsg = "Cliente: " & !NOME
               sMsg = Mid(sMsg, 1, 28) & IIf(Len(sMsg) > 30, "...", "") & vbNewLine
               sMsg = sMsg & "Possui fatura " & StrZero(!IDFATURA, 6) & " 'Em Aberto'"
               Call ExibirMensagemI(mvarSys, sMsg)
            End If
         End With
      End If
   End With
End Sub
Private Sub mvarMe_BtnCustomPropertiesClick()
'    If m_pEditingEvent Is Nothing Then
'        Exit Sub
'    End If
'
'    frmCustomEventProperties.SetEvent m_pEditingEvent
'
'    frmCustomEventProperties.Show vbModal, Me
End Sub

Private Sub mvarMe_BtnRecurrenceClick()
   
   Dim bisDirt As Boolean
   Dim MyRec   As TL_Recorrencia
   
   Call PopulaClasse_Evento
   
   Set MyRec = New TL_Recorrencia
   With MyRec
      Set .CalEvent = mvarCalEvent.CloneEvent
      'If mvarCalControl.ActiveView.Selection.IsValid Then
      '   .CalBegin mvarCalControl.ActiveView.Selection.Begin
      'End If
      .CalBegin = DateFromString(mvarMe.CmbStartDate, mvarMe.CmbStartTime)
      .Show
      'bisDirt = (mvarCalEvent.RecurrenceState <> .CalEvent.RecurrenceState)
      
      If .isDirt Then
         Set mvarCalEvent = .CalEvent
         Call PopulaTela_Evento
      End If
   End With
End Sub
Private Sub mvarMe_ChkAllDayEventClick()
   Dim iTab As Integer
   Dim i    As Integer
   
   If mvarMe.ChkMeeting.Value = 1 Then
      iTab = 1
   Else
      For i = 0 To mvarMe.TabEvento.ItemCount - 1
         If mvarMe.TabEvento(i).Selected Then
            iTab = i
            Exit For
         End If
      Next
   End If
   If mvarMe.ChkAllDayEvent.Value = 1 Then
      mvarMe.ChkMeeting.Value = 0
      mvarMe.ChkMeeting.Enabled = False
      mvarMe.TabEvento.SelectedItem = 0
      mvarMe.TabEvento.Item(1).Visible = False
      mvarMe.TxtLocation.Enabled = True
      mvarMe.CmbLabel.ListIndex = 1
      mvarMe.CmbLabel.Enabled = False
      
      mvarMe.CmbEndDate.Enabled = False
      Call SetTag(mvarMe.CmbEndDate, "TEXT", mvarMe.CmbEndDate.Value)
      Call SetTag(mvarMe.CmbStartTime, "TEXT", mvarMe.CmbStartTime.Text)
      Call SetTag(mvarMe.CmbEndTime, "TEXT", mvarMe.CmbEndTime.Text)
      mvarMe.CmbEndTime.Enabled = False
      mvarMe.CmbStartTime.Enabled = False
      mvarMe.CmbEndDate.Value = mvarMe.CmbStartDate.Value
      mvarMe.CmbEndTime.Text = ""
      mvarMe.CmbStartTime.Text = ""
      
      'mvarMe.TabEvento(1).Visible = False
      'mvarMe.TabEvento(2).Visible = (mvarMe.ChkMeeting.Value = 0)
      mvarMe.TabEvento(0).Selected = True
      
    Else
      mvarMe.TabEvento.Item(1).Visible = True
      mvarMe.CmbLabel.ListIndex = -1
      mvarMe.CmbLabel.Enabled = True
      mvarMe.ChkMeeting.Enabled = True
      mvarMe.CmbEndDate.Enabled = False
      mvarMe.CmbEndTime.Enabled = True
      mvarMe.CmbStartTime.Enabled = True
      If mvarMe.Visible Then
         If IsDate(GetTag(mvarMe.CmbEndDate, "TEXT", mvarMe.CmbEndDate.Value)) Then
            'mvarMe.CmbEndDate.Value = GetTag(mvarMe.CmbEndDate, "TEXT", mvarMe.CmbEndDate.Value)
            mvarMe.CmbEndDate.Value = mvarMe.CmbStartDate.Value
         End If
         mvarMe.CmbEndTime.Text = GetTag(mvarMe.CmbEndTime, "TEXT", mvarMe.CmbEndTime)
         mvarMe.CmbStartTime.Text = GetTag(mvarMe.CmbStartTime, "TEXT", mvarMe.CmbStartTime)
      End If
      
      'mvarMe.TabEvento(1).Visible = True
      'mvarMe.TabEvento(2).Visible = (mvarMe.ChkMeeting.Value = 0)
      mvarMe.TabEvento(1).Selected = True
   End If
   'mvarMe.cmbEndTime.Enabled = IIf(mvarMe.ChkAllDayEvent.Value = 1, False, True)
   'mvarMe.cmbStartTime.Enabled = IIf(mvarMe.ChkAllDayEvent.Value = 1, False, True)
   
   If mvarMe.TabEvento(iTab).Visible Then
      mvarMe.TabEvento(iTab).Selected = True
   End If
End Sub

Private Sub mvarMe_ChkFLGAVALIACAOClick()
   Dim Sql As String
   Dim MyRS As Object
   Dim nLin As Integer
   Dim nCol As Integer
   Dim bAval   As Boolean
      
   If mvarSys.Propriedades("TPSERV") = "0" And mvarMe.ChkFLGAVALIACAO.Value = vbChecked Then
      mvarMe.ChkFLGAVALIACAO.Value = vbUnchecked
      Exit Sub
   End If
   With mvarMe.GrdSERVICOEVT
      bAval = (mvarMe.ChkFLGAVALIACAO.Value = vbChecked)
      If mvarSys.Propriedades("TPSERV") = "0" Then
         .ColTag("IDTPSERVICO") = SetTag(.ColTag("IDTPSERVICO"), "EDIT", False)
      Else
         .ColTag("IDTPSERVICO") = SetTag(.ColTag("IDTPSERVICO"), "EDIT", IIf(bAval, False, (.Combos("IDTPSERVICO").ListCount > 2)))
      End If
      
      If mvarMe.Visible Then
         With .Combos("IDTPTRATAMENTO")
            .Clear
            Sql = Sql_OTPTRATAMENTO(mvarIDLOJA)
            Sql = Sql & " And ATIVO=1" + IIf(bAval, " And FLGAVALIACAO=1", "")
            Sql = Sql & IIf(bAval, " And FLGAVALIACAO=1", "")
            Sql = Sql & " Order By DSCTRATAMENTO"
            If mvarSys.xDb.AbreTabela(Sql, MyRS) Then
               While Not MyRS.EOF
                  .AddItem MyRS("DSCTRATAMENTO").Value, MyRS("IDTPTRATAMENTO").Value
                  MyRS.MoveNext
               Wend
            End If
            .AutoAdjustWidth
         End With
      End If
      
      For nLin = 1 To .RowCount - 1
         .CellValue(nLin, "IDTPSERVICO") = IIf(bAval, 1, 2)
         For nCol = 1 To .ColCount
            If GetTag(.ColTag(nCol), "EDIT", True) = False Then
               .CellForeColor(nLin, nCol) = vbGrayText
            Else
               .CellForeColor(nLin, nCol) = vbBlack
            End If
         Next
      Next
   End With
End Sub

Private Sub mvarMe_ChkMeetingClick()
   Dim vCtrl   As Control
   Dim nCol    As Long
   Dim nNew    As Long
   Dim sNewTxt As String
   Dim bCheked As Boolean
   Dim bCancel As Boolean
   
   With mvarMe
      bCheked = (.ChkMeeting.Value = vbChecked)
      bCancel = (.ChkFLGCANCELADO.Value = vbChecked)
      
      If GetTag(.GrpSessao, "IDATENDIMENTO", "0") = "0" Then
         .GrpSessao.Enabled = bCheked
         For Each vCtrl In .Controls
            If TypeName(vCtrl) <> "ImageList" Then
               If vCtrl.Container Is .GrpSessao Then
                  vCtrl.Enabled = .GrpSessao.Enabled
               End If
            End If
         Next
      End If
      .ChkFLGCANCELADO.Enabled = False
      .ChkFLGREMARCADO.Enabled = False

      .CmbShowTimeAs.Enabled = Not .GrpSessao.Enabled
      Call MontaCmbShowTimeAs(bCheked)
      
      If .GrpSessao.Enabled And .GrdSERVICOEVT.RowCount > 1 Then
         nCol = .GrdSERVICOEVT.ColIndex("IDTPSERVICO")
         nNew = .GrdSERVICOEVT.CellValue(1, "IDTPSERVICO")
         sNewTxt = .GrdSERVICOEVT.CellText(1, "IDTPSERVICO")
         If mvarSys.Propriedades("TPSERV") <> "0" Then
            If nNew = 0 Then mvarMe.ChkFLGAVALIACAO.Value = vbChecked
         End If
         If mvarMe.GrdSERVICOEVT.RowCount > 1 Then
            Call mvarMe_GrdSERVICOEVTBeforeCommitEdit(1, nCol, igEditResCommit, sNewTxt, nNew, 0)
         
            nCol = .GrdSERVICOEVT.ColIndex("IDTPTRATAMENTO")
            nNew = .GrdSERVICOEVT.CellValue(1, "IDTPTRATAMENTO")
            sNewTxt = .GrdSERVICOEVT.CellText(1, "IDTPSERVICO")
            Call mvarMe_GrdSERVICOEVTBeforeCommitEdit(1, nCol, igEditResCommit, sNewTxt, nNew, 0)
         End If
      End If
      
      .CmdAtendimento.Enabled = bCheked And (dToday >= CDate(Format(.CmbEndDate.Value, "dd/mm/yyyy"))) And .CmdOk.Enabled 'And xVal(mvarMe.TxtIDCLIENTE) <> 0
      .TxtLocation.Enabled = Not bCheked
      .TxtLocation.Visible = Not bCheked
      .CmbSchedule.Visible = bCheked
      '.TabEvento(2).Visible = Not bCheked
      
      If bCheked Then
         mvarMe.TabEvento.Item(1).Selected = True
      Else
         mvarMe.TabEvento.Item(0).Selected = True
      End If
      mvarMe.TabEvento.Item(1).Visible = bCheked
      mvarMe.TabEvento.Item(2).Visible = bCheked
      
      If mvarMe.TxtNOME.Enabled And mvarMe.TxtNOME.Visible Then
         mvarMe.TxtNOME.SetFocus
      Else
         If mvarMe.CmbNOME.Enabled And mvarMe.CmbNOME.Visible Then
            mvarMe.CmbNOME.SetFocus
         ElseIf mvarMe.TxtSubject.Enabled And mvarMe.TxtSubject.Visible Then
            mvarMe.TxtSubject.SetFocus
         End If
      End If
   End With
End Sub

Private Sub mvarMe_ChkReminderClick()
   With mvarMe
      .CmbReminder.Enabled = IIf(.ChkReminder.Value > 0, True, False)
      .CmbReminder.BackColor = IIf(.ChkReminder.Value > 0, RGB(255, 255, 255), RGB(210, 210, 210))
   End With
End Sub
Private Sub mvarMe_CmbEndTimeChange()
   Call ExibeDuracao
End Sub
Private Sub mvarMe_CmbEndTimeClick()
   Dim Index As Long
   Index = InStr(1, mvarMe.CmbEndTime.Text, "(")
   If Index > 0 Then
      mvarMe.CmbEndTime.Text = Left(mvarMe.CmbEndTime.Text, Index - 2)
   End If
'   Call ExibeDuracao
End Sub

Private Sub mvarMe_CmbEndTimeLostFocus()
   Dim sTime   As String
   
   mvarMe.CmbEndTime.Text = FormatarHora(mvarMe.CmbEndTime.Text)
    
   sTime = mvarMe.CmbEndTime.Text
   If Not IsDate(sTime) Then
      sTime = Format(DateAdd("n", 30, mvarMe.CmbStartTime.Text), "hh:mm")
      mvarMe.CmbEndTime.Text = sTime
   End If
   Call ExibeDuracao
End Sub

Private Sub mvarMe_CmbLabelClick()
   Dim pLabel As CalendarEventLabel
   Dim nLabelID As Long
   
   With mvarMe
      If .CmbLabel.ListIndex >= 0 Then
         If .CmbLabel.ItemData(.CmbLabel.ListIndex) >= 6 Or (.CmbLabel.ListIndex = 0 And .ChkMeeting.Value = 0) Then
            nLabelID = .CmbLabel.ItemData(.CmbLabel.ListIndex)
            Set pLabel = mvarCalControl.DataProvider.LabelList.Find(nLabelID)
            If Not pLabel Is Nothing Then
               .ctrlColor.BackColor = pLabel.Color
            End If
         End If
      End If
   End With
End Sub
Private Sub ExibeDuracao()
   Dim sTempo As String
   Dim nTempo As Integer
   Dim nInt   As Integer
   
   On Error Resume Next
   nTempo = DateDiff("n", mvarMe.CmbStartTime.Text, mvarMe.CmbEndTime.Text)
   sTempo = ""
   
   mvarMe.LblDuracao.Caption = "Duração: "
   If nTempo > 0 Then
      nInt = Int(nTempo / 60)
      sTempo = StrZero(nInt, 2) & ":" & StrZero(nTempo Mod (IIf(nInt = 0, 1, nInt) * 60), 2)
   End If
   If sTempo <> "" Then
      mvarMe.LblDuracao.Caption = mvarMe.LblDuracao.Caption & sTempo & "h"
   End If
End Sub
Private Sub mvarMe_CmbNOMEChange()
   mvarMe.TxtNOME.Text = mvarMe.CmbNOME.Text
   'Call MontaSubject
End Sub
Private Sub mvarMe_CmbStartTimeChange()
'   Call ExibeDuracao
End Sub
Private Sub mvarMe_CmbStartTimeClick()
'   Call ExibeDuracao
   Call DefineEndTime
End Sub
Private Sub DefineEndTime()
   Dim sTime  As String
    
   mvarMe.CmbStartTime.Text = FormatarHora(mvarMe.CmbStartTime.Text)
   sTime = mvarMe.LblDuracao.Caption
   sTime = Trim(Replace(sTime, "Duração:", ""))
   sTime = Trim(Replace(sTime, "Duração :", ""))
   sTime = Trim(Replace(sTime, "h", ""))
   If IsDate(sTime) And IsDate(mvarMe.CmbStartTime.Text) Then
      mvarMe.CmbEndTime.Text = Format(CDate(mvarMe.CmbStartTime.Text) + CDate(sTime), "hh:mm")
   End If

End Sub

Private Sub mvarMe_CmbStartTimeLostFocus()
   Call DefineEndTime
   Call ExibeDuracao
End Sub
Private Sub mvarMe_CmdAtendimentoClick()
   Dim TlAtend As TL_Atendimento
   
   Screen.MousePointer = vbHourglass
'   If Not VerificaSessao() Then Exit Sub
   Call SalvarEvento

   If mvarCalEvent.Id > 0 Then
      Set TlAtend = New TL_Atendimento
      With TlAtend
         Set .Sys = mvarSys
         Set .CalEvent = mvarCalEvent
         .IDLOJA = mvarIDLOJA
         .IDSALA = xVal(mvarCalEvent.CustomProperties("IDSALA"))
         .IDEVENTO = mvarCalEvent.Id
         
         
         mvarMe.CmdAtendimento.Enabled = False
         mvarMe.CmdOk.Enabled = False
         
         mvarMe.Hide
         
         .Show vbModal
         
         'If .IDATENDIMENTO <> 0 Then mvarMe.Hide
         'mvarMe.CmdOk.Enabled = True
         
         Call PopulaTela_Evento
         
         If .IDATENDIMENTO <> 0 Then
            If Not mvarCalEvent Is Nothing Then
               If mvarCalEvent.Id > 0 Then
                  mvarCalEvent.CustomProperties("EXIBEMSG") = False
                  mvarCalControl.DataProvider.ChangeEvent mvarCalEvent
                  mvarCalControl.Enabled = True
                  mvarCalControl.RedrawControl
                  mvarCalControl.Populate
                  mvarCalEvent.CustomProperties("EXIBEMSG") = True
                  Unload mvarMe
               Else
                  On Error Resume Next
                  mvarMe.Show
                  mvarMe.CmdAtendimento.Enabled = True
               End If
            End If
         Else
            On Error Resume Next
            mvarMe.Show
            mvarMe.CmdAtendimento.Enabled = True
         End If
      End With
   End If
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_CmdCancelClick()
   On Error Resume Next
   'If Not mvarCalEvent Is Nothing Then
   '   If mvarCalEvent.Id > 0 Then
   '      mvarCalControl.DataProvider.ChangeEvent mvarCalEvent
   '   End If
   'End If
   Unload mvarMe
End Sub

Private Sub mvarMe_CmdOkClick()
   If Not VerificaDatas() Then Exit Sub
   If Not VerificaSessao() Then Exit Sub

   Call SalvarEvento

   On Error Resume Next
   If mvarCalEvent.Id > 0 Then
      Unload mvarMe
   End If
   
'   If mvarMe.ChkMeeting.Value = vbChecked Then
'      If GetTag(mvarMe.ChkMeeting, "bResult", "True") = "True" Then
'         Unload mvarMe
'      End If
'   Else
'      Unload mvarMe
'   End If
   
   
End Sub
Private Sub SalvarEvento()
   If mvarMe.ChkMeeting.Value Then
      If mvarMe.GrdSERVICOEVT.RowCount <= 2 Then
         If Val(mvarMe.GrdSERVICOEVT.CellValue(1, "IDTPSERVICO")) = 0 Then
            mvarMe.CmbEndTime.Text = mvarMe.CmbStartTime.Text
         End If
      End If
   End If
   
   Call PopulaClasse_Evento

   If bEventoNovo Then
      mvarCalControl.DataProvider.AddEvent mvarCalEvent
      bEventoNovo = (mvarCalEvent.Id <= 0)
   Else
      mvarCalControl.DataProvider.ChangeEvent mvarCalEvent
   End If
End Sub

Private Sub mvarMe_GrdSERVICOEVTRequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
   With mvarMe.GrdSERVICOEVT
      bCancel = (GetTag(.ColTag(lCol), "EDIT", True) = False)
      If .CellFmtString(lRow, lCol) <> "" Then sText = .CellValue(lRow, lCol)
      eTextEditOpt = GetTag(.ColTag(lCol), "EDITOPT", 0)
   End With
End Sub

Private Sub mvarMe_KeyUp(KeyCode As Integer, Shift As Integer)
   Dim lRow As Long
   Dim nIDAREA As Integer
   Dim nIDTPTRATAMENTO As Integer
   
   If KeyCode = vbKeyDelete Then
      If mvarMe.ActiveControl Is mvarMe.GrdSERVICOEVT Then
         With mvarMe.ActiveControl
            lRow = .CurRow
            If lRow = 0 Then Exit Sub
            nIDAREA = xVal(.CellValue(lRow, "IDAREA"))
            nIDTPTRATAMENTO = xVal(.CellValue(lRow, "IDTPTRATAMENTO"))
            'If mvarMe.GrdSERVICOEVT.ColVisible("IDAREA") Then
            '   If nIDAREA = 0 Then Exit Sub
            'End If
            If lRow <> .RowCount And lRow > 0 Then
               .RemoveRow (.CurRow)
               .Tag = SetTag(.Tag, "ISDIRT", 1)
               .SetFocus
               If lRow < .RowCount Then
                  .SetCurCell lRow, 1
               ElseIf lRow > 1 Then
                  .SetCurCell lRow - 1, 1
               End If
            End If
            For lRow = 1 To .RowCount - 1
               .CellValue(lRow, 1) = lRow
            Next
            
            '* Excluir Tempo
            Dim MyArea As Object
            Dim nTempo As Integer
            Dim nIDTPMAN As Integer
            Dim sSexo As String
         
            nIDTPMAN = 1
            sSexo = "F"
            Set MyArea = CriarObjeto("BANCO_3R.TB_OAREA_TRATAMENTO")
            Set MyArea.xDb = mvarSys.xDb
            nTempo = 0
            If MyArea.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDAREA:=nIDAREA, CH_IDTPTRATAMENTO:=nIDTPTRATAMENTO, Ch_IDTPMANIPULO:=nIDTPMAN) Then
               nTempo = -1 * (MyArea.TEMPOMED * IIf(sSexo = "F", 1, MyArea.INDICEH))
            End If
            If DateDiff("n", mvarMe.CmbStartTime.Text, mvarMe.CmbEndTime.Text) + nTempo > 0 Then
               mvarMe.CmbEndTime.Text = DateAdd("n", nTempo, mvarMe.CmbEndTime.Text)
            End If
         End With
      End If
      
   ElseIf KeyCode = vbKeyF2 Then
      Call mvarMe_CmdAtendimentoClick
   End If
End Sub

Private Sub mvarMe_GrdSERVICOEVTAfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
   Dim MyArea As Object
   Dim MyCliente As Object
   Dim Sql As String
   Dim nTempo As Integer
   Dim nIDTPMAN As Integer
   Dim sSexo As String
   Dim sTag As String
   Dim sEndTime As String
   Dim bTeste  As Boolean
   Dim i As Integer
   
   With mvarMe.GrdSERVICOEVT
      If lCol < .ColCount Then
         Call .SetCurCell(lRow, lCol + 1)
      Else
         Call .SetCurCell(lRow, lCol - 1)
      End If
      Call .SetCurCell(lRow, lCol)
   
      sTag = ""
      If Not IsMissing(.RowTag(lRow)) Then
         sTag = .RowTag(lRow)
         If GetTag(sTag, "IDAREA2", 0) = 0 Then
            .RowTag(lRow) = SetTag(sTag, "IDAREA2", GetTag(sTag, "IDAREA", 0))
            '.RowTag(lRow) = SetTag(sTag, "IDAREA2", .CellValue(lRow, "IDAREA"))
         End If
      End If

      If .ColKey(lCol) = "IDAREA" Or .ColKey(lCol) = "IDTPSERVICO" Then
         If .RowCount <= 2 Then
            If xVal(.CellValue(1, "IDTPSERVICO")) = 0 Then
               .CellValue(1, "IDTPSERVICO") = 2
            End If
            If xVal(.CellValue(1, "IDTPTRATAMENTO")) = 0 Then
               .CellValue(1, "IDTPTRATAMENTO") = 1
            End If
         End If
         
         '**************
         '* Incluir Tempo
         If GetTag(mvarMe.GrdSERVICOEVT, "CARREGANDO", 0) = 0 Then
            If GetTag(sTag, "IDAREA2", 0) <> .CellValue(lRow, "IDAREA") Or GetTag(sTag, "IDTPSERVICO2", 0) <> .CellValue(lRow, "IDTPSERVICO") Then
               nIDTPMAN = 1
               sSexo = "F"
               If Val(mvarMe.TxtIDCLIENTE) <> 0 Then
                  Sql = "Select SEXO"
                  Sql = Sql & " From OCLIENTE"
                  Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
                  Sql = Sql & " And IDCLIENTE=" & SqlNum(mvarMe.TxtIDCLIENTE.Text)
                  If mvarSys.xDb.AbreTabela(Sql) Then
                     sSexo = mvarSys.xDb.RsAux("SEXO")
                  End If
               End If
               
               '**********
               '* Deduzir o tempo da área modificada
               Sql = "Select *"
               Sql = Sql & " From OAREA_TRATAMENTO"
               Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
               Sql = Sql & " And IDAREA=" & SqlNum(GetTag(sTag, "IDAREA2", 0))
               Sql = Sql & " And IDTPTRATAMENTO=" & SqlNum(.CellValue(1, "IDTPTRATAMENTO"))
               Sql = Sql & " And IDTPMANIPULO=" & nIDTPMAN
               If mvarSys.xDb.AbreTabela(Sql) Then
                  If GetTag(sTag, "IDTPSERVICO2", 0) = 1 Then
                     If .CellValue(1, "IDTPSERVICO") = 1 Then
                        nTempo = -10
                     End If
                  Else
                     nTempo = -1 * (mvarSys.xDb.RsAux("TEMPOMED") * IIf(sSexo = "F", 1, mvarSys.xDb.RsAux("INDICEH")))
                  End If
               ElseIf lRow = 1 Then
                  nTempo = DateDiff("n", mvarMe.CmbEndTime.Text, mvarMe.CmbStartTime.Text)
               End If
                              
               '**********
               '* Adicionar tempo da área atual
               If .CellValue(lRow, "IDTPSERVICO") = 1 Then
               'GetTag(sTag, "IDAREA2", 0)
                  '* Se for teste, tempo padrão de 10 minutos
                  If .RowCount = 2 Then
                     nTempo = nTempo + IIf(lRow = 1, 10, 0)
                  End If
               Else
                  Sql = "Select *"
                  Sql = Sql & " From OAREA_TRATAMENTO"
                  Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
                  Sql = Sql & " And IDAREA=" & SqlNum(.CellValue(lRow, "IDAREA"))
                  Sql = Sql & " And IDTPTRATAMENTO=" & SqlNum(.CellValue(lRow, "IDTPTRATAMENTO"))
                  Sql = Sql & " And IDTPMANIPULO=" & nIDTPMAN
                  If mvarSys.xDb.AbreTabela(Sql) Then
                     nTempo = nTempo + (mvarSys.xDb.RsAux("TEMPOMED") * IIf(sSexo = "F", 1, mvarSys.xDb.RsAux("INDICEH")))
                  End If
               End If
               If .CellValue(lRow, "IDTPSERVICO") = 1 Then
                  bTeste = True
                  For i = 1 To .RowCount - 1
                     bTeste = bTeste And (.CellValue(i, "IDTPSERVICO") = 1)
                  Next
                  If bTeste Or DateDiff("n", mvarMe.CmbStartTime.Text, mvarMe.CmbEndTime.Text) + nTempo < 10 Then
                     nTempo = 10 - DateDiff("n", mvarMe.CmbStartTime.Text, mvarMe.CmbEndTime.Text)
                  End If
               End If
               
               If DateDiff("n", mvarMe.CmbStartTime.Text, mvarMe.CmbEndTime.Text) + nTempo > 0 Then
                  sEndTime = Format(DateAdd("n", nTempo, mvarMe.CmbEndTime.Text), "hh:mm")
                  If -1 = LocalizarCombo(mvarMe.CmbEndTime, sEndTime) Then
                     mvarMe.CmbEndTime.Text = sEndTime
                  End If
               End If
               .RowTag(lRow) = SetTag(sTag, "IDAREA2", .CellValue(lRow, "IDAREA"))
               .RowTag(lRow) = SetTag(sTag, "IDTPSERVICO2", .CellValue(lRow, "IDTPSERVICO"))
               Call ExibeDuracao
            End If
         End If
      ElseIf .ColKey(lCol) = "IDTPTRATAMENTO" Then
         If ClFLGAREA("k" & .CellValue(lRow, "IDTPTRATAMENTO")) = 0 Then
            .CellValue(lRow, "IDAREA") = 0
         Else
            If xVal(.CellValue(lRow, "IDAREA")) = 0 Then .CellValue(lRow, "IDAREA") = ""
         End If
         If ClFLGAVALIACAO("k" & .CellValue(lRow, "IDTPTRATAMENTO")) = 0 Then
            .CellValue(lRow, "IDTPSERVICO") = 2
         End If
      
      End If
   End With
End Sub

Private Sub mvarMe_GrdSERVICOEVTBeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
   Dim pLabel     As CalendarEventLabel
   Dim nLabelID   As Long
   Dim sKey       As String
   Dim nIDSERV    As Integer
   Dim nIDTRAT    As Integer
   Dim nCancelado As Integer
   
   With mvarMe
      sKey = .GrdSERVICOEVT.ColKey(lCol)
      If lRow = 1 Then
         If sKey = "IDTPSERVICO" Or sKey = "IDTPTRATAMENTO" Then
            nIDSERV = xVal(IIf(sKey = "IDTPSERVICO", vNewValue, .GrdSERVICOEVT.CellValue(lRow, "IDTPSERVICO")))
            nIDTRAT = xVal(IIf(sKey = "IDTPTRATAMENTO", vNewValue, .GrdSERVICOEVT.CellValue(lRow, "IDTPTRATAMENTO")))
            If xVal(mvarMe.TxtIDCLIENTE.Text) = 0 Then
               nIDSERV = IIf(nIDSERV = 0, 1, nIDSERV)
               nIDTRAT = IIf(nIDTRAT = 0, 1, nIDTRAT)
               .GrdSERVICOEVT.CellValue(1, "IDTPSERVICO") = nIDSERV
               .GrdSERVICOEVT.CellValue(1, "IDTPTRATAMENTO") = nIDTRAT
            End If
            nCancelado = .ChkFLGCANCELADO.Value
            
            .CmbShowTimeAs.Enabled = False
            .CmbShowTimeAs.ListIndex = DefineShowAs(nIDSERV)
            
            nLabelID = DefineLabelID(nIDTRAT, nIDSERV, nCancelado)
            If Not mvarCalControl Is Nothing Then
               Set pLabel = mvarCalControl.DataProvider.LabelList.Find(nLabelID)
               If Not pLabel Is Nothing Then
                  .ctrlColor.BackColor = pLabel.Color
               End If
            End If
         End If
         If Not IsMissing(.GrdSERVICOEVT.RowTag(lRow)) Then
            .GrdSERVICOEVT.RowTag(lRow) = SetTag(.GrdSERVICOEVT.RowTag(lRow), "IDAREA2", .GrdSERVICOEVT.CellValue(lRow, "IDAREA"))
            .GrdSERVICOEVT.RowTag(lRow) = SetTag(.GrdSERVICOEVT.RowTag(lRow), "IDTPSERVICO2", .GrdSERVICOEVT.CellValue(lRow, "IDTPSERVICO"))
         End If
         
      End If
   End With
End Sub
Private Sub mvarMe_GrdSERVICOEVTColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
   bDoDefault = False
   With mvarMe.GrdSERVICOEVT
      If .RowCount > 1 Then
         If .ColKey(lCol) = "IDTPSERVICO" Then
            If xVal(.CellValue(1, lCol)) = 0 Then
               .CellValue(1, lCol) = 2
               If .Combos("IDTPSERVICO").ListCount > 0 Then
                  Call mvarMe_GrdSERVICOEVTBeforeCommitEdit(1, lCol, igEditResCommit, .Combos("IDTPSERVICO").ItemText(1), 2, 0)
               End If
            End If
         ElseIf .ColKey(lCol) = "IDTPTRATAMENTO" Then
            If xVal(.CellValue(1, lCol)) = 0 Then
               .CellValue(1, lCol) = 1
               If .Combos("IDTPTRATAMENTO").ListCount > 0 Then
                  Call mvarMe_GrdSERVICOEVTBeforeCommitEdit(1, lCol, igEditResCommit, .Combos("IDTPTRATAMENTO").ItemText(0), 1, 0)
               End If
            End If
         End If
      End If
   End With
End Sub
Private Sub mvarMe_GrdSERVICOEVTLostFocus()
   Dim sTag As String
   Dim i    As Integer
   Dim j    As Integer
   
   With mvarMe.GrdSERVICOEVT
      '.Tag = SetTag(.Tag, "ISDIRT", 0)
      If .RowCount <= 1 Then .Tag = SetTag(.Tag, "ISDIRT", 1)
      For i = 1 To .RowCount - 1
         If Not IsMissing(.RowTag(i)) Then sTag = .RowTag(i)
         For j = 1 To .ColCount
            If GetTag(sTag, .ColKey(j), 0) <> .CellValue(i, j) Then .Tag = SetTag(.Tag, "ISDIRT", 1)
            If GetTag(.Tag, "ISDIRT", 0) = 1 Then Exit For
         Next
         If GetTag(.Tag, "ISDIRT", 0) = 1 Then Exit For
      Next
   End With
End Sub

Private Sub mvarMe_GrdSERVICOEVTMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
   If GetTag(mvarMe, "LOADED", "N") = "N" Then Exit Sub
   With mvarMe.GrdSERVICOEVT
      If lRow = .RowCount Then
         If lRow > 1 Then
            If Val(.CellValue(lRow - 1, "IDTPSERVICO")) = 0 Then
               If mvarMe.ChkFLGAVALIACAO.Value Then
                  .CellValue(lRow - 1, "IDTPSERVICO") = 1
               Else
                  .CellValue(lRow - 1, "IDTPSERVICO") = 2
               End If
            End If
            If Val(.CellValue(lRow - 1, "IDSERVICOEVT")) = 0 Then Exit Sub
            If .ColVisible("IDAREA") Then
               If Val(.CellValue(lRow - 1, "IDTPTRATAMENTO")) = 0 Then Exit Sub
               If Val(.CellValue(lRow - 1, "IDAREA")) = 0 Then
                  If ClFLGAREA("k" & .CellValue(lRow - 1, "IDTPTRATAMENTO")) <> 0 Then
                     Exit Sub
                  End If
               End If
            End If
         End If
         If .ColVisible("IDAREA") Then
            Call IncluiGrdLinha
            If .RowCount = 2 And xVal(mvarMe.TxtIDCLIENTE.Text) = 0 Then
               If (mvarSys.Propriedades("TPSERV") = "1") Then
                  If xVal(.CellValue(1, "IDTPSERVICO")) = 0 Then
                     .CellValue(1, "IDTPSERVICO") = 1
                  End If
               End If
               If mvarSys.Propriedades("TPAREA") = "1" Then
                  If xVal(.CellValue(1, "IDTPTRATAMENTO")) = 0 Then
                     .CellValue(1, "IDTPTRATAMENTO") = 1
                     Call .SetCurCell(1, "IDAREA")
                  End If
               End If
            End If
         Else
            Call AdicionarTratamento
         End If
      End If
   End With
End Sub

Public Sub mvarMe_Load()
   If mvarSys Is Nothing Then Exit Sub
   mvarMe.Icon = mvarSys.MDI.Icon
   mvarisLoad = True
         
   '* Monta e Popula Tela
   Call MontaTela
   Call PopulaTela
      
   ModalFormsRunningCounter = ModalFormsRunningCounter + 1
   Screen.MousePointer = vbDefault
End Sub
Private Sub F_EventoNovo()
   Dim nScheduleNr As Integer
   Dim i As Integer
   Dim n As Variant
   Dim nAux As Integer
   Dim bMult As Integer
   
   If mvarMe.CmbSchedule.ListCount > 0 Then
      nScheduleNr = mvarCalControl.ActiveView.Selection.GroupIndex
      If nScheduleNr = 0 Then
         i = -1
         nAux = -1
         For Each n In mvarSys.Objetos("LOJAS")(CStr(mvarIDLOJA)).Salas
            i = i + 1
            If n.Selecionada Then
               bMult = bMult + 1
               nAux = i
            End If
         Next
         If bMult = 1 Then nScheduleNr = nAux
      End If
      mvarMe.CmbSchedule.ListIndex = nScheduleNr
   End If
        
    'Set mvarCalEvent = pCalEvent
    mvarCalEvent.StartTime = mvarBeginSelection
    mvarCalEvent.EndTime = mvarEndSelection
    
    SetStartEnd mvarBeginSelection, mvarEndSelection, mvarAllDay
    
    mvarMe.ChkAllDayEvent.Value = IIf(mvarAllDay, 1, 0)
    mvarMe.TxtSubject = "Novo"

   If mvarMe.CmbShowTimeAs.ListCount > 3 Then
      mvarMe.CmbShowTimeAs.ListIndex = IIf(mvarAllDay, 0, 2)
   End If
   If mvarMe.CmbLabel.ListCount > 0 Then
      mvarMe.CmbLabel.ListIndex = 0
   End If
    
   Call mvarMe_ChkReminderClick
   mvarMe.CmbReminder.Text = "15 minutos"
End Sub
Private Sub MontaTela()
   
   mvarMe.TxtTEL1.MaxLength = 20 'mvarSys.Propriedades("MaxLenTel")
   mvarMe.TxtTEL2.MaxLength = mvarMe.TxtTEL1.MaxLength
   mvarMe.TabEvento(1).Visible = True
   mvarMe.TabEvento(2).Visible = True
   mvarMe.TabEvento(0).Selected = True
   
   mvarMe.TxtIDCLIENTE.Text = ""
            
   Call MontaCmbTime
   Call MontaCmbLabel
   Call MontaCmbShowTimeAs(mvarMe.ChkMeeting.Value)
   Call MontaCmbSchedule
   Call CarregaComboTempo(mvarMe.CmbReminder, False)
   'mvarMe.TabEvento(1).Selected = True
   Call MontaGrdServico
   'mvarMe.TabEvento(0).Visible = True
   If mvarMe.TabEvento(1).Visible Then
      mvarMe.TabEvento(1).Selected = True
   End If
   mvarMe.CmbEndDate.Enabled = False
     
   mvarMe.ChkFLGAVALIACAO.Visible = False '(mvarSys.Propriedades("TPSERV") <> "0")
End Sub
Private Sub PopulaCmbCli()
   Dim Sql As String
   
   mvarMe.CmbNOME.Clear
   If bEventoNovo Then
      Sql = ""
      Sql = Sql & "Select Top 5 C.IDCLIENTE, C.NOME" & vbNewLine
      Sql = Sql & " , CASE WHEN C.TIMESTAMP > MAX(A.TIMESTAMP) THEN  C.TIMESTAMP  ELSE MAX(A.TIMESTAMP) END [TIMESTAMP]" & vbNewLine
      Sql = Sql & " , DateDiff(mi, CASE WHEN C.TIMESTAMP > MAX(A.TIMESTAMP) THEN  C.TIMESTAMP  ELSE MAX(A.TIMESTAMP) END, GetDate()) [DIF]" & vbNewLine
      Sql = Sql & " From OCLIENTE C" & vbNewLine
      Sql = Sql & " Join OATENDIMENTO A On A.IDLOJA=C.IDLOJA And A.IDCLIENTE=C.IDCLIENTE" & vbNewLine
      Sql = Sql & " Where C.IDLOJA = " & mvarIDLOJA & vbNewLine
      Sql = Sql & " GROUP BY C.IDCLIENTE, C.NOME, C.TIMESTAMP" & vbNewLine
      Sql = Sql & " Having DateDiff(mi,CASE WHEN C.TIMESTAMP > MAX(A.TIMESTAMP) THEN  C.TIMESTAMP  ELSE MAX(A.TIMESTAMP) END, GetDate()) <=5" & vbNewLine
      Sql = Sql & " And DateDiff(mi,CASE WHEN C.TIMESTAMP > MAX(A.TIMESTAMP) THEN  C.TIMESTAMP  ELSE MAX(A.TIMESTAMP) END, GetDate()) >=0" & vbNewLine
      Sql = Sql & " ORDER BY DIF"
      If mvarSys.xDb.AbreTabela(Sql) Then
         While Not mvarSys.xDb.RsAux.EOF
            mvarMe.CmbNOME.AddItem mvarSys.xDb.RsAux("NOME")
            mvarMe.CmbNOME.ItemData(mvarMe.CmbNOME.NewIndex) = mvarSys.xDb.RsAux("IDCLIENTE")
            mvarSys.xDb.RsAux.MoveNext
         Wend
      End If
   End If
   mvarMe.TxtNOME.Visible = (mvarMe.CmbNOME.ListCount = 0)
   mvarMe.CmbNOME.Visible = Not mvarMe.TxtNOME.Visible
End Sub
Private Sub MontaGrdServico()
   Dim Sql As String
   Dim MyRS As Object
   Dim i As Integer
   
   With mvarMe.GrdSERVICOEVT
      .ShowControlsInAllCells = True
      .Header.DragCols = False
      .Width = 5600
      .Left = 200
      
      '.Header.ImageList = mvarMe.IlsIcons
      '.ImageList = mvarMe.IlsIcons
            
      '************
      '* Montar Combos do Grid
      With .Combos
         For i = .Count To 1 Step -1
            .Remove i
         Next
         With .Add("IDTPSERVICO")
            Sql = Sql_OTPSERVICO(mvarIDLOJA)
            Sql = Sql & " And ATIVO=1"
            Sql = Sql & " Order By IDTPSERVICO"
            If mvarSys.xDb.AbreTabela(Sql, MyRS) Then
               While Not MyRS.EOF
                  .AddItem MyRS("DSCSERVICO").Value, MyRS("IDTPSERVICO").Value
                  MyRS.MoveNext
               Wend
            End If
            .AutoAdjustWidth
         End With
         
         Set ClFLGAREA = New Collection
         Set ClFLGAVALIACAO = New Collection
         Set ClFLGDISPARO = New Collection
         'Set ClFLGSESSAO = New Collection
         
         With .Add("IDTPTRATAMENTO")
            Sql = Sql_OTPTRATAMENTO(mvarIDLOJA)
            Sql = Sql & " And ATIVO=1"
            Sql = Sql & " Order By DSCTRATAMENTO"
            If mvarSys.xDb.AbreTabela(Sql, MyRS) Then
               While Not MyRS.EOF
                  .AddItem MyRS("DSCTRATAMENTO").Value, MyRS("IDTPTRATAMENTO").Value
                  ClFLGAREA.Add MyRS("FLGAREA").Value, "k" & CStr(MyRS("IDTPTRATAMENTO").Value)
                  ClFLGAVALIACAO.Add MyRS("FLGAVALIACAO").Value, "k" & CStr(MyRS("IDTPTRATAMENTO").Value)
                  ClFLGDISPARO.Add MyRS("FLGDISPARO").Value, "k" & CStr(MyRS("IDTPTRATAMENTO").Value)
                  'ClFLGSESSAO.Add MyRs("FLGSESSAO").Value, "k" & CStr(MyRs("IDTPTRATAMENTO").Value)
                  MyRS.MoveNext
               Wend
            End If
            .MaxHeightInItems = 10
            .AutoAdjustWidth
         End With
         With .Add("IDAREA")
            Sql = Sql_OAREA(mvarIDLOJA)
            Sql = Sql & " And ATIVO in " & IIf(mvarSys.Propriedades("TPAREA") = "0", "(0,1)", "(1)")
            Sql = Sql & " Order By DSCAREA"
            If mvarSys.xDb.AbreTabela(Sql, MyRS) Then
               While Not MyRS.EOF
                  .AddItem MyRS("DSCAREA").Value, MyRS("IDAREA").Value
                  MyRS.MoveNext
               Wend
            End If
            .MaxHeightInItems = 10
            .AutoAdjustWidth
         End With
      End With
      '************
      '* Montar colunas do Grid
      For i = .ColCount To 1 Step -1
         .RemoveCol i
      Next
      With .AddCol(sKey:="IDSERVICOEVT", sHeader:="IDSERVICOEVT", lWidth:=85, bVisible:=False)
         .eType = igCellText
         .sCtrlKey = "IDSERVICOEVT"
      End With
      With .AddCol(sKey:="IDTPSERVICO", sHeader:="Serviço", lWidth:=90, bVisible:=True)
         .eType = igCellCombo
         .eTypeFlags = igComboBtnFlat
         .sCtrlKey = "IDTPSERVICO"
      End With
      With .AddCol(sKey:="IDTPTRATAMENTO", sHeader:="Tratamento", lWidth:=120, bVisible:=True)
         .eType = igCellCombo
         .eTypeFlags = igComboBtnFlat
         .sCtrlKey = "IDTPTRATAMENTO"
      End With
      With .AddCol(sKey:="IDAREA", sHeader:="Área", lWidth:=140, bVisible:=True)
         .eType = igCellCombo
         .eTypeFlags = igComboBtnFlat
         .sCtrlKey = "IDAREA"
      End With
      '************
      '* Configurar Estrutura do Serviço
      If mvarSys.Propriedades("TPSERV") = "0" Then
         If mvarSys.Propriedades("TPAREA") = "0" Then
            .ColHeaderText("IDTPSERVICO") = "Serviços"
            .ColHeaderText("IDTPTRATAMENTO") = ""
         End If
         .ColVisible("IDTPSERVICO") = True
         '.ColWidth("IDTPTRATAMENTO") = .ColWidth("IDTPTRATAMENTO") + .ColWidth("IDTPSERVICO")
      End If
      If mvarSys.Propriedades("TPAREA") = "0" Then
         .ColVisible("IDAREA") = False
         .ColWidth("IDTPTRATAMENTO") = .ColWidth("IDTPTRATAMENTO") + .ColWidth("IDAREA")
      End If
      .Combos("IDTPTRATAMENTO").Width = .ColWidth("IDTPTRATAMENTO")
      .ColVisible("IDTPTRATAMENTO") = (mvarSys.Propriedades("TPTRAT") = 1 Or mvarSys.Propriedades("TPTRAT") = "")
      
      For i = 1 To .ColCount
         .ColTag(i) = ""
      Next
      
      .ColTag("IDSERVICOEVT") = SetTag(.ColTag("IDSERVICOEVT"), "EDIT", False)
      .ColTag("IDTPSERVICO") = SetTag(.ColTag("IDTPSERVICO"), "EDIT", (.Combos("IDTPSERVICO").ListCount > 2) And (mvarSys.Propriedades("TPSERV") <> "0"))
      .ColTag("IDTPTRATAMENTO") = SetTag(.ColTag("IDTPTRATAMENTO"), "EDIT", True)
      .ColTag("IDAREA") = SetTag(.ColTag("IDAREA"), "EDIT", (.Combos("IDAREA").ListCount > 2) And (mvarSys.Propriedades("TPAREA") <> "0"))
            
      .ColTag("IDSERVICOEVT") = SetTag(.ColTag("IDSERVICOEVT"), "EDITOPT", igTextEditNumberOnly)
      
      Call mvarMe_ChkFLGAVALIACAOClick
      
      IncluiGrdLinha
   End With
End Sub
Public Sub IncluiGrdLinha()
   Dim lRow As Long
   Dim i As Integer
   Dim bPernaD As Boolean
   Dim bPernaE As Boolean
   Dim bBracoD As Boolean
   Dim bBracoE As Boolean
   
   On Error Resume Next
   If mvarMe.GrdSERVICOEVT Is Nothing Then Exit Sub
   
   With mvarMe.GrdSERVICOEVT
      lRow = .RowCount
      If lRow > 0 Then .RemoveRow lRow
      If lRow > 1 Then
         If GetTag(.ColTag("IDTPSERVICO"), "EDIT", True) = False Then
            .CellForeColor(lRow - 1, "IDTPSERVICO") = vbGrayText
         Else
            .CellForeColor(lRow - 1, "IDTPSERVICO") = vbBlack
         End If
      End If
      .CellForeColor(lRow - 1, "IDTPTRATAMENTO") = vbBlack
      
      If lRow > 0 Or mvarIDEVENTO > 0 Then
         .AddRow
         lRow = .RowCount
   
         .CellType(lRow, "IDSERVICOEVT") = igCellText
         .CellValue(lRow, "IDSERVICOEVT") = lRow
         .CellForeColor(lRow, "IDSERVICOEVT") = vbBlack
         
         .CellType(lRow, "IDTPSERVICO") = igCellCombo
         .CellValue(lRow, "IDTPSERVICO") = IIf(lRow = 1, 1, .CellValue(lRow - 1, "IDTPSERVICO"))
         If GetTag(.ColTag("IDTPSERVICO"), "EDIT", True) = False Then
            .CellForeColor(lRow, "IDTPSERVICO") = vbGrayText
         Else
            .CellForeColor(lRow, "IDTPSERVICO") = vbBlack
         End If
   
         .CellType(lRow, "IDTPTRATAMENTO") = igCellCombo
         If lRow > 1 Then
            If ClFLGAREA("k" & mvarMe.GrdSERVICOEVT.CellValue(lRow - 1, "IDTPTRATAMENTO")) = 1 Then
               .CellValue(lRow, "IDTPTRATAMENTO") = IIf(lRow = 1, 1, .CellValue(lRow - 1, "IDTPTRATAMENTO"))
            End If
         End If
         .CellForeColor(lRow, "IDTPTRATAMENTO") = vbBlack
   
         .CellType(lRow, "IDAREA") = igCellCombo
         .CellValue(lRow, "IDAREA") = ""
         .CellForeColor(lRow, "IDAREA") = vbBlack
         If lRow > 1 And .ColVisible("IDAREA") Then
            bPernaD = False
            bPernaE = False
            bBracoD = False
            bBracoE = False
            For i = 1 To .RowCount - 1
               bPernaD = bPernaD Or (.CellValue(i, "IDAREA") = 16)
               bPernaE = bPernaE Or (.CellValue(i, "IDAREA") = 17)
               bBracoD = bBracoD Or (.CellValue(i, "IDAREA") = 4)
               bBracoE = bBracoE Or (.CellValue(i, "IDAREA") = 5)
            Next
            If .CellValue(lRow - 1, "IDAREA") = "16" And Not bPernaE Then
               .CellValue(lRow, "IDAREA") = 17
               Call mvarMe_GrdSERVICOEVTAfterCommitEdit(lRow, .ColIndex("IDAREA"))
            ElseIf .CellValue(lRow - 1, "IDAREA") = "17" And Not bPernaD Then
               .CellValue(lRow, "IDAREA") = 16
               Call mvarMe_GrdSERVICOEVTAfterCommitEdit(lRow, .ColIndex("IDAREA"))
            ElseIf .CellValue(lRow - 1, "IDAREA") = "4" And Not bBracoE Then
               .CellValue(lRow, "IDAREA") = 5
               Call mvarMe_GrdSERVICOEVTAfterCommitEdit(lRow, .ColIndex("IDAREA"))
            ElseIf .CellValue(lRow - 1, "IDAREA") = "5" And Not bBracoD Then
               .CellValue(lRow, "IDAREA") = 4
               Call mvarMe_GrdSERVICOEVTAfterCommitEdit(lRow, .ColIndex("IDAREA"))
            End If
         End If
         If GetTag(.ColTag("IDTPSERVICO"), "EDIT", True) = False Then
            .CellForeColor(lRow, "IDTPSERVICO") = vbGrayText
         Else
            .CellForeColor(lRow, "IDTPSERVICO") = vbBlack
         End If

      End If

      .AddRow
      lRow = 0
      For i = 1 To .ColCount
         .CellType(.RowCount, i) = igCellText
         .CellForeColor(.RowCount, i) = vbGrayText
      Next
      lRow = .RowCount
      If lRow > 0 Then
         For i = 1 To .ColCount
            If .ColVisible(i) Then
               .CellValue(.RowCount, i) = "Clique para incluir nova linha."
               .CellForeColor(.RowCount, i) = vbGrayText
               .CellTextFlags(.RowCount, i) = igTextCenter
               .CellTextFlags(.RowCount, i) = igTextNoClip Or igTextCenter
               i = .ColCount
            End If
         Next
      End If
      .RowMode = (.RowCount = 1)
      If .RowCount > 1 Then
         If .ColVisible("IDAREA") Then
            .SetCurCell lRow, "IDAREA"
         Else
            .SetCurCell lRow, "IDTPTRATAMENTO"
         End If
      End If
      If .Visible Then
         .SetFocus
      End If
      .Redraw = True
   End With
End Sub

'Default Background Color  Default Name Default LabelID
'RGB(255,255,255)          None              0
'RGB(255,148,132)          Important         1
'RGB(132,156,231)          Business          2
'RGB(165,222, 99)          Personal          3
'RGB(231,231,214)          Vacation          4
'RGB(255,181,115)          Must Attend       5
'RGB(132,239,247)          Travel Required   6
'RGB(214,206,132)          Needs Preparation 7
'RGB(198,165,247)          Birthday          8
'RGB(165,206,198)          Anniversary       9
'RGB(255,231,115)          Phone Call        10
Private Sub MontaCmbLabel()
   Dim pLabel As CalendarEventLabel
   Dim i As Integer
   
   With mvarMe
'      If False Then
'         With mvarCalControl.DataProvider.LabelList
'            .RemoveAll
'            .AddLabel 0, RGB(255, 255, 255), " "
'            .AddLabel 1, RGB(255, 231, 115), "Teste"
'            .AddLabel 2, RGB(165, 222, 99), "Depilação"
'            .AddLabel 3, RGB(165, 206, 198), "Rejuvenescimento"
'            .AddLabel 4, RGB(214, 206, 132), "Acne"
'            .AddLabel 5, RGB(255, 148, 132), "Importante"
'         End With
'      End If
      If mvarCalControl Is Nothing Then
         .CmbLabel.Clear
         .CmbLabel.AddItem " "
         .CmbLabel.ItemData(.CmbLabel.NewIndex) = 0
         .CmbLabel.AddItem "All Day "
         .CmbLabel.ItemData(.CmbLabel.NewIndex) = 6
      Else
         .CmbLabel.Clear
         For i = 0 To mvarCalControl.DataProvider.LabelList.Count - 1
            Set pLabel = mvarCalControl.DataProvider.LabelList.Item(i)
            .CmbLabel.AddItem pLabel.Name
            .CmbLabel.ItemData(.CmbLabel.NewIndex) = pLabel.LabelID
         Next
         
'         Set pLabel = mvarCalControl.DataProvider.LabelList.Item(0)
'         .CmbLabel.Clear
'         .CmbLabel.AddItem pLabel.Name
'         .CmbLabel.ItemData(.CmbLabel.NewIndex) = pLabel.LabelID
'
'         For i = 6 To mvarCalControl.DataProvider.LabelList.Count - 1
'            Set pLabel = mvarCalControl.DataProvider.LabelList.Item(i)
'            .CmbLabel.AddItem pLabel.Name
'            .CmbLabel.ItemData(.CmbLabel.NewIndex) = pLabel.LabelID
'         Next
      End If
   End With
End Sub
Private Sub MontaCmbShowTimeAs(bSessao As Boolean)
   With mvarMe.CmbShowTimeAs
      .Clear
      .AddItem IIf(bSessao, "Avaliação", "Disponível"), 0
      .AddItem IIf(bSessao, "Manutenção", "Provisório"), 1
      .AddItem IIf(bSessao, "Tratamento", "Ocupado"), 2
      .AddItem "Outros", 3
      .ListIndex = IIf(bSessao, 0, 2)
   End With
End Sub
Private Sub MontaCmbSchedule()
   Dim vSala As Variant
   
   Call VerificaLoja
   With mvarMe
      If mvarCalControl Is Nothing Then
         .CmbSchedule.Enabled = False
      Else
         .CmbSchedule.Clear
         .CmbSchedule.Enabled = True
         If mvarSys.ExisteObjeto("LOJAS") Then
            If ExisteItem(mvarSys.Objetos("LOJAS"), CStr(mvarIDLOJA)) Then
               For Each vSala In mvarSys.Objetos("LOJAS")(CStr(mvarIDLOJA)).Salas
                  .CmbSchedule.AddItem vSala.NMSALA
                  .CmbSchedule.ItemData(.CmbSchedule.NewIndex) = vSala.IDSALA
               Next
            End If
         End If
         If .CmbSchedule.ListCount = 0 Then
            .CmbSchedule.AddItem IIf(mvarSys.Propriedades("ExibeSala") = 0, "", "Sala ") & "01"
            .CmbSchedule.ItemData(.CmbSchedule.NewIndex) = 1
         End If
      End If
   End With
End Sub
Private Sub mvarMe_Resize()
   With mvarMe.ctrlColor
      .Appearance = 1
      .BackStyle = 1
      .Visible = True
      .BorderStyle = 0
      .Left = 5
      .Top = 0
      .Width = mvarMe.ScaleWidth
      .Height = 120
      .Refresh
   End With
End Sub

Private Sub mvarMe_TxtIDCLIENTEChange()
   Dim Sql As String
   Dim MyRS As Object
      
   Sql = Sql_OCLIENTE(mvarIDLOJA, pIDCLIENTE:=xVal(mvarMe.TxtIDCLIENTE.Text))
   If mvarSys.xDb.AbreTabela(Sql, MyRS) Then
      mvarMe.TxtIDCLIENTE.Text = MyRS("IDCLIENTE")
      mvarMe.TxtNOME.Text = MyRS("NOME")
      mvarMe.TxtTEL1.Text = MyRS("TEL1") & ""
      mvarMe.TxtTEL2.Text = MyRS("TEL2") & ""
      mvarMe.TxtEMAIL.Text = MyRS("EMAIL") & ""
   Else
      mvarMe.TxtIDCLIENTE.Text = ""
      mvarMe.TxtNOME.Text = ""
      mvarMe.TxtTEL1.Text = ""
      mvarMe.TxtTEL2.Text = ""
      mvarMe.TxtEMAIL.Text = ""
   End If
End Sub

Private Sub mvarMe_TxtNOMEChange()
   Call MontaSubject
   mvarMe.CmbNOME.Text = mvarMe.TxtNOME.Text
   mvarMe.CmbNOME.SelStart = Len(mvarMe.CmbNOME.Text)
   mvarMe.CmbNOME.SelLength = 0
End Sub
Private Sub mvarMe_TxtTEL1Change()
   Call MontaSubject
End Sub
Private Sub mvarMe_TxtTEL1LostFocus()
   mvarMe.TxtTEL1.Text = Trim(Replace(mvarMe.TxtTEL1.Text, " ", ""))
   mvarMe.TxtTEL1.Text = Trim(Replace(mvarMe.TxtTEL1.Text, ".", ""))
End Sub
Private Sub mvarMe_TxtTEL2Change()
   Call MontaSubject
End Sub
Private Sub mvarMe_TxtTEL2LostFocus()
   mvarMe.TxtTEL2.Text = Trim(Replace(mvarMe.TxtTEL2.Text, " ", ""))
   mvarMe.TxtTEL2.Text = Trim(Replace(mvarMe.TxtTEL2.Text, ".", ""))
End Sub
Private Sub mvarMe_Unload(Cancel As Integer)
   ModalFormsRunningCounter = ModalFormsRunningCounter - 1
   mvarisLoad = False
   Set mvarCalEvent = Nothing
   'Cancel = True
   On Error Resume Next
   'mvarMe.Hide
   Set mvarMe = Nothing
End Sub
Public Sub PopulaTela()
   'mvarMe.TxtNOME.Text = ""
   'mvarMe_ChkMeetingClick
   If bEventoNovo Then
     Call F_EventoNovo
     Call PopulaCmbCli
   Else
      If Not mvarCalEvent Is Nothing Then
         If mvarCalEvent.RecurrenceState = xtpCalendarRecurrenceOccurrence Then
            mvarCalEvent.MakeAsRException
         End If
         mvarIDEVENTO = mvarCalEvent.Id
         Call PopulaTela_Evento
      ElseIf mvarIDEVENTO <> 0 Then
         Call PopulaTela_Compromisso
      End If
   End If
End Sub
Private Sub PopulaTela_Compromisso()
   Dim i As Long
   Dim bDatesVisible As Boolean
   Dim TBEvento As Object
   
   Set TBEvento = CriarObjeto("BANCO_3R.TB_OEVENTOAGENDA")
   Set TBEvento.xDb = mvarSys.xDb
   
   If TBEvento.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDEVENTO:=mvarIDEVENTO) Then
      With mvarMe
         ' Restore Event base properties
         .TxtSubject = TBEvento.Subject
         .txtBody = TBEvento.Body
         .TxtLocation = TBEvento.Location
         
         .ChkAllDayEvent = IIf(TBEvento.isAllDayEvent, 1, 0)
         
         Call LocalizarCombo(.CmbLabel, TBEvento.LabelID, True, True)
         Call LocalizarCombo(.CmbSchedule, TBEvento.ScheduleID, True, True)
         
         ' Restore other properties
         .CmbShowTimeAs.ListIndex = TBEvento.BusyStatus
         
         .ChkPrivate.Value = IIf(TBEvento.isPrivate, 1, 0)
         .ChkMeeting.Value = IIf(TBEvento.ISMEETING, 1, 0)
         
         SetStartEnd TBEvento.StartDateTime, TBEvento.EndDateTime, TBEvento.isAllDayEvent
         
         If xVal(TBEvento.IDEVENTO) <> 0 Then
            .Caption = "Compromisso: Nº" & StrZero(TBEvento.IDEVENTO, 6)
         ElseIf (TBEvento.Subject <> "") Then
            .Caption = TBEvento.Subject & " - Compromisso"
         End If
         
   
         bDatesVisible = TBEvento.RecurrenceState <> xtpCalendarRecurrenceMaster
            
         .lblStartTime.Visible = bDatesVisible
         .lblEndTime.Visible = bDatesVisible
         .CmbStartDate.Visible = bDatesVisible
         .CmbStartTime.Visible = bDatesVisible
         .CmbEndDate.Visible = bDatesVisible
         .CmbEndTime.Visible = bDatesVisible
         .ChkAllDayEvent.Visible = bDatesVisible
             
         If bDatesVisible Then
            Call mvarMe_ChkAllDayEventClick
         End If
         
         If TBEvento.isReminder Then
            .ChkReminder.Value = Checked
            .CmbReminder.Text = CalcStandardDurations_0m_2wString(TBEvento.ReminderMinutesBeforeStart)
         End If
         .CmdOk.Enabled = False
      End With
      Call ExibeObsCli(TBEvento.IDCLIENTE)

      Dim MyNG As New NG_Calendario
      With MyNG
         Set .Sys = mvarSys
         Set .TlEvt = Me
         .IDAGENDA = TBEvento.IDAGENDA
         '.IDCLIENTE = TBEvento.IDCLIENTE
         '.IDCOLIGADA = TBEvento.IDCOLIGADA
         .IDLOJA = TBEvento.IDLOJA
         Call .MPopulaAtributos(TBEvento)
      End With
   End If
   Set TBEvento = Nothing
   Set MyNG = Nothing
End Sub
Private Sub PopulaTela_Evento()
   Dim i As Long
   Dim bDatesVisible As Boolean
   
   If mvarCalEvent Is Nothing Then
      Exit Sub
      Set mvarCalEvent = mvarCalControl.DataProvider.CreateEvent
   End If
   
   mvarCalControl.DataProvider.ClearCache
   If mvarCalEvent.Id > 0 Then
      Set mvarCalEvent = mvarCalControl.DataProvider.GetEvent(mvarCalEvent.Id)
   Else
       mvarCalControl.DataProvider.GetEvent (mvarCalEvent.Id)
   End If
   
   If mvarCalEvent Is Nothing Then Exit Sub
   
   Call PopulaCmbCli
   If mvarCalEvent.Id < 0 Then
      mvarMe.TxtNOME.Visible = True
      mvarMe.CmbNOME.Visible = False
   End If
   
   With mvarMe
      ' Restore Event base properties
      .TxtSubject.Text = mvarCalEvent.Subject
      .txtBody.Text = mvarCalEvent.Body
      .TxtLocation.Text = mvarCalEvent.Location
      Call LocalizarCombo(.CmbSchedule, mvarCalEvent.Location)
      If .CmbSchedule.ListCount <= 0 Then
         .CmbSchedule.AddItem .TxtLocation.Text, 0
         Call LocalizarCombo(.CmbSchedule, mvarCalEvent.Location)
      End If
      
      .ChkAllDayEvent = IIf(mvarCalEvent.AllDayEvent, 1, 0)
      
      Call LocalizarCombo(.CmbLabel, mvarCalEvent.Label, True, True)
      Call LocalizarCombo(.CmbSchedule, xVal(mvarCalEvent.CustomProperties("IDSALA")), True, True)
      
      ' Restore other properties
      .CmbShowTimeAs.ListIndex = mvarCalEvent.BusyStatus
      
      .ChkPrivate.Value = IIf(mvarCalEvent.PrivateFlag, 1, 0)
      .ChkMeeting.Value = IIf(mvarCalEvent.MeetingFlag, 1, 0)
      
      If mvarCalEvent.Id > 0 Then
         .CmbStartTime.Text = Format(mvarCalEvent.StartTime, "hh:mm")
         .CmbEndTime.Text = Format(mvarCalEvent.EndTime, "hh:mm")
      End If
      SetStartEnd mvarCalEvent.StartTime, mvarCalEvent.EndTime, mvarCalEvent.AllDayEvent
         
      
      
      If xVal(mvarCalEvent.Id) <> 0 Then
         .Caption = "Compromisso: Nº" & StrZero(mvarCalEvent.Id, 6)
      ElseIf (mvarCalEvent.Subject <> "") Then
         .Caption = mvarCalEvent.Subject & " - Compromisso"
      End If

      bDatesVisible = mvarCalEvent.RecurrenceState <> xtpCalendarRecurrenceMaster
         
      .lblStartTime.Visible = bDatesVisible
      .lblEndTime.Visible = bDatesVisible
      .CmbStartDate.Visible = bDatesVisible
      .CmbStartTime.Visible = bDatesVisible
      .CmbEndDate.Visible = bDatesVisible
      .CmbEndTime.Visible = bDatesVisible
      .ChkAllDayEvent.Visible = bDatesVisible
          
      If bDatesVisible Then
         Call mvarMe_ChkAllDayEventClick
      End If
      
      If mvarCalEvent.Reminder Then
         .ChkReminder.Value = Checked
         .CmbReminder.Text = CalcStandardDurations_0m_2wString(mvarCalEvent.ReminderMinutesBeforeStart)
      End If
   End With
   Call ExibeObsCli(Val(mvarCalEvent.CustomProperties("IDCLIENTE")))
End Sub
Private Sub ExibeObsCli(pIDCLIENTE As Long)
   Dim Sql As String
   Dim MyRS As Object
   Dim sAux As String
   Dim sAux2 As String
   Dim nDias As Double
   Dim nQtd As Integer
   Dim nIDLOJA As Integer
      
   nIDLOJA = mvarIDLOJA
   
   Sql = "Select Max(A.DTATEND) as [DTATEND], S.IDTPSERVICO "
   Sql = Sql & " From OATENDIMENTO  A "
   Sql = Sql & " Join OSESSAO S ON A.IDLOJA=S.IDLOJA And A.IDATENDIMENTO=S.IDATENDIMENTO "
   'Sql = Sql & " AND S.IDTPSERVICO<>1"
   Sql = Sql & " Where A.IDCLIENTE = " & SqlNum(pIDCLIENTE)
   Sql = Sql & " And A.IDLOJA =" & SqlNum(mvarIDLOJA)
   Sql = Sql & " And A.DTATEND <" & SqlDate(mvarMe.CmbStartDate.Value)
   Sql = Sql & " Group By A.IDCLIENTE, S.IDTPSERVICO"
   Sql = Sql & " Order By S.IDTPSERVICO"
   
   mvarMe.LblObsCli.Caption = ""
   With mvarSys.xDb
      If .AbreTabela(Sql, MyRS) Then
         If MyRS("IDTPSERVICO") = 1 Then
            nDias = DateDiff("d", MyRS("DTATEND"), mvarMe.CmbStartDate.Value)
            sAux = "Avaliação: "
            sAux = sAux & Format(MyRS("DTATEND"), "dd/mm")
            sAux = sAux & " (" & CStr(nDias)
            sAux = sAux & " dias)"

            nDias = 21
            MyRS.MoveNext
         End If
         If Not MyRS.EOF Then
            nDias = DateDiff("d", MyRS("DTATEND"), mvarMe.CmbStartDate.Value)
            
            sAux = "Sessão: "
            sAux = sAux & Format(MyRS("DTATEND"), "dd/mm")
            sAux = sAux & " (" & CStr(nDias)
            sAux = sAux & " dias)"
         
         End If
      End If
   End With
   
   Sql = "Select V.IDCLIENTE, V.DTVENDA, I.IDPROD, I.QTDVENDA"
   Sql = Sql & " From CVENDA V"
   Sql = Sql & " Join CITENSVENDA I On V.IDLOJA=I.IDLOJA And V.IDVENDA=I.IDVENDA"
   Sql = Sql & " Join SPRODUTO P On P.IDLOJA=I.IDLOJA And P.IDPROD=I.IDPROD And P.ESERVICO=0"
   Sql = Sql & " Where V.IDCLIENTE = " & SqlNum(pIDCLIENTE)
   Sql = Sql & " And V.IDLOJA =" & SqlNum(nIDLOJA)
   Sql = Sql & " And Not V.DTVENDA is Null"
   Sql = Sql & " Order By V.DTVENDA"
   
   nQtd = 0
   nDias = 0
   With mvarSys.xDb
      If .AbreTabela(Sql, MyRS) Then
         While Not MyRS.EOF
            nQtd = nQtd + MyRS("QTDVENDA")
            MyRS.MoveNext
         Wend
         MyRS.MoveLast
         nDias = DateDiff("d", MyRS("DTVENDA"), mvarSys.xDb.Sysdate)
         
         sAux2 = "Creme: " & nQtd
'         sAux2 = sAux2 & " (" + Format(MyRs("DTVENDA"), "dd/mm")
'         sAux2 = sAux2 & " - "
'         If nDias <= 30 Then
'            sAux2 = sAux2 & CStr(nDias) & " dias)"
'         Else
'            sAux2 = sAux2 & CInt(nDias / 30) & " " & IIf(CInt(nDias / 30) <= 1, "mês)", "meses)")
'         End If
      Else
         sAux2 = "Creme: 0"
      End If
      If nQtd = 0 And Trim(sAux2) <> "" Then
         mvarMe.LblObsCli.ForeColor = 192
      Else
         mvarMe.LblObsCli.ForeColor = &H404040
      End If
      If Trim(sAux) = "" Then
         mvarMe.LblObsCli.Caption = Trim(sAux2)
      Else
         mvarMe.LblObsCli.Caption = Trim(sAux) & IIf(Trim(sAux2) = "", "", " / ") & Trim(sAux2)
      End If
   End With
End Sub
Public Sub NewEvent(pCalEvent As CalendarEvent, pBeginSelection As Date, pEndSelection As Date, pAllDay As Boolean)
   Dim nScheduleNr As Integer
   Dim i As Integer
   Dim n As Variant
   Dim nAux As Integer
   Dim bMult As Integer
   
   
   If Not VerificaLoja Then Exit Sub
   
   bEventoNovo = True
   
   Set mvarCalEvent = pCalEvent
   mvarBeginSelection = pBeginSelection
   mvarEndSelection = pEndSelection
   mvarAllDay = pAllDay
   mvarMe.Show vbModal
End Sub
Private Function VerificaLoja() As Boolean
   Dim TlCal As TL_CALENDARIO
   Dim bOk   As Boolean
   
   If mvarSys.ExisteObjeto("LOJAS") Then
      If Not ExisteItem(mvarSys.Objetos("LOJAS"), CStr(mvarIDLOJA)) Then
         Set TlCal = New TL_CALENDARIO
         Set TlCal.Sys = mvarSys
         Call TlCal.MontarUnidades
         Set TlCal = Nothing
      End If
   End If
   If mvarSys.ExisteObjeto("LOJAS") Then
      If ExisteItem(mvarSys.Objetos("LOJAS"), CStr(mvarIDLOJA)) Then
         bOk = True
      End If
   End If
   If Not bOk Then
      Call ExibirAviso("Sistema não achou estrutura Loja/Sala", "Ambiente")
   End If
   VerificaLoja = bOk
End Function
Private Function VerificaDatas() As Boolean
   With mvarMe
      VerificaDatas = False
      If (Not IsDate(.CmbStartDate.Value)) Then
          .CmbStartDate.SetFocus
          Exit Function
      End If
      If (Not IsDate(.CmbStartTime.Text)) And .CmbStartTime.Enabled And .CmbStartTime.Visible Then
          .CmbStartTime.SetFocus
          Exit Function
      End If
      If (Not IsDate(.CmbEndDate.Value)) And mvarMe.CmbEndDate.Enabled Then
          .CmbEndDate.SetFocus
          Exit Function
      End If
      If (Not IsDate(.CmbEndTime.Text)) And .CmbEndTime.Enabled And .CmbEndTime.Visible Then
          .CmbEndTime.SetFocus
          Exit Function
      End If
      VerificaDatas = True
   End With
End Function
Private Function VerificaSessao() As Boolean
   Dim bOk As Boolean
   If mvarMe.ChkMeeting.Value Then
      If Trim(mvarMe.TxtTEL1.Text) = "" And Trim(mvarMe.TxtTEL2.Text) = "" And Trim(mvarMe.TxtEMAIL.Text) = "" Then
         Call ExibirInformacao("Informe ao menos um contato do cliente.")
         If Trim(mvarMe.TxtTEL1.Text) = "" Then
            If mvarMe.TxtTEL1.Enabled And mvarMe.TxtTEL1.Visible Then
               mvarMe.TxtTEL1.SetFocus
            End If
         End If
         Exit Function
      End If
      If Trim(mvarMe.TxtNOME.Text) = "" Then
         Call ExibirInformacao("Informe o nome do Cliente.")
         mvarMe.TxtNOME.SetFocus
         Exit Function
      End If
      If mvarMe.GrdSERVICOEVT.RowCount <= 1 Then
         Call ExibirInformacao("Serviço inválido.")
         mvarMe.GrdSERVICOEVT.SetFocus
         Exit Function
      Else
         bOk = (Val(mvarMe.GrdSERVICOEVT.CellValue(1, "IDTPSERVICO")) = 0)
         bOk = bOk Or (Val(mvarMe.GrdSERVICOEVT.CellValue(1, "IDTPTRATAMENTO")) = 0)
         bOk = bOk Or (Val(mvarMe.GrdSERVICOEVT.CellValue(1, "IDAREA")) = 0)
         If bOk Then
            If Not mvarMe.GrdSERVICOEVT.ColVisible("IDTPSERVICO") Then
               mvarMe.GrdSERVICOEVT.CellValue(1, "IDTPSERVICO") = 2
            End If
            If Not mvarMe.GrdSERVICOEVT.ColVisible("IDAREA") Then
               mvarMe.GrdSERVICOEVT.CellValue(1, "IDAREA") = 0
            End If
            
            If Val(mvarMe.GrdSERVICOEVT.CellValue(1, "IDTPSERVICO")) = 0 Then
               Call ExibirInformacao("Serviço inválida.")
               bOk = False
            End If
            If Val(mvarMe.GrdSERVICOEVT.CellValue(1, "IDTPTRATAMENTO")) = 0 Then
               Call ExibirInformacao("Tratamento inválido.")
               bOk = False
            End If

            If Val(mvarMe.GrdSERVICOEVT.CellValue(1, "IDAREA")) = 0 And mvarMe.GrdSERVICOEVT.ColVisible("IDAREA") Then
               If mvarMe.GrdSERVICOEVT.CellValue(1, "IDAREA") = "" Then
                  Call ExibirInformacao("Área inválida.")
                  bOk = False
               ElseIf ClFLGAREA("k" & mvarMe.GrdSERVICOEVT.CellValue(1, "IDTPTRATAMENTO")) = 1 Then
                  Call ExibirInformacao("Área inválida.")
                  bOk = False
               End If
            End If
            If Not bOk Then
               mvarMe.GrdSERVICOEVT.SetFocus
               Exit Function
            End If
         End If
      End If
'      If Val(GetTag(mvarMe.TxtNOME, "IDCLIENTE", "0")) <= 0 Then
'         Call ExibirInformacao("Código do cliente inválido. (" & GetTag(mvarMe.TxtNOME, "IDCLIENTE", "0") & ")")
'         mvarMe.CmdIDCLIENTE.SetFocus
'         Exit Function
'      End If
   End If
   VerificaSessao = True
End Function
Private Sub PopulaClasse_Evento()
   Dim StartTime As Date
   Dim EndTime As Date
   Dim nIdSala As Integer
   Dim oLoja   As Object
   
   StartTime = DateFromString(mvarMe.CmbStartDate.Value, mvarMe.CmbStartTime.Text)
   EndTime = DateFromString(mvarMe.CmbEndDate.Value, mvarMe.CmbEndTime.Text)
   
   If mvarMe.ChkAllDayEvent.Value = 1 Then
       If DateDiff("s", TimeValue(EndTime), 0) = 0 Then
           EndTime = EndTime + 1
       End If
   End If
   
   If mvarCalEvent.RecurrenceState <> xtpCalendarRecurrenceMaster Then
       mvarCalEvent.StartTime = StartTime
       mvarCalEvent.EndTime = EndTime
   End If
   
   mvarCalEvent.Subject = mvarMe.TxtSubject.Text
   If mvarMe.CmbSchedule.Visible Then
      mvarCalEvent.Location = mvarMe.CmbSchedule.Text
   Else
      mvarCalEvent.Location = mvarMe.TxtLocation.Text
   End If
   mvarCalEvent.Body = mvarMe.txtBody
   mvarCalEvent.AllDayEvent = (mvarMe.ChkAllDayEvent.Value = 1)
   
   mvarCalEvent.Label = 0
   If mvarMe.CmbLabel.ListIndex > 0 Then
      mvarCalEvent.Label = mvarMe.CmbLabel.ItemData(mvarMe.CmbLabel.ListIndex)
   ElseIf mvarMe.ChkMeeting.Value Then
      mvarCalEvent.Label = DefineLabelID(mvarMe.GrdSERVICOEVT.CellValue(1, "IDTPTRATAMENTO"), mvarMe.GrdSERVICOEVT.CellValue(1, "IDTPSERVICO"), mvarMe.ChkFLGCANCELADO.Value)
   End If
   

   mvarCalEvent.BusyStatus = mvarMe.CmbShowTimeAs.ListIndex
   If mvarMe.CmbSchedule.ListIndex >= 0 And mvarMe.CmbSchedule.ListIndex < mvarMe.CmbSchedule.ListCount Then
      nIdSala = mvarMe.CmbSchedule.ItemData(mvarMe.CmbSchedule.ListIndex)
      Set oLoja = mvarSys.Objetos("LOJAS")(CStr(mvarIDLOJA))
      If Not ExisteItem(oLoja.Salas, CStr(nIdSala)) Then
         Dim TlCal As TL_CALENDARIO
         Set TlCal = New TL_CALENDARIO
         Set TlCal.Sys = mvarSys
         Call TlCal.MontarUnidades
         Set TlCal = Nothing
         If Not ExisteItem(oLoja.Salas, CStr(nIdSala)) Then
            Call ExibirAviso("Verifique o carregamento de Salas")
            Exit Sub
         End If
      End If
      mvarCalEvent.ScheduleID = oLoja.Salas(CStr(nIdSala)).ScheduleID
      mvarCalEvent.CustomProperties("IDSALA") = nIdSala
   End If
   If mvarCalEvent.ScheduleID = 0 Then mvarCalEvent.ScheduleID = mvarIDLOJA * 1000 + 1
   If xVal(mvarCalEvent.CustomProperties("IDSALA")) = 0 Then mvarCalEvent.CustomProperties("IDSALA") = 1
     
   mvarCalEvent.PrivateFlag = (mvarMe.ChkPrivate.Value = 1)
   mvarCalEvent.MeetingFlag = (mvarMe.ChkMeeting.Value = 1)
   
   If Not mvarMe.ChkReminder.Value = mvarCalEvent.Reminder Then
      mvarCalEvent.Reminder = mvarMe.ChkReminder.Value
      mvarCalEvent.ReminderSoundFile = "D:\Backup_10_12\Desktop\mustbuild.wav"
   End If
   
   If mvarMe.ChkReminder.Value Then
      If Not Val(mvarMe.CmbReminder.Text) = mvarCalEvent.ReminderMinutesBeforeStart Then
         mvarCalEvent.ReminderMinutesBeforeStart = CalcStandardDurations_0m_2wLong(mvarMe.CmbReminder.Text)
      End If
   End If
End Sub
Private Sub MontaCmbTime()
   Dim i As Long
   Dim sHH As String
   Dim BeginTime As Date
   
   On Error GoTo Error
   
   BeginTime = #12:00:00 AM#
   mvarMe.CmbStartTime.Clear
   mvarMe.CmbEndTime.Clear
   For i = 1 To 47
      sHH = Format(TimeValue(BeginTime + i / 24 / 2), "hh:mm")
      mvarMe.CmbStartTime.AddItem sHH
      mvarMe.CmbEndTime.AddItem sHH
   Next i
   
Error:
    
End Sub
Private Sub SetStartEnd(BeginSelection As Date, EndSelection As Date, AllDay As Boolean)
    Dim StartDate As Date
    Dim StartTime As Date
    Dim EndDate   As Date
    Dim EndTime   As Date
    Dim Sql       As String
    Dim MyRS      As Object

    StartDate = DateValue(BeginSelection)
    EndDate = DateValue(EndSelection)
    
    StartTime = TimeValue(BeginSelection)
    EndTime = TimeValue(EndSelection)

   If AllDay Then
      mvarMe.CmbEndDate.Enabled = False
      mvarMe.CmbEndTime.Enabled = False
      mvarMe.CmbStartTime.Enabled = False
      'If DateDiff("s", EndTime, 0) < 0 Then
      '   EndDate = EndDate - 1
      'End If
   End If
    
   mvarMe.CmbStartDate.Value = StartDate
   If mvarMe.CmbStartTime = "" Then
      'DoEvents
      Sql = ""
      Sql = Sql & "Select IDEVENTO, SUBJECT, STARTDATETIME, ENDDATETIME "
      Sql = Sql & " From OEVENTOAGENDA"
      Sql = Sql & " Where ISMEETING = 1"
      Sql = Sql & " And FLGCANCELADO<>1"
      Sql = Sql & " And IDLOJA=" & SqlNum(mvarIDLOJA)
      Sql = Sql & " And (StartDateTime <=" & SqlDate(StartDate & " " & StartTime) & " And EndDateTime >" & SqlDate(StartDate & " " & StartTime) & ")"
      If mvarMe.CmbSchedule.ListIndex < 0 Then
         Sql = Sql & " And IDSALA=1"
      Else
         Sql = Sql & " And IDSALA=" & mvarMe.CmbSchedule.ItemData(mvarMe.CmbSchedule.ListIndex)
      End If
      Sql = Sql & " Order by StartDateTime"
      If mvarSys.xDb.AbreTabela(Sql, MyRS) Then
         If Format(MyRS("ENDDATETIME"), "hh:mm") > Format(BeginSelection, "hh:mm") Then
            mvarMe.CmbStartTime.Text = Format(MyRS("ENDDATETIME"), "hh:mm")
         End If
      End If
      Sql = ""
      Sql = Sql & "Select IDEVENTO, SUBJECT, STARTDATETIME, ENDDATETIME "
      Sql = Sql & " From OEVENTOAGENDA"
      Sql = Sql & " Where ISMEETING = 1"
      Sql = Sql & " And FLGCANCELADO<>1"
      Sql = Sql & " And IDLOJA=" & SqlNum(mvarIDLOJA)
      Sql = Sql & " And (StartDateTime >" & SqlDate(StartDate & " " & StartTime) & " And StartDateTime <" & SqlDate(EndDate & " " & EndTime) & ")"
      If mvarMe.CmbSchedule.ListIndex < 0 Then
         Sql = Sql & " And IDSALA=1"
      Else
         Sql = Sql & " And IDSALA=" & mvarMe.CmbSchedule.ItemData(mvarMe.CmbSchedule.ListIndex)
      End If
      Sql = Sql & " Order by StartDateTime"
      If mvarSys.xDb.AbreTabela(Sql, MyRS) Then
         If Format(MyRS("STARTDATETIME"), "hh:mm") < Format(EndSelection, "hh:mm") Then
            mvarMe.CmbEndTime.Text = Format(MyRS("STARTDATETIME"), "hh:mm")
         End If
      End If
   End If
   If Not IsDate(mvarMe.CmbStartTime.Text) Then
      If -1 = LocalizarCombo(mvarMe.CmbStartTime, Format(BeginSelection, "hh:mm")) Then
         mvarMe.CmbStartTime.Text = Format(BeginSelection, "hh:mm")
      End If
   End If
   If Not IsDate(mvarMe.CmbEndTime.Text) Then
      If -1 = LocalizarCombo(mvarMe.CmbEndTime, Format(EndSelection, "hh:mm")) Then
         mvarMe.CmbEndTime.Text = Format(EndSelection, "hh:mm")
      End If
   End If
    
'    UpdateEndTimeCombo
    mvarMe.CmbEndDate.Value = EndDate
'    Call LocalizarCombo(mvarMe.cmbEndTime, Hour(EndTime) * 60 + Minute(EndTime), True, True)
'    mvarMe.cmbEndTime.Text = CStr(EndTime)
End Sub
'Private Sub UpdateEndTimeCombo(Optional pEndTime As Date)
'   On Error GoTo Error
'   Dim BeginTime  As Date
'   Dim EndTime    As Date
'   Dim i          As Long
'
'   With mvarMe.cmbEndTime
'      .Clear
'      BeginTime = TimeValue(mvarMe.cmbStartTime.Text)
'
'      For i = 1 To 48
'         EndTime = TimeValue(BeginTime + (i / 48))
'         .AddItem EndTime
'         .ItemData(.NewIndex) = Hour(EndTime) * 60 + Minute(EndTime)
'
'         If i = 1 Then
'            .List(.NewIndex) = .List(.NewIndex) & " (30 minutos)"
'         ElseIf i = 2 Then
'            .List(.NewIndex) = .List(.NewIndex) & " (1 hora)"
'         Else
'            .List(.NewIndex) = .List(.NewIndex) & " (" & i / 2 & " horas)"
'         End If
'
'      Next i
'      Call SendMessage(.hwnd, CB_SETDROPPEDWIDTH, 200, 0)
'
'      If Not IsMissing(pEndTime) Then
'         Call LocalizarCombo(mvarMe.cmbEndTime, Hour(pEndTime) * 60 + Minute(pEndTime), True, True)
'         mvarMe.cmbEndTime.Text = CStr(pEndTime)
'      End If
'   End With
'
'Error:
'
'End Sub
Private Sub MontaSubject()
   Dim sNome As String
   Dim sTel  As String
   Dim sAux  As String
   Dim i     As Integer
   
   
   With mvarMe
      If .ChkMeeting.Value = vbChecked Then
         sNome = Trim(.TxtNOME.Text)
         sTel = IIf(Trim(.TxtTEL2.Text) = "", Trim(.TxtTEL1.Text), Trim(.TxtTEL2.Text))
         
         For i = 1 To Len(sTel)
            If isAlfaNum(Mid(sTel, i, 1)) Then
               sAux = sAux & Mid(sTel, i, 1)
            End If
         Next
         sTel = sAux
         
         .TxtSubject.Text = sNome & IIf(sTel = "", "", " - ") & sTel
      End If
   End With
End Sub
Private Sub AdicionarTratamento()
   Dim i As Integer
   Dim MySel As Collection
   Dim PreSel As Collection
   Dim nChave As Collection
   
   Screen.MousePointer = vbHourglass
   
   Set PreSel = New Collection
   With mvarMe.GrdSERVICOEVT
      For i = 1 To .RowCount - 1
         If Not ExisteItem(PreSel, "k" & .CellValue(i, "IDTPTRATAMENTO")) Then
            PreSel.Add .CellValue(i, "IDTPTRATAMENTO"), "k" & .CellValue(i, "IDTPTRATAMENTO")
         End If
      Next
      Set MySel = SelecionarTratamento(PreSel)
      If Not MySel Is Nothing Then
         .Clear
         Call IncluiGrdLinha
         If .RowCount > 1 Then
            .RemoveRow 1
         End If
         i = 0
         For Each nChave In MySel
            Call IncluiGrdLinha
            i = i + 1
            .CellValue(i, "IDSERVICOEVT") = i
            If i = 1 Then
               .CellValue(i, "IDTPSERVICO") = IIf(mvarMe.ChkFLGAVALIACAO.Value, 1, 2)
            Else
               .CellValue(i, "IDTPSERVICO") = .CellValue(i - 1, "IDTPSERVICO")
            End If
            .CellValue(i, "IDTPTRATAMENTO") = xVal(nChave("IDTPTRATAMENTO"))
            .CellValue(i, "IDAREA") = 0
         Next
      End If
   End With
   Screen.MousePointer = vbDefault
End Sub
Private Function SelecionarTratamento(Optional pPreSel As Collection) As Collection
   Dim MyAdd As Object
   Dim Sql As String
   
   Sql = "Select '' [CHK], IDTPTRATAMENTO, DSCTRATAMENTO"
   Sql = Sql + " From OTPTRATAMENTO"
   Sql = Sql + " Where ATIVO =1"
   Sql = Sql + " Order By DSCTRATAMENTO"
   
   Set MyAdd = CriarObjeto("XActive.DsAdd", False)
   If Not MyAdd Is Nothing Then
      With MyAdd
        .Limpar
         Set .xDb = mvarSys.xDb
         .Caption = "Tratamentos"
         Set .Chaves = New Collection
         .Chaves.Add "IDTPTRATAMENTO", "IDTPTRATAMENTO"
         
         With .TITULO
            '0-eAlin.Esquerda 1-eAlin.Direita 2-eAlin.Centro)
            Call .Add("CHK", "", 3, 0)
            Call .Add("IDTPTRATAMENTO", "Id.", 0, 0)
            Call .Add("DSCTRATAMENTO", "Descrição", 50, 0)
         End With
         Set .Preselecao = pPreSel
         .CheckBox = True
         .Query = Sql
   
         .Show
         Set SelecionarTratamento = .Selecao
      End With
   End If
End Function

