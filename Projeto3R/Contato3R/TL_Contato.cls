VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_Contato"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private WithEvents mvarMe     As FrmContato
Attribute mvarMe.VB_VarHelpID = -1
Private WithEvents mvarPane   As FrmPaneContato
Attribute mvarPane.VB_VarHelpID = -1
Private WithEvents mvarCmmdBars As XtremeCommandBars.CommandBars
Attribute mvarCmmdBars.VB_VarHelpID = -1

Private rsContatos As Object

Private Enum IdTaskItem
   ID_TASKITEM_ADDCONTATO = 1
   ID_TASKITEM_EDTCONTATO = 2
   ID_TASKITEM_DELCONTATO = 3
   ID_TASKITEM_MALADIRETA = 4
End Enum
Private Enum eMenuCon
   Mnu_Contato = 2200
   Mnu_MalaDireta
End Enum
'Const ID_TASKITEM_NEWFOLDER = 4
'Const ID_TASKITEM_PUBLISH = 5
'Const ID_TASKITEM_SHARE = 6
'Const ID_TASKITEM_MYCOMPUTER = 7
'Const ID_TASKITEM_MYDOCUMENTS = 8
'Const ID_TASKITEM_SHAREDDOCUMENTS = 9
'Const ID_TASKITEM_MYNETWORKPLACES = 10

Const FCONTROL = 8
Const iId = 2
Dim iCol As Collection
Private mvarSys As Object
Private mvarIDLOJA As Integer
Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
    
   mvarIDLOJA = mvarSys.Propriedades("IDLOJA")
   Set mvarCmmdBars = mvarSys.MDI.CommandBars
End Property
Public Property Get Sys() As Object
   Set Sys = mvarSys
End Property
Public Property Get IDLOJA() As Integer
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDLOJA(ByVal vNewValue As Integer)
   mvarIDLOJA = vNewValue
End Property
Public Property Get Pane() As Object
   Set Pane = mvarPane
End Property
Public Sub Show(Optional Reload As Boolean = False)
   If Reload Then
      Set Sys = Sys
      Unload mvarMe
   End If
   
   '* Acoplar Form no MDI
   Load mvarMe
   Call AcoplarForm(mvarMe, 2, mvarSys)
   mvarMe.Show
   
   '* Exibir Menu no MDI
   With mvarSys.MDI.CommandBars
      If Not .FindControl(, Mnu_Contato, False) Is Nothing Then
         .FindControl(, Mnu_Contato).Visible = True
      End If
   End With

   
   '* Definir foco no Lista de Clientes
   With mvarMe.GrdContato
      If .Visible And .Enabled Then
         .SetFocus
      End If
   End With
   If Reload Then Call mvarMe_Rezise
End Sub
Public Sub Hide()
   mvarMe.Hide
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmContato
   Set mvarPane = New FrmPaneContato
End Sub
Private Sub Class_Terminate()
   On Error Resume Next
   Unload mvarPane
   Set mvarPane = Nothing
   Set mvarMe = Nothing
End Sub

Private Sub mvarCmmdBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
   Dim n As Object
   
   DoEvents
   Select Case UCase(Control.Category)
      Case "CONTATO"
         Select Case Control.Id
            Case Mnu_MalaDireta: Call EnviarMalaDireta
         End Select
   End Select
End Sub

Private Sub mvarMe_GRdContatoRowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
   Dim TlContato As New TL_CadCliente
   
   Screen.MousePointer = vbHourglass
   With TlContato
      Set .Sys = mvarSys
      'mvarMe.GRdContato.SelectedRows(0).Record(iId).Value
      .IDLOJA = Val(Row.Record.Item(iCol("IDLOJA")).Value)
      .IDCLIENTE = Val(Row.Record.Item(iCol("Id")).Value)
      Call .Show
      If .IsDirt Then
         f_CarregaContatos
      End If
   End With
   Set TlContato = Nothing
   
   mvarMe.GrdContato.Redraw
   mvarMe.GrdContato.SetFocus
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_Activate()
   'Call PopulaTela
End Sub
Private Sub mvarMe_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF5 Then
      Call f_CarregaContatos
   ElseIf KeyCode = vbKeyEscape Then
         mvarMe.txtFiltrar.Text = "Pesquisar Contatos"
         Call f_CarregaContatos
   ElseIf (Asc(UCase(Chr(KeyCode))) >= Asc("A") And Asc(UCase(Chr(KeyCode))) >= Asc("Z")) Or (Asc(UCase(Chr(KeyCode))) >= Asc("0") And Asc(UCase(Chr(KeyCode))) >= Asc("9")) Then
      If Not mvarMe.ActiveControl Is mvarMe.txtFiltrar Then
         mvarMe.txtFiltrar.Text = UCase(Chr(KeyCode))
         mvarMe.txtFiltrar.SetFocus
         mvarMe.txtFiltrar.SelStart = Len(mvarMe.txtFiltrar.Text)
      End If
   End If
End Sub
Private Sub mvarMe_Load()
   Call MontarMenu
   Call PopulaTela
   Screen.MousePointer = vbDefault
End Sub
Private Sub PopulaTela()
   'ReportControlGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
   'ShortcutBarGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
   With mvarMe
      .SccContato.VisualTheme = xtpShortcutThemeOffice2003
      .SccContato.GradientHorizontal = False
      With .GrdContato
         .BorderStyle = xtpBorderFrame
         .PaintManager.ColumnStyle = xtpColumnOffice2007
         .PaintManager.ColumnOffice2007CustomThemeBaseColor = -1
         .PaintManager.RefreshMetrics
         .Redraw
      End With
   End With
   
   Call f_CarregaContatos
   mvarMe_Rezise
End Sub
Private Sub MontarMenu()
   Dim oToolBar      As CommandBars
   Dim oCommBar      As CommandBar
   Dim oMenuMain     As CommandBarControl
   Dim oMenuItem     As CommandBarControl
   Dim oMenuSubItem  As CommandBarControl
   Dim oMenuItemP    As CommandBarPopup
   Dim bMenuCarregado As Boolean
   
   On Error GoTo TrataErro
      
   Set oToolBar = mvarSys.MDI.CommandBars
   With oToolBar
      If .ActiveMenuBar Is Nothing Then
         Set oCommBar = .Add("Menu Contato", xtpBarTop)
         oCommBar.Closeable = False
         oCommBar.Customizable = False
         oCommBar.EnableDocking xtpFlagHideWrap
      Else
         Set oCommBar = .ActiveMenuBar
      End If
      
      Set oMenuMain = oCommBar.FindControl(, Mnu_Contato)
      If Not oMenuMain Is Nothing Then
         If (oMenuMain.Category = "CONTATO") Then
            oMenuMain.Visible = True
            GoTo Saida
         End If
      End If
      
      With oCommBar
         '************
         '* Menu Contato
         Set oMenuMain = AddButtonBar(.Controls, Mnu_Contato, "Contato", False, xtpControlPopup, "CONTATO")
         With oMenuMain.CommandBar
            AddButtonBar .Controls, Mnu_MalaDireta, "Mala Direta", Category:="CONTATO"
         End With
         'Call MontarContextMenu
      End With
      .RecalcLayout
   End With
GoTo Saida

TrataErro:
   'Resume Next
   MsgBox "Erro em MontarMenu. Erro: " & Err.Number & "-" & Err.Description
Saida:
End Sub
Private Sub mvarMe_Rezise()
   On Error Resume Next
   mvarMe.SccContato.Move 0, 50, mvarMe.ScaleWidth - 50, 420

   mvarMe.GrdContato.Left = 0
   mvarMe.GrdContato.Top = mvarMe.SccContato.Top + mvarMe.SccContato.Height + 30
   mvarMe.GrdContato.Width = mvarMe.ScaleWidth - 50
   mvarMe.GrdContato.Height = mvarMe.ScaleHeight - mvarMe.GrdContato.Top - mvarMe.LblStGrd.Height - 70
   
   mvarMe.LblStGrd.Top = mvarMe.GrdContato.Top + mvarMe.GrdContato.Height - 20
   mvarMe.LblStGrd.Left = mvarMe.GrdContato.Left
   mvarMe.LblStGrd.Width = mvarMe.GrdContato.Width

   mvarMe.txtFiltrar.Left = mvarMe.Width - 2800
   mvarMe.imgLupa.Left = mvarMe.Width - 680

   mvarMe.txtFiltrar.Top = 90
   mvarMe.imgLupa.Top = 100
   mvarMe.imgLupa.Height = mvarMe.txtFiltrar.Height - 25

End Sub

Private Sub mvarMe_txtFiltrarGotFocus()
   mvarMe.txtFiltrar.ForeColor = vbBlack
   
   If mvarMe.txtFiltrar.Text = "Pesquisar Contatos" Then
      mvarMe.txtFiltrar.Text = ""
   End If
End Sub

Private Sub mvarMe_txtFiltrarKeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      mvarMe.GrdContato.FilterText = mvarMe.txtFiltrar.Text
      mvarMe.GrdContato.Populate
   End If
End Sub

Private Sub mvarMe_txtFiltrarLostFocus()
   
   If mvarMe.txtFiltrar.Text = "" Then mvarMe.txtFiltrar.Text = "Pesquisar Contatos"
   mvarMe.txtFiltrar.ForeColor = 12632256
End Sub

Private Sub mvarMe_Unload(Cancel As Integer)
  On Error Resume Next
'   Unload FrmLembrete
End Sub
'=============================
'=============================
Public Property Get PaneContact() As Object
Attribute PaneContact.VB_Description = "Formulário"
Attribute PaneContact.VB_UserMemId = 0
   Set PaneContact = mvarPane
End Property
Private Sub mvarPane_ChkFiltroClick()
   With mvarPane
      If .ChkAtivo = xtpUnchecked And .ChkEmEspera.Value = xtpUnchecked And .ChkInativo = xtpUnchecked Then
         Call ExibirInformacao("você deve deixar ao menos uma opção selecionada.")
         .ChkAtivo = xtpChecked
         Exit Sub
      End If
   End With
   Call f_CarregaContatos
End Sub
Private Sub mvarPane_CmdPerquisaClick()
   Dim Sql As String
   Dim sAux As String
   
   Sql = ""
   With mvarPane
      If .txtNome.Text <> "" Then
         sAux = IIf(InStr(.txtNome.Text, "%") = 0, "%" + .txtNome.Text + "%", .txtNome.Text)
         Sql = Sql + " And NOME Like '" + sAux + "'"
      End If
      If .txtTel.Text <> "" Then
         sAux = IIf(InStr(.txtTel.Text, "%") = 0, "%" + .txtTel.Text + "%", .txtTel.Text)
         Sql = Sql + "And (TEL1 Like '" + sAux + "'"
         Sql = Sql + " Or TEL2 Like '" + sAux + "'"
         Sql = Sql + " Or FAX Like '" + sAux + "'"
         Sql = Sql + " )"
      End If
      If .TxtBairro.Text <> "" Then
         sAux = IIf(InStr(.TxtBairro.Text, "%") = 0, "%" + .TxtBairro.Text + "%", .TxtBairro.Text)
         Sql = Sql + "And BAIRRO Like '" + sAux + "'"
      End If
   
   End With

   Call f_CarregaContatos(Sql)
End Sub
Private Sub mvarPane_KeyPress(KeyAscii As Integer)
   If Between(UCase(Chr(KeyAscii)), "A", "Z") Then
'      mvarMe.txtFiltrar.Text = mvarMe.txtFiltrar.Text & Chr(KeyCode)
'      mvarMe.txtFiltrar.SelStart = Len(mvarMe.txtFiltrar.Text)
      If Not (mvarPane.ActiveControl Is mvarPane.txtNome Or mvarPane.ActiveControl Is mvarPane.txtTel Or mvarPane.ActiveControl Is mvarPane.TxtBairro Or mvarPane.ActiveControl Is mvarPane.ChkAtivo Or mvarPane.ActiveControl Is mvarPane.ChkEmEspera Or mvarPane.ActiveControl Is mvarPane.ChkInativo) Then
         DoEvents
         mvarMe.txtFiltrar.SetFocus
      End If
   End If

End Sub
Private Sub mvarPane_KeyUp(KeyCode As Integer, Shift As Integer)
'   If Mid(mvarMe.txtFiltrar.Text, 1, 9) = "Pesquisar" Then
'      mvarMe.txtFiltrar.Text = ""
'   End If
   If Between(UCase(Chr(KeyCode)), "A", "Z") Then
'      mvarMe.txtFiltrar.Text = mvarMe.txtFiltrar.Text & Chr(KeyCode)
'      mvarMe.txtFiltrar.SelStart = Len(mvarMe.txtFiltrar.Text)
      If Not (mvarPane.ActiveControl Is mvarPane.txtNome Or mvarPane.ActiveControl Is mvarPane.txtTel Or mvarPane.ActiveControl Is mvarPane.TxtBairro Or mvarPane.ActiveControl Is mvarPane.ChkAtivo Or mvarPane.ActiveControl Is mvarPane.ChkEmEspera Or mvarPane.ActiveControl Is mvarPane.ChkInativo) Then
         DoEvents
         mvarMe.txtFiltrar.SetFocus
      End If
   End If
End Sub
Private Sub mvarPane_Load()
   'MsgBox "mvarPaneLoad"
   DoEvents
   Call CreateTaskPanel
End Sub
Private Sub mvarPane_Resize()
   mvarPane.SccContato.Move 0, 0, mvarPane.ScaleWidth, 420
   mvarPane.SccContato2.Move 0, mvarPane.SccContato.Height, mvarPane.SccContato.Width, 285
   
   mvarPane.Picture1.Move 0, mvarPane.SccContato2.Top + mvarPane.SccContato2.Height, mvarPane.ScaleWidth, mvarPane.ScaleHeight - mvarPane.SccContato.Height
   mvarPane.wndTaskPanel.Move 0, 0, mvarPane.Picture1.ScaleWidth, mvarPane.Picture1.ScaleHeight
   
End Sub
Private Sub f_CarregaContatos(Optional pWhere As String)
   Dim Sql As String
   Dim i As Long
   Dim sFiltroAtivo As String
      
   sFiltroAtivo = ""
   sFiltroAtivo = sFiltroAtivo & IIf(mvarPane.ChkAtivo.Value = xtpChecked, IIf(sFiltroAtivo = "", "", ", ") & "Null", "")
   sFiltroAtivo = sFiltroAtivo & IIf(mvarPane.ChkInativo.Value = xtpChecked, IIf(sFiltroAtivo = "", "", ", ") & "0", "")
   sFiltroAtivo = sFiltroAtivo & IIf(mvarPane.ChkAtivo.Value = xtpChecked, IIf(sFiltroAtivo = "", "", ", ") & "1", "")
   sFiltroAtivo = sFiltroAtivo & IIf(mvarPane.ChkEmEspera.Value = xtpChecked, IIf(sFiltroAtivo = "", "", ", ") & "2", "")
   
   
   Sql = "SELECT '' [Img], '' as [ATIVO], '' [IDLOJA], '' [IDCOLIGADA], '' [Id], '' [Nome]"
   Sql = Sql & ", '' [Tel. Celular], '' [Tel. Resid.], '' [eMail], '' [Fototipo], '' [Prox.Sessão]"
   Sql = Sql & ", '' [Últ.Tempo]"
   'Sql = Sql & " FROM OCLIENTE "
   Call mvarSys.xDb.Abretabela(Sql, rsContatos)
   Call FillRCFromRS(rsContatos, mvarMe.GrdContato)
   mvarMe.GrdContato.Records.DeleteAll
         
   Sql = ""
   'Sql = Sql & " Select Row_Number() Over (Order By NOME ) [Lin], *"
   'Sql = Sql & " From "
   'Sql = Sql & "("
   Sql = Sql & "Select '' As Img, C.ATIVO, C.IDLOJA, C.IDCOLIGADA, C.IDCLIENTE As Id, C.NOME As Nome"
   Sql = Sql & ", C.TEL1 [Tel. Celular], C.TEL2 [Tel. Resid.], C.EMAIL [eMail]"
   Sql = Sql & ", C.FOTOTIPO [Fototipo]"
   'Sql = Sql & ", E.StartDateTime [Prox.Sessão]"
   Sql = Sql & ", Convert(varchar(5),E.StartDateTime , 103) + ' - ' + Convert(varchar(5),E.StartDateTime , 108)"
   Sql = Sql & " [Prox.Sessão]"
   
   Sql = Sql + " , Right('00'+ Cast(DATEDIFF(mi,A.HHINI, A.HHFIM)/60 As varchar),2)+':'"
   Sql = Sql + " + Right('00'+ Cast(DATEDIFF(mi,A.HHINI, A.HHFIM)%60 As varchar),2)"
   Sql = Sql & " [Últ.Tempo]"
   
   
   Sql = Sql & " From OCLIENTE C"
   Sql = Sql & " Left Join OEVENTOAGENDA E On E.IDLOJA=C.IDLOJA AND E.IDCLIENTE=C.IDCLIENTE And E.FLGCANCELADO=0 And E.StartDateTime>=GetDate() And E.StartDateTime=(Select Min(E2.StartDateTime) From OEVENTOAGENDA E2 Where E2.IDLOJA=C.IDLOJA AND E2.IDCLIENTE=C.IDCLIENTE And E2.StartDateTime>=GetDate())"
   Sql = Sql & " Left Join OATENDIMENTO  A On A.IDLOJA=C.IDLOJA AND A.IDCLIENTE=C.IDCLIENTE And A.DTATEND=(Select Max(A2.DTATEND) From OATENDIMENTO A2 Where A2.IDLOJA=C.IDLOJA AND A2.IDCLIENTE=C.IDCLIENTE)"
   Sql = Sql & " Where C.IDLOJA=" & mvarIDLOJA
   If Trim(sFiltroAtivo) <> "" Then
      Sql = Sql & " And C.ATIVO in (" & sFiltroAtivo & ")"
   End If
   If pWhere <> "" Then
      Sql = Sql & pWhere
   End If
   'Sql = Sql & ") TbRow"
   Sql = Sql & " Order By C.NOME"
   
   With mvarMe.GrdContato
      .SetImageList mvarMe.ImageList1
     
      If mvarSys.xDb.Abretabela(Sql, rsContatos) Then
         Call FillRCFromRS(rsContatos, mvarMe.GrdContato)
      End If
      DoEvents
      
      Set iCol = New Collection
      For i = 0 To rsContatos.Fields.Count - 1
         iCol.Add i, rsContatos.Fields(i).Name
      Next
      .Columns(iCol("ATIVO")).Visible = False
      .Columns(iCol("IDCOLIGADA")).Visible = False
      .Columns(iCol("IDLOJA")).Visible = False
      .Columns(iCol("Id")).Visible = False
      
      For i = 0 To .Records.Count - 1
         .Records(i).Item(iCol("Img")).Icon = .Records(i).Item(iCol("ATIVO")).Value
         '.Records(i).Item(0).Value = .Records(i).Item(1).Value
         '.Records(i).Item(0).BackColor = &HE0E0E0
         '.Records(i).Item(0).ForeColor = .Records(i).Item(0).BackColor
      Next
      
      .Columns(iCol("Tel. Celular")).HeaderAlignment = xtpAlignmentLeft
      .Columns(iCol("Tel. Resid.")).HeaderAlignment = xtpAlignmentLeft
      .Columns(iCol("eMail")).HeaderAlignment = xtpAlignmentLeft
      .Columns(iCol("Fototipo")).HeaderAlignment = xtpAlignmentCenter
      .Columns(iCol("Prox.Sessão")).HeaderAlignment = xtpAlignmentCenter
      .Columns(iCol("Últ.Tempo")).HeaderAlignment = xtpAlignmentCenter
      
      .Columns(iCol("Tel. Celular")).Alignment = xtpAlignmentLeft
      .Columns(iCol("Tel. Resid.")).Alignment = xtpAlignmentLeft
      .Columns(iCol("eMail")).Alignment = xtpAlignmentLeft
      .Columns(iCol("Fototipo")).Alignment = xtpAlignmentCenter
      .Columns(iCol("Prox.Sessão")).Alignment = xtpAlignmentCenter
      .Columns(iCol("Últ.Tempo")).Alignment = xtpAlignmentCenter
           
      For i = 0 To .Columns.Count - 1
         .Columns(i).BestFit
      Next
      '.Columns(0).Width = 20
      '.Columns(0).Caption = "#"
      .Columns(iCol("Img")).Width = 20
      .Columns(iCol("Img")).Caption = ""
      .Columns(iCol("Img")).Icon = 3
      
      .Columns(iCol("Nome")).Width = 180
      If .Columns(iCol("Nome")).Width < 200 Then .Columns(iCol("Nome")).Width = 200
      If .Columns(iCol("Tel. Celular")).Width < 100 Then .Columns(iCol("Tel. Celular")).Width = 100
      If .Columns(iCol("Tel. Resid.")).Width < 100 Then .Columns(iCol("Tel. Resid.")).Width = 100
      If .Columns(iCol("eMail")).Width < 200 Then .Columns(iCol("eMail")).Width = 200
      If .Columns(iCol("Fototipo")).Width < 60 Then .Columns(iCol("Fototipo")).Width = 60
      If .Columns(iCol("Prox.Sessão")).Width < 80 Then .Columns(iCol("Prox.Sessão")).Width = 80
      If .Columns(iCol("Prox.Sessão")).Width < 60 Then .Columns(iCol("Prox.Sessão")).Width = 60
                  
      If mvarMe.txtFiltrar.Text = "" Or mvarMe.txtFiltrar.Text = "Pesquisar Contatos" Then
         .FilterText = ""
      Else
         .FilterText = mvarMe.txtFiltrar.Text
      End If
      If mvarMe.GrdContato.Records.Count = 1 Then
         mvarMe.LblStGrd.Caption = "Total: ( 1 Item )"
      Else
         mvarMe.LblStGrd.Caption = "Total: ( " & mvarMe.GrdContato.Records.Count & " Itens )"
      End If
      .Populate
      .Redraw
   End With

End Sub
Private Sub CreateTaskPanel()
   Dim Group As TaskPanelGroup
   Dim Item As TaskPanelGroupItem
   
   mvarPane.wndTaskPanel.SetImageList mvarPane.imlTaskPanelIcons
   
   Set Group = mvarPane.wndTaskPanel.Groups.Add(0, "Operações")
   Group.Items.Add ID_TASKITEM_ADDCONTATO, "Adicionar novo contato", xtpTaskItemTypeLink, 17
   Group.Items.Add ID_TASKITEM_EDTCONTATO, "Editar contato selecionado", xtpTaskItemTypeLink, 18
   Group.Items.Add ID_TASKITEM_DELCONTATO, "Excluir contato selecionado", xtpTaskItemTypeLink, 19
   Group.Expandable = False
    
   Set Group = mvarPane.wndTaskPanel.Groups.Add(0, "Filtros")
   Set Item = Group.Items.Add(0, "", xtpTaskItemTypeControl)
    Set Item.Control = mvarPane.FraFiltro
    mvarPane.FraFiltro.BackColor = Item.BackColor
    mvarPane.ChkAtivo.BackColor = Item.BackColor
    mvarPane.ChkEmEspera.BackColor = Item.BackColor
    mvarPane.ChkInativo.BackColor = Item.BackColor
'   Group.Items.Add ID_TASKITEM_NEWFOLDER, "Visualizar tratamentos", xtpTaskItemTypeControl, 1
'   Group.Items.Add ID_TASKITEM_PUBLISH, "Últimas comunicações", xtpTaskItemTypeLink, 2
'   Group.Items.Add ID_TASKITEM_SHARE, "Histório de compra do creme", xtpTaskItemTypeLink, 3
   Group.Expanded = False
    
   Set Group = mvarPane.wndTaskPanel.Groups.Add(0, "Busca específica")
   Set Item = Group.Items.Add(0, "", xtpTaskItemTypeControl)
   Set Item.Control = mvarPane.fraBuscaDetalhada
   mvarPane.fraBuscaDetalhada.BackColor = Item.BackColor
   Group.Expanded = False
    
   'Set Group = mvarPane.wndTaskPanel.Groups.Add(0, "Ferramentas")
   'Set Item = Group.Items.Add(ID_TASKITEM_MALADIRETA, "Mala Direta", xtpTaskItemTypeLink)
   'Group.Expanded = False
    
    mvarPane.wndTaskPanel.Reposition
End Sub
Private Sub mvarPane_wndTaskPanelItemClick(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem)
   Dim oCAD As TL_CadCliente
   
   Select Case Item.Id
      Case 1:       'ADD CONTATO
         Set oCAD = New TL_CadCliente
         With oCAD
            Set .Sys = mvarSys
            Call .Show
            If .IsDirt = True Then f_CarregaContatos
         End With
         Set oCAD = Nothing
      Case 2:       'EDT CONTATO
         If mvarMe.GrdContato.SelectedRows.Count > 0 Then
            Set oCAD = New TL_CadCliente
            With oCAD
               Set .Sys = mvarSys
               .IDLOJA = mvarIDLOJA
               .IDCLIENTE = mvarMe.GrdContato.SelectedRows(0).Record.Item(iId).Value
               Call .Show
               If .IsDirt = True Then f_CarregaContatos
            End With
            Set oCAD = Nothing
         End If
      Case 3:       'DEL CONTATO
         If mvarMe.GrdContato.SelectedRows.Count > 0 Then
            Set oCAD = New TL_CadCliente
            With oCAD
               Set .Sys = mvarSys
               .IDCLIENTE = mvarMe.GrdContato.SelectedRows(0).Record.Item(iId).Value
               Call .ExcluirCliente
               If .IsDirt = True Then f_CarregaContatos
            End With
            Set oCAD = Nothing
         End If
      Case ID_TASKITEM_MALADIRETA:
          Call EnviarMalaDireta
      Case 5:       '
      Case 6:       '
      
   End Select

End Sub
Private Sub EnviarMalaDireta()
'   Dim bResult As Boolean
'   Dim nMouse  As Integer
'   Dim xMail   As Object
'   Dim sFile As String
'   Dim ODlg  As Object
'   Dim sMessage As String
'   Dim sTitulo  As String
'   Dim sPath As String
'   Dim i As Integer
'   Dim Sql As String
'   Dim oRs As Object
   Dim bAcesso As Boolean
                
   If InStr(mvarSys.ExePath, "\Sistemas\") = 0 Then
      'bAcesso = AcessopEspecial(mvarSys, "CADCLI_MALADIR")
      bAcesso = SenhaMestre(mvarSys, 2)
   Else
      bAcesso = True
   End If
   
   If bAcesso Then
      Dim TlMala As TL_MalaDireta
      Set TlMala = New TL_MalaDireta
      Set TlMala.Sys = mvarSys
      TlMala.Show
      Set TlMala = Nothing
   Else
      Call ExibirAviso("Acesso Inválido.")
   End If
   
'   If bAcesso Then
'      Screen.MousePointer = vbHourglass
'
'      Set ODlg = CreateObject("MSComDlg.CommonDialog")
'      sFile = SelecionarArquivo(ODlg, "Arquivo", "", "*.htm|*.htm")
'      If ExisteArquivo(sFile) Then
'         sMessage = ReadTextFile(sFile)
'         If InStr(sFile, "VERAO.HTM") <> 0 Then
'            sTitulo = InputBox("Informe o título.", "Mala Direta", "Dpil - Promoção de Verão 2012")
'         Else
'            sTitulo = InputBox("Informe o título.", "Mala Direta", "Dpil - Promoção de Natal 2011")
'         End If
'      End If
'      If Trim(sMessage) = "" Then
'         Call ExibirStop("Mensagem Inválida")
'         Exit Sub
'      End If
'      If Trim(sTitulo) = "" Then
'         Call ExibirStop("Título Inválida")
'         Exit Sub
'      End If
'
'      If Trim(sMessage) <> "" And Trim(sTitulo) <> "" Then
'         Set xMail = CriarObjeto("CAMail.SendMail")
'         With xMail
'            .UseAuthentication = (mvarSys.GetParam("UseAuthentication") = xtpChecked)
'            .UsePopAuthentication = (mvarSys.GetParam("UsePopAuthentication") = xtpChecked)
'
'            .POP3Host = mvarSys.GetParam("POP3Host") ' "pop3.bol.com.br"
'            .SMTPHost = mvarSys.GetParam("SMTPHost") ' "smtps.bol.com.br"
'            .SMTPPort = mvarSys.GetParam("SMTPPort") ' 587
'            .Username = mvarSys.GetParam("MailUID")  ' "diogenes72@bol.com.br"
'            .Password = Decrypt2(mvarSys.GetParam("MailPWD"))
'
'            .FromDisplayName = mvarSys.GetParam("FromDisplayName") 'FromDisplayName ' "Diogenes"
'
'            .Receipt = True
'            .SMTPHostValidation = 0 'VALIDATE_HOST_NONE
'
'            .Subject = sTitulo
'            .Message = sMessage
'
'            .RecipientDisplayName = "Cliente"           ' "DiSantos"
'            .From = .Username
'            .AsHTML = True
'
'            sFile = GetNameFromPath(sFile, sPath)
'            .Attachment = sPath & "Promo.jpg;" & sPath & "Dpil.jpg"
'            .Subject = sTitulo
'
'            Dim x As ProgressBar
'            Sql = "Select NOME, EMAIL "
'            Sql = Sql & " From OCLIENTE"
'            Sql = Sql & " Where IDLOJA=" & mvarIDLOJA
'            Sql = Sql & " And ATIVO<>0"
'            Sql = Sql & " And EMAIL<>''"
'            Sql = Sql & " And Not EMAIL is Null"
'            If mvarSys.xDb.AbreTabela(Sql, oRs) Then
'               Set x = mvarSys.MDI.ProgressBar
'               mvarSys.MDI.CommandBars.StatusBar.Pane(3).Handle = x.hwnd
'               x.Visible = True
'               x.Value = 0
'               x.ZOrder 0
'               x.DragMode = 1
'               x.UseVisualStyle = True
'               x.Min = 1
'               x.Max = oRs.recordcount
'               While Not oRs.EOF
'                  .Recipient = oRs("EMAIL")
'                  .Connect
'                  .Send
'                  .Disconnect
'
'                  x.Value = oRs.AbsolutePosition
'                  mvarSys.MDI.Caption = oRs.AbsolutePosition & " / " & oRs.recordcount
'                  oRs.MoveNext
'               Wend
'               x.Visible = False
'            End If
''            For i = 1 To 2
''               If i = 1 Then
''                  .Recipient = "diogenes72@bol.com.br"  ' "disantos@ig.com.br"
''                Else
''                  .Recipient = "adriane76@bol.com.br"  ' "disantos@ig.com.br"
''                End If
''               .Connect
''               .Send
''               .Disconnect
''            Next
'            bResult = .SendSuccesful
'
'         End With
'         Set xMail = Nothing
'      End If
'
'      Screen.MousePointer = vbDefault
'      MsgBox IIf(bResult, "Ok.", "Não Ok")
'   Else
'      Call ExibirAviso("Acesso Inválido.")
'   End If
End Sub

