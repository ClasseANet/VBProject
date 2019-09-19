VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_VENDAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe     As FrmVendas
Attribute mvarMe.VB_VarHelpID = -1
Private WithEvents PopupMenu  As CommandBars
Attribute PopupMenu.VB_VarHelpID = -1

Private Enum eIconId
   Novo = 1
   Editar
   RPS
   NF
   ExpRPS
   ImpNF
   Sair
   Cliente
   BaixarSaldo
End Enum
Private Enum ePopVendas
   Pop_Vendas = 100
End Enum

Dim iCampo As Collection

Private RsCad           As Object
Private mvarSys         As Object
Private mvarConsulta    As String
Private mvarIDLOJA      As Integer
Public Sub F_Refresh()
   Call PopulaGrdVendas(True)
End Sub
Public Property Let IDLOJA(ByVal vData As Integer)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Integer
   IDLOJA = mvarIDLOJA
End Property
Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
   mvarIDLOJA = mvarSys.Propriedades("IDLOJA")
End Property
Public Property Get Sys() As Object
   Set Sys = mvarSys
End Property
Public Property Get Consulta() As String
   Consulta = mvarConsulta
End Property
Public Property Let Consulta(ByVal vData As String)
   mvarConsulta = vData
End Property
Public Sub Show()
   mvarMe.Show vbModal
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmVendas
End Sub
Private Sub Class_Terminate()
   On Error Resume Next
   Unload mvarMe
   Set mvarMe = Nothing
End Sub
Private Sub mvarMe_CommandBarsExecute(ByVal Control As XtremeCommandBars.ICommandBarControl)
   Select Case Control.Id
      Case eIconId.Novo:    Call NovaVenda
      Case eIconId.Editar: Call EditarVenda(mvarMe.GrdVendas.SelectedRows(0).Record(iCampo("Nº VENDA")).Value)
      Case eIconId.RPS:    Call EditarRPS
      Case eIconId.NF:     'Call EditarNF
      Case eIconId.ExpRPS: Call ExportarRPS
      Case eIconId.ImpNF:  Call ImportarNF
      Case eIconId.Sair:   Unload mvarMe
      Case "Sair":
   End Select
End Sub
Private Sub EditarRPS()
   Dim Sql As String
   Dim MyRs As Object
   Dim MyRs2 As Object
   
   If False Then Exit Sub
   
   Sql = "Select * "
   Sql = Sql & " From CPGTOSVENDA"
   Sql = Sql & " Where IDLOJA=3"
   Sql = Sql & " And IDVENDA>=1309"
   Sql = Sql & " And IDVENDA<=2423"
   If mvarSys.xDb.AbreTabela(Sql, MyRs) Then
      While Not MyRs.EOF
         Sql = "Update CPGTOSVENDA"
         Sql = Sql & " set IDVENDA=IDVENDA-1"
         Sql = Sql & " , TIMESTAMP=GetDate()"
         Sql = Sql & " , ALTERSTAMP=1"
         Sql = Sql & " Where IDLOJA=" & MyRs("IDLOJA")
         Sql = Sql & " And IDVENDA=" & MyRs("IDVENDA")
         If Not mvarSys.xDb.Executa(Sql) Then
            Sql = Sql
         End If
         MyRs.MoveNext
      Wend
   End If
   
   Sql = "Select * "
   Sql = Sql & " From CVENDA"
   Sql = Sql & " Where IDLOJA=3"
   Sql = Sql & " And IDVENDA>=1308"
   Sql = Sql & " And IDVENDA<=2425"
   If mvarSys.xDb.AbreTabela(Sql, MyRs) Then
      While Not MyRs.EOF
         Sql = "Select A.IDATENDIMENTO"
         Sql = Sql & " from OATENDIMENTO A"
         Sql = Sql & " JOIN OATENDIMENTO_VENDA AV ON A.IDLOJA=AV.IDLOJA AND A.IDATENDIMENTO=AV.IDATENDIMENTO"
         Sql = Sql & " JOIN CVENDA V ON V.IDLOJA=AV.IDLOJA AND V.IDVENDA=AV.IDVENDA"
         Sql = Sql & " Where A.IDLOJA =" & MyRs("IDLOJA")
         Sql = Sql & " AND AV.IDVENDA=" & MyRs("IDVENDA")
         Sql = Sql & " AND A.IDCLIENTE<>V.IDCLIENTE"
         If mvarSys.xDb.AbreTabela(Sql, MyRs2) Then
            While Not MyRs2.EOF
               Sql = ""
               Sql = "DELETE OATENDIMENTO_VENDA"
               Sql = Sql & " WHERE IDLOJA=" & MyRs("IDLOJA")
               Sql = Sql & " AND IDVENDA=" & MyRs("IDVENDA")
               Sql = Sql & " AND IDATENDIMENTO=" & MyRs2("IDATENDIMENTO")
               If Not mvarSys.xDb.Executa(Sql) Then
                  Sql = Sql
               End If
               
               Sql = Sql & ""
               Sql = Sql & "INSERT INTO DELETEDROWS"
               Sql = Sql & " (QUERY) VALUES"
               Sql = Sql & " ('"
               Sql = Sql & "DELETE OATENDIMENTO_VENDA "
               Sql = Sql & " WHERE IDLOJA=" & MyRs("IDLOJA")
               Sql = Sql & " AND IDVENDA=" & MyRs("IDVENDA")
               Sql = Sql & " AND IDATENDIMENTO=" & MyRs2("IDATENDIMENTO")
               Sql = Sql & ""
               Sql = Sql & "')"
               If Not mvarSys.xDb.Executa(Sql) Then
                  Sql = Sql
               End If
               MyRs2.MoveNext
            Wend
         End If
         MyRs.MoveNext
      Wend
   End If
   
   
End Sub
Private Sub mvarMe_GrdVendasBeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
   If Not iCampo Is Nothing Then
      'If Item.Index = iCampo("Depósito") Or Item.Index = iCampo("Pagamento") Or Item.Index = iCampo("Saldo") Then
      '   Item.Value = IIf(xVal(Item.Value & "") = 0, "", ValBr(Item.Value & ""))
      'End If
   End If
   If ((Row.Index Mod 2 = 0) And (Row.GroupRow = False)) Then
       Metrics.BackColor = RGB(245, 245, 245)
   Else
      'Metrics.BackColor = RGB(65, 165, 241)
   End If

End Sub

Private Sub mvarMe_GrdVendasKeyUp(KeyCode As Integer, Shift As Integer)
   Call mvarMe_KeyUp(KeyCode, Shift)
End Sub

Private Sub mvarMe_GrdVendasMouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
   Dim xMenu      As CommandBar
   Dim oMenuItem  As CommandBarControl
   Dim L As Long
   Dim T As Long
   Dim XX As Long
   Dim YY As Long

   XX = 0
   YY = 0
   While mvarMe.GrdVendas.HitTest(XX, YY).ht <> xtpHitTestReportArea
     YY = YY + 1
   Wend
   XX = 0
   While mvarMe.GrdVendas.HitTest(XX, YY).ht = xtpHitTestReportArea
      XX = XX + 1
   Wend
   XX = XX - 1
   While mvarMe.GrdVendas.HitTest(XX, YY).ht = xtpHitTestReportArea
      YY = YY + 1
   Wend
   YY = YY - 1
   If Button = 2 And mvarMe.GrdVendas.HitTest(x, y).ht = xtpHitTestReportArea Then
      '**************
      '* Pop-up Edit
      On Error Resume Next
      If PopupMenu Is Nothing Then
         Set PopupMenu = mvarSys.MDI.CommandBars
      End If
      mvarSys.MDI.CommandBars.ContextMenus(99).Delete
      
      Set xMenu = PopupMenu.ContextMenus.Add(99, "Menu Aux")
      With xMenu
         Set oMenuItem = AddButtonBar(.Controls, eIconId.Editar, "Abrir Venda", False, 1, "VENDAS")
         oMenuItem.DefaultItem = True
         AddButtonBar .Controls, eIconId.Novo, "Nova Venda", True, 1, "VENDAS"
         AddButtonBar .Controls, eIconId.Cliente, "Cliente", False, 1, "VENDAS"
         AddButtonBar .Controls, eIconId.BaixarSaldo, "Baixar Saldo", True, 1, "VENDAS"
      End With
      With mvarMe.GrdVendas
         xMenu.Controls.Find(, eIconId.BaixarSaldo).Enabled = (xVal(.SelectedRows(0).Record(iCampo("Saldo Serv.")).Value) <> 0)
         L = .Container.Left + .Left
         T = .Container.Top + .Top + 320
         'L = L + (x / 993 * .Width)
         'T = T + (y / 610 * .Height)
         L = L + (x / XX * .Width)
         T = T + (y / YY * .Height)
         xMenu.ShowPopup , L, T
      End With
   End If
End Sub
Private Sub PopupMenu_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
   If Control.Category = "VENDAS" Then
      mvarMe.MousePointer = vbHourglass

      Select Case Control.Id
         Case eIconId.Novo:   Call NovaVenda
         Case eIconId.Editar: Call EditarVenda(mvarMe.GrdVendas.SelectedRows(0).Record(iCampo("Nº VENDA")).Value)
         Case eIconId.RPS:    Call EditarRPS
         Case eIconId.NF:     'Call EditarNF
         Case eIconId.ExpRPS: Call ExportarRPS
         Case eIconId.ImpNF:  Call ImportarNF
         Case eIconId.Sair:   Unload mvarMe
         Case eIconId.Cliente: Call ShowCliente(mvarSys, mvarIDLOJA, mvarMe.GrdVendas.SelectedRows(0).Record(iCampo("IDCLIENTE")).Value)
         Case eIconId.BaixarSaldo: Call BaixarSaldoV
      End Select
      mvarMe.MousePointer = vbDefault
   End If
End Sub
'Private Sub MontarPopupMenu()
'   Dim SysMenu As CommandBars
'   Dim PopupMenu   As CommandBar
'   Dim oMenuItem     As CommandBarControl
'
'   '**************
'   '* Pop-up Vendas
'   Set SysMenu = mvarSys.MDI.CommandBars
'   Set PopupMenu = SysMenu.ContextMenus.Add(ePopVendas.Pop_Vendas, "Context PopUp Menu")
'   With PopupMenu
'      AddButtonBar .Controls, Mnu_CalEditOpen, "Nova Venda", False, "VENDAS"
'      AddButtonBar .Controls, Mnu_CalEditDelete, "Abrir Venda", True, "VENDAS"
'      AddButtonBar .Controls, Mnu_CalEditCliente, "Cliente", False, "VENDAS"
'      AddButtonBar .Controls, Mnu_CalEditVenda, "Baixar Saldo", False, "VENDAS"
'      AddButtonBar .Controls, Mnu_CalEditConfirm, "Exportar RPS", True, "VENDAS"
'      AddButtonBar .Controls, Mnu_CalEditConfirm, "Importar NF", False, "VENDAS"
'
'   End With
'   PopupMenu.Controls.CreateOriginalControls
'
'
'End Sub
Private Sub mvarMe_GrdVendasRowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
   Call EditarVenda(mvarMe.GrdVendas.SelectedRows(0).Record(iCampo("Nº VENDA")).Value)
End Sub
Private Sub mvarMe_Activate()
   mvarMe_Resize
   If GetTag(mvarMe, "1Vez", 0) = 0 Then
      Call SetTag(mvarMe, "1Vez", 1)
      Call PosicionaForm
   End If

Call MontarToolbar

   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF5 Then
      Call PopulaGrdVendas(True)
   ElseIf KeyCode = vbKeyEscape Then
         mvarMe.txtFiltrar.Text = mvarMe.txtFiltrar.Tag
         Call PopulaGrdVendas(True)
   ElseIf (Asc(UCase(Chr(KeyCode))) >= Asc("A") And Asc(UCase(Chr(KeyCode))) >= Asc("Z")) Or (Asc(UCase(Chr(KeyCode))) >= Asc("0") And Asc(UCase(Chr(KeyCode))) >= Asc("9")) Then
      If Not mvarMe.ActiveControl Is mvarMe.txtFiltrar Then
         mvarMe.txtFiltrar.Text = UCase(Chr(KeyCode))
         mvarMe.txtFiltrar.SetFocus
         mvarMe.txtFiltrar.SelStart = Len(mvarMe.txtFiltrar.Text)
      End If
   End If
End Sub
Private Sub mvarMe_Load()
   mvarMe.txtFiltrar.Tag = mvarMe.txtFiltrar.Text
   
   Set mvarSys.ObjAux = mvarMe
   Call PosicionaForm
   
   Call MontarTela
   'Call MontarMenu
'   Call MontarToolbar
   'Call MontarStatusBar
   Call PopulaTela
   
   Screen.MousePointer = vbDefault
End Sub
Private Sub MontarTela()
   'ReportControlGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
   'ShortcutBarGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
   With mvarMe
'      With .SccContato
'         .VisualTheme = xtpShortcutThemeOffice2003
'         .GradientHorizontal = False
'      End With
      With .GrdVendas
         .BorderStyle = xtpBorderFrame
         .PaintManager.ColumnStyle = xtpColumnOffice2007
         .PaintManager.ColumnOffice2007CustomThemeBaseColor = -1
         .PaintManager.RefreshMetrics
         .Redraw
      End With
   End With
End Sub
Private Sub PopulaTela()
   Call PopulaGrdVendas
   mvarMe_Resize
End Sub
Private Sub mvarMe_Resize()
   On Error Resume Next
   'mvarMe.SccContato.Move 0, 0, mvarMe.ScaleWidth, 400

   mvarMe.GrdVendas.Left = 0
   'mvarMe.GrdVendas.Top = mvarMe.SccContato.Top + mvarMe.SccContato.Height
   mvarMe.GrdVendas.Top = 400
   mvarMe.GrdVendas.Width = mvarMe.ScaleWidth - 50
   mvarMe.GrdVendas.Height = mvarMe.ScaleHeight - mvarMe.GrdVendas.Top - 50

   mvarMe.txtFiltrar.Left = mvarMe.Width - 2800
   mvarMe.imgLupa.Left = mvarMe.Width - 680

   'mvarMe.txtFiltrar.Top = (mvarMe.SccContato.Height - mvarMe.txtFiltrar.Height) / 2
   mvarMe.txtFiltrar.Top = (400 - mvarMe.txtFiltrar.Height) / 2
   mvarMe.imgLupa.Top = mvarMe.txtFiltrar.Top + 5
   mvarMe.imgLupa.Height = mvarMe.txtFiltrar.Height - 30
   
   mvarMe.txtFiltrar.ZOrder 0
   mvarMe.imgLupa.ZOrder 0
End Sub
Private Sub PosicionaForm()
   Dim Left As Long
   Dim Top As Long
   Dim Right As Long
   Dim Bottom As Long
   On Error Resume Next
   mvarSys.MDI.CommandBars.GetClientRect Left, Top, Right, Bottom
   mvarSys.MDI.CommandBars(1).GetWindowRect Left, Top, Right, Bottom
      
   mvarMe.Top = Bottom
   
   mvarMe.Left = mvarSys.MDI.Left + 60
   mvarMe.Width = mvarSys.MDI.Width - mvarMe.Left - 120
   mvarMe.Height = mvarSys.MDI.Height - mvarMe.Top - 420
End Sub
Private Sub mvarMe_txtFiltrarGotFocus()
   mvarMe.txtFiltrar.ForeColor = vbBlack
   
   If mvarMe.txtFiltrar.Text = mvarMe.txtFiltrar.Tag Then
      mvarMe.txtFiltrar.Text = ""
   End If
End Sub

Private Sub mvarMe_txtFiltrarKeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      mvarMe.GrdVendas.FilterText = mvarMe.txtFiltrar.Text
      mvarMe.GrdVendas.Populate
   End If
End Sub

Private Sub mvarMe_txtFiltrarLostFocus()
   
   If mvarMe.txtFiltrar.Text = "" Then mvarMe.txtFiltrar.Text = mvarMe.txtFiltrar.Tag
   mvarMe.txtFiltrar.ForeColor = 12632256
End Sub
Private Sub PopulaGrdVendas(Optional bRefresh As Boolean)
   Dim i As Long
   Dim nRow As Long
   Dim Sql  As String
      
   With mvarMe.GrdVendas
      .SetImageList mvarMe.ImageList1
      If bRefresh Then
         nRow = .SelectedRows(0).Record.Item(0).Value
      End If
      If mvarConsulta = "" Then
         Sql = "Select Right('000000'+Rtrim(Cast(V.IDVENDA As Char)),6) [Nº Venda]" & vbNewLine
         Sql = Sql & ", Cast(Convert(Char(10),V.DTVENDA, 103) As SmallDatetime) [Data]" & vbNewLine
         Sql = Sql & ", Right('000000'+Rtrim(Cast(R.IDRECIBO As Char)),6) [Nº Recibo]" & vbNewLine
         Sql = Sql & ", Right('000000'+Rtrim(Cast(N.NUMNOTA As Char)),6) [Nº N.F.]" & vbNewLine
         'Sql = Sql & ", V.VLVENDA-V.VLDESC [Valor]"
         Sql = Sql & ", V.VLVENDA [Valor]"
         Sql = Sql & ", C.IDCLIENTE [IDCLIENTE], C.NOME [Cliente]" & vbNewLine
         
            Sql = Sql & " , Cast(( IsNull((SELECT SUM(I.QTDVENDA)" & vbNewLine
            Sql = Sql & "     FROM CITENSVENDA I" & vbNewLine
            Sql = Sql & "     JOIN SPRODUTO P ON I.IDLOJA=P.IDLOJA And I.IDPROD=P.IDPROD And P.ESERVICO=1 And P.EVENDA=1" & vbNewLine
            Sql = Sql & "     Where I.IDLOJA = V.IDLOJA" & vbNewLine
            Sql = Sql & "     And I.IDVENDA=V.IDVENDA),0)" & vbNewLine
            Sql = Sql & "   -" & vbNewLine
            Sql = Sql & "     IsNull((SELECT COUNT(S.IDSESSAO)" & vbNewLine
            Sql = Sql & "     From OSESSAO S" & vbNewLine
            Sql = Sql & "     Where S.IDVENDA = V.IDVENDA" & vbNewLine
            Sql = Sql & "     And isNull(IDITEM,0)<>0" & vbNewLine
            Sql = Sql & "     And S.IDTPSERVICO<>1),0)" & vbNewLine
            Sql = Sql & "   ) as Integer) [Saldo Serv.]" & vbNewLine
                  
         Sql = Sql & " From CVENDA V" & vbNewLine
         Sql = Sql & " Join OCLIENTE C On C.IDLOJA=V.IDLOJA And V.IDCLIENTE=C.IDCLIENTE" & vbNewLine
         Sql = Sql & " Left Join FRECIBO R On V.IDLOJA=R.IDLOJA And V.IDVENDA=R.IDVENDA" & vbNewLine
         Sql = Sql & " Left Join FNOTAFISCAL N On V.IDLOJA=N.IDLOJA And V.IDVENDA=N.IDVENDA And N.FLGCANCELADA=0" & vbNewLine
         Sql = Sql & " Where V.IDLOJA=" & SqlNum(mvarIDLOJA) & vbNewLine
         'Sql = Sql & " And V.SITVENDA='10'" & VBNEWLINE
         'Sql = Sql & " And V.VLVENDA-V.VLDESC>0" & VBNEWLINE
         Sql = Sql & " Order By Data Desc, V.IDVENDA Desc"
      Else
         Sql = mvarConsulta
      End If
            
      If mvarSys.xDb.AbreTabela(Sql, RsCad) Then
         Call FillRCFromRS(RsCad, mvarMe.GrdVendas, pCollColumn:=iCampo, pCurrency:=True)
      End If
      
      'For i = 0 To .Records.Count - 1
      '   .Records(i).Item(1).Icon = 0
      'Next
      
      If mvarMe.txtFiltrar.Text = "" Or mvarMe.txtFiltrar.Text = mvarMe.txtFiltrar.Tag Then
         .FilterText = ""
      Else
         .FilterText = mvarMe.txtFiltrar.Text
      End If
      
      
      .Populate
      If bRefresh Then
         For i = 0 To .Rows.Count - 1
            If nRow = .Rows(i).Record.Item(0).Value Then
               nRow = i
               i = .Rows.Count
            End If
         Next
         If nRow > .Rows.Count - 1 Then
            nRow = 0 '.Rows.Count - 1
         End If
         .SetFocus
         .SelectedRows.DeleteAll
         .Rows(nRow).EnsureVisible
         .Rows(nRow).Selected = True
      End If

'      Call DefineGrdIcones
'      Call DefineGrdFiltro
'      Call CalculaSaldo
      Call DefineGrdAlinha
      
      
      .Redraw
   End With

End Sub
Private Sub DefineGrdAlinha()
   Dim vCol As ReportColumn
   If iCampo Is Nothing Then Exit Sub
   With mvarMe.GrdVendas
      For Each vCol In .Columns
         vCol.HeaderAlignment = xtpAlignmentCenter
         'vCol.Sortable = False
      Next
      If iCampo.Count <> 0 Then
         .Columns(iCampo("Nº Venda")).Alignment = xtpAlignmentCenter
         .Columns(iCampo("Nº Recibo")).Alignment = xtpAlignmentCenter
         .Columns(iCampo("Nº N.F.")).Alignment = xtpAlignmentCenter
         .Columns(iCampo("Data")).Alignment = xtpAlignmentCenter
         .Columns(iCampo("Valor")).Alignment = xtpAlignmentRight
         .Columns(iCampo("Cliente")).Alignment = xtpAlignmentLeft
         .Columns(iCampo("Saldo Serv.")).Alignment = xtpAlignmentCenter
         
         .AutoColumnSizing = False
         .Columns(iCampo("Nº Venda")).Width = 60
         .Columns(iCampo("Nº Recibo")).Width = 60
         .Columns(iCampo("Nº N.F.")).Width = 60
         .Columns(iCampo("Data")).Width = 90
         .Columns(iCampo("Valor")).Width = 100
         .Columns(iCampo("IDCLIENTE")).Visible = False
         .Columns(iCampo("Cliente")).Width = (.Width / Screen.TwipsPerPixelX) - 271 - 240
         .Columns(iCampo("Saldo Serv.")).Width = 80
         '.Populate
      End If
   End With
End Sub

Private Sub MontarToolbar()
   Dim MyToolBar As CommandBar
   
   With mvarMe.CommandBars
'      .DeleteAll
      .EnableCustomization False
      Set .Icons = mvarMe.ImgToobar.Icons
      Set MyToolBar = mvarMe.CommandBars.ActiveMenuBar
      With MyToolBar
         .EnableDocking xtpFlagStretched
         '.DefaultButtonStyle = xtpButtonAutomatic '= xtpButtonIcon
         '.Customizable = True
         '.Closeable = False
         '.Position = xtpBarTop
         '.ShowExpandButton = False
         '.ShowGripper = True
         '.ShowTextBelowIcons = False
         '.ContextMenuPresent = True
            
         Call CriarButtonToolbar(MyToolBar, XTPControlType.xtpControlButton, eIconId.Novo, pIconId:=10001, pCaption:="Novo")
         Call CriarButtonToolbar(MyToolBar, XTPControlType.xtpControlButton, eIconId.Editar, pIconId:=10002, pCaption:="Editar")
         Call CriarButtonToolbar(MyToolBar, XTPControlType.xtpControlButton, eIconId.RPS, pIconId:=10003, pCaption:="RPS", pBeginGroup:=True)
         Call CriarButtonToolbar(MyToolBar, XTPControlType.xtpControlButton, eIconId.NF, pIconId:=10004, pCaption:="N.F.")
         Call CriarButtonToolbar(MyToolBar, XTPControlType.xtpControlButton, eIconId.ExpRPS, pIconId:=10005, pCaption:="Exportar RPS", pBeginGroup:=True)
         Call CriarButtonToolbar(MyToolBar, XTPControlType.xtpControlButton, eIconId.ImpNF, pIconId:=10006, pCaption:="Importar NF")
         Call CriarButtonToolbar(MyToolBar, XTPControlType.xtpControlButton, eIconId.Sair, pIconId:=10007, pCaption:="Sair", pBeginGroup:=True)
         
         Dim n As Object
         For Each n In .Controls
            n.Category = "BAR_VENDA"
         Next
         
         .Position = 0
         .Visible = True
      End With
   End With
End Sub
Private Sub NovaVenda()
   Dim MyVenda As TL_Venda
   
   Set MyVenda = New TL_Venda
   With MyVenda
      Set .Sys = mvarSys
      .Show vbModal
   End With
   Set MyVenda = Nothing
End Sub
Private Sub EditarVenda(pIDVENDA As Long)
   Dim MyVenda As TL_Venda
   
   Screen.MousePointer = vbHourglass
   Set MyVenda = New TL_Venda
   With MyVenda
      Set .Sys = mvarSys
      .IDLOJA = mvarIDLOJA
      .IDVENDA = pIDVENDA
      Set .TlVendas = Me
      .Show vbModal
   End With
   Call PopulaGrdVendas(True)
   Screen.MousePointer = vbDefault
   Exit Sub
Saida:
   Screen.MousePointer = vbDefault
   ShowError
End Sub
Private Sub ExportarRPS()
   Dim TLExp As TL_ExportRPS

   Set TLExp = New TL_ExportRPS
   With TLExp
      Set .Sys = mvarSys
      .Show
   End With
End Sub
Private Sub BaixarSaldoV()
   Dim TLBaixar As TL_BaixarSaldoV

   Set TLBaixar = New TL_BaixarSaldoV
   With TLBaixar
      Set .Sys = mvarSys
      .IDVENDA = mvarMe.GrdVendas.SelectedRows(0).Record(iCampo("Nº VENDA")).Value
      .Show
      mvarMe.GrdVendas.SelectedRows(0).Record(iCampo("Saldo Serv.")).Value = .SaldoNovo
      mvarMe.GrdVendas.Populate
   End With
End Sub
Private Sub ImportarNF()
   Dim NgFin As NG_Financeiro
   Dim sArq  As String
   Dim ODlg  As Object
     
   On Error GoTo TrataErro
   Set ODlg = CreateObject("MSComDlg.CommonDialog")
   sArq = SelecionarArquivo(ODlg, "Notas Fiscais", "", "*.xml|*.xml|*.txt|*.txt|*.csv|*.csv")
   If ExisteArquivo(sArq) Then
      Set NgFin = New NG_Financeiro
      Set NgFin.Sys = mvarSys
      Call NgFin.ImportarNF(sArq)
   End If
   Exit Sub
TrataErro:
   If Err = 429 Then
      Set ODlg = mvarMe.CommonDialog1
      Resume Next
   End If
   ExibirStop Error, Err
End Sub
Public Sub Refresh()
   Call PopulaGrdVendas
End Sub

