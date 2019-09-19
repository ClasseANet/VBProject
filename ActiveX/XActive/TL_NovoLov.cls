VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_NovoLov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe  As frmNovoLov
Attribute mvarMe.VB_VarHelpID = -1

Private mvarSys As Object

Private mvarSql As String
Private mvarRsItens As Object
Private mvarFiltro As String
Private mvarIdSelecionado As Integer
Public Property Set Sys(ByVal vData As Object)
    Set mvarSys = vData
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Public Sub Show(Optional pSql, Optional pRsItens)
   
   If Trim(mvarFiltro) <> "" Then
      mvarMe.TxtFiltrar.Text = mvarFiltro
   End If
   
   If Not pRsItens Is Nothing Then
      Set mvarRsItens = pRsItens.Clone
   Else
      If Trim(pSql) <> "" Then
         ' executa a query
         mvarSql = pSql
         Call mvarSys.XDb.AbreTabela(mvarSql, mvarRsItens)
      End If
   End If
   Call FillRCFromRS(mvarRsItens, mvarMe.grdItens)
   Call mvarMe_CmdFiltrarClick
      
   If Not mvarMe.Visible Then
      mvarMe.Show vbModal
   End If
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New frmNovoLov
End Sub
Private Sub Class_Terminate()
   On Error Resume Next
   Unload mvarMe
End Sub

Public Sub ShowClientes(Optional pSql, Optional pIDLOJA)
   Dim Sql As String
   Dim MyRs As Object
   Dim sAtivo As String
   
   mvarMe.Caption = "Listagem de Clientes"
   
   If IsMissing(pSql) Then
      If IsMissing(pIDLOJA) Then
         pIDLOJA = mvarSys.Propriedades("IDLOJA")
      End If
      
      sAtivo = ""
      sAtivo = sAtivo & IIf(mvarMe.ChkAtivo.Value, "1", "")
      sAtivo = sAtivo & IIf(sAtivo <> "" And mvarMe.ChkInativo.Value, ", ", "")
      sAtivo = sAtivo & IIf(mvarMe.ChkInativo.Value, "0", "")
      sAtivo = sAtivo & IIf(sAtivo <> "" And mvarMe.ChkEmEspera.Value, ", ", "")
      sAtivo = sAtivo & IIf(mvarMe.ChkEmEspera.Value, "2", "")
      
      Sql = "Select IDCLIENTE [Id], NOME [Nome], TEL1 [Tel. Residencial], TEL2 [Tel. Celular]"
      Sql = Sql & " From OCLIENTE "
      Sql = Sql & " Where IDLOJA=" & pIDLOJA
      Sql = Sql & " And ATIVO in (" & sAtivo & ")"
      Sql = Sql & " Order By NOME"
   Else
      Sql = pSql
   End If
   If mvarSys.XDb.AbreTabela(Sql, MyRs) Then
      Call Show(pRsItens:=MyRs)
   End If
End Sub

Private Sub mvarMe_Activate()
   If mvarMe.TxtFiltrar.Visible And mvarMe.TxtFiltrar.Enabled Then
      mvarMe.TxtFiltrar.SetFocus
      mvarMe.TxtFiltrar.SelStart = 0
      mvarMe.TxtFiltrar.SelLength = Len(mvarMe.TxtFiltrar.Text)
   ElseIf mvarMe.grdItens.Visible And mvarMe.grdItens.Enabled Then
      mvarMe.grdItens.SetFocus
   End If
End Sub

Private Sub mvarMe_ChkAtivoClick()
   Call ShowClientes
End Sub
Private Sub mvarMe_ChkEmEsperaClick()
   Call ShowClientes
End Sub
Private Sub mvarMe_ChkInativoClick()
   Call ShowClientes
End Sub
Private Sub mvarMe_CmdFiltrarClick()
   mvarMe.grdItens.FilterText = mvarMe.TxtFiltrar.Text
   mvarMe.grdItens.Populate
   If mvarMe.grdItens.Visible And mvarMe.grdItens.Enabled Then
      mvarMe.grdItens.SetFocus
   End If
End Sub
Private Sub mvarMe_CmdSairClick()
   mvarIdSelecionado = 0
   Unload mvarMe
End Sub
Private Sub mvarMe_cmdSelecionarClick()
   Dim bSel As Boolean
   If mvarMe.ActiveControl Is mvarMe.TxtFiltrar Then
      bSel = (mvarMe.grdItens.Rows.Count = 1)
   Else
      bSel = True
   End If
   If bSel Then
      mvarIdSelecionado = xVal(mvarMe.grdItens.SelectedRows(0).Record(0).Value)
      Unload mvarMe
   End If
End Sub

Private Sub mvarMe_FormKeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      If mvarMe.ActiveControl Is mvarMe.TxtFiltrar Or mvarMe.ActiveControl Is mvarMe.cmdFiltrar Then
         Call mvarMe_CmdFiltrarClick
      ElseIf Not mvarMe.ActiveControl Is mvarMe.cmdSair Then
         Call mvarMe_cmdSelecionarClick
      Else
         Call mvarMe_CmdFiltrarClick
      End If
   End If
End Sub

Private Sub mvarMe_FormKeyUp(KeyCode As Integer, Shift As Integer)
'mvarMe.Caption = KeyCode
End Sub

Private Sub mvarMe_grdItensKeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      Call mvarMe_cmdSelecionarClick
   End If
End Sub

Private Sub mvarMe_Load()
   mvarMe.Icon = mvarSys.MDI.Icon
   mvarMe.KeyPreview = True
End Sub
Private Sub mvarMe_Resize()
   On Error Resume Next
   mvarMe.lblFiltrar.Top = mvarMe.ScaleHeight - 500
   mvarMe.TxtFiltrar.Top = mvarMe.lblFiltrar.Top
   mvarMe.cmdFiltrar.Top = mvarMe.lblFiltrar.Top
   mvarMe.cmdSelecionar.Top = mvarMe.lblFiltrar.Top
   mvarMe.cmdSair.Top = mvarMe.lblFiltrar.Top

   mvarMe.cmdSelecionar.Left = mvarMe.ScaleWidth - 3000
   mvarMe.cmdSair.Left = mvarMe.cmdSelecionar.Left + 1500


   mvarMe.grdItens.Width = mvarMe.ScaleWidth - 300
   mvarMe.grdItens.Height = mvarMe.ScaleHeight - 1000
   mvarMe.ChkAtivo.Top = mvarMe.grdItens.Top + mvarMe.grdItens.Height + 30
   mvarMe.ChkInativo.Top = mvarMe.ChkAtivo.Top
   mvarMe.ChkEmEspera.Top = mvarMe.ChkAtivo.Top

End Sub
Public Property Get IdSelecionado() As Integer
   IdSelecionado = mvarIdSelecionado
End Property
Public Property Let IdSelecionado(ByVal vNewValue As Integer)
   mvarIdSelecionado = vNewValue
End Property
Public Property Get Filtro() As String
   Filtro = mvarFiltro
End Property
Public Property Let Filtro(ByVal vNewValue As String)
   mvarFiltro = vNewValue
End Property
Private Sub mvarMe_TxtFiltrarKeyDown(KeyCode As Integer, Shift As Integer)
'   If KeyCode = vbKeyReturn Then
'      Call mvarMe_CmdFiltrarClick
'   End If
End Sub
Private Sub mvarMe_txtFiltrarKeyPress(KeyAscii As Integer)
'   If KeyAscii = vbKeyReturn Then
'      Call mvarMe_CmdFiltrarClick
'   End If
End Sub
Private Sub mvarMe_TxtFiltrarKeyUp(KeyCode As Integer, Shift As Integer)
   mvarFiltro = mvarMe.TxtFiltrar.Text
'   If KeyCode = vbKeyReturn Then
'      Call mvarMe_CmdFiltrarClick
'   End If
End Sub
