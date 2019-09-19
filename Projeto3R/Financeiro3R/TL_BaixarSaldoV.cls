VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_BaixarSaldoV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe  As FrmBaixarSaldoV
Attribute mvarMe.VB_VarHelpID = -1
Private mvarSys      As Object
Private mvarIDLOJA   As Integer
Private mvarIDVENDA  As Integer
Private mvarSaldoNovo As Integer
Public Property Get SaldoNovo() As Integer
   SaldoNovo = mvarSaldoNovo
End Property
Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
   mvarIDLOJA = Val(mvarSys.Propriedades("IDLOJA"))
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Public Property Let IDVENDA(ByVal vData As Long)
   mvarIDVENDA = vData
End Property
Public Property Get IDVENDA() As Long
   IDVENDA = mvarIDVENDA
End Property
Public Property Let IDLOJA(ByVal vData As String)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As String
   IDLOJA = mvarIDLOJA
End Property
Public Sub Show()
   mvarMe.Show vbModal
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmBaixarSaldoV
End Sub
Private Sub mvarMe_Activate()
   If GetTag(mvarMe, "1VEZ", 1) = 1 Then
      Call SetTag(mvarMe, "1VEZ", 0)
      If mvarIDVENDA = 0 Then
         mvarMe.TxtIDVENDA.SetFocus
      Else
         mvarMe.TxtSaldoNovo.Text = mvarMe.TxtSaldo.Text
         mvarMe.TxtSaldoNovo.SetFocus
      End If
      
      'Call MontaTela
      'Call PopulaTela
   End If
End Sub
Private Sub mvarMe_CmdSairClick()
   mvarSaldoNovo = xVal(mvarMe.TxtSaldo.Text)
   Unload mvarMe
End Sub
Private Sub mvarMe_CmdOkClick()
   Dim NgFin As NG_Financeiro
   Dim bOk As Boolean
   bOk = False
   If xVal(mvarMe.TxtSaldo.Text) > xVal(mvarMe.TxtSaldoNovo.Text) Then
      bOk = True
   ElseIf (xVal(mvarMe.TxtSaldo.Text) = xVal(mvarMe.TxtSaldoNovo.Text)) And Mid(mvarMe.TxtSaldoNovo.Text, 1, 1) = "0" Then
      bOk = True
   End If
   If bOk Then
      Set NgFin = New NG_Financeiro
      Set NgFin.Sys = mvarSys
      Call NgFin.BaixarSaldo(mvarMe.TxtIDVENDA, xVal(mvarMe.TxtSaldoNovo.Text))
      Set NgFin = Nothing
      mvarSaldoNovo = xVal(mvarMe.TxtSaldoNovo.Text)
   Else
      mvarSaldoNovo = xVal(mvarMe.TxtSaldo.Text)
   End If
   Unload mvarMe
End Sub
Private Sub mvarMe_Load()
   DoEvents
   Call MontaTela
   Screen.MousePointer = vbDefault
End Sub
Private Sub MontaTela()
   Dim NgFin As NG_Financeiro
   Dim Sql As String
   Dim MyRs As Object
   
   Set NgFin = New NG_Financeiro
   Set NgFin.Sys = mvarSys
   Sql = NgFin.GetSqlVendaSaldo(mvarIDLOJA, mvarIDVENDA)
   
   If mvarSys.xDb.AbreTabela(Sql, MyRs) Then
      mvarMe.TxtIDVENDA.Text = MyRs("Venda")
      mvarMe.TxtDTVENDA.Text = MyRs("Data")
      mvarMe.TxtIDCLIENTE.Text = MyRs("IDCLIENTE")
      mvarMe.TxtNOME.Text = MyRs("Nome")
      mvarMe.TxtSaldo.Text = MyRs("Sld")
      mvarMe.TxtSaldoNovo.Text = MyRs("Sld")
   Else
       mvarIDVENDA = 0
       Call LimparTela(mvarMe)
       mvarMe.TxtIDVENDA.SetFocus
   End If
End Sub
Private Sub mvarMe_TxtIDVENDALostFocus()
   mvarIDVENDA = xVal(mvarMe.TxtIDVENDA.Text)
   Call MontaTela
End Sub

Private Sub mvarMe_TxtSaldoNovoLostFocus()
   If Not (xVal(mvarMe.TxtSaldo) >= xVal(mvarMe.TxtSaldoNovo)) Or xVal(mvarMe.TxtSaldoNovo) < 0 Then
      Call ExibirInformacao("Saldo inválido!" & vbNewLine & "O novo saldo deve ser menor ou igual ao atual")
      mvarMe.TxtSaldoNovo.Text = mvarMe.TxtSaldo.Text
      mvarMe.TxtSaldoNovo.SetFocus
   End If
End Sub
