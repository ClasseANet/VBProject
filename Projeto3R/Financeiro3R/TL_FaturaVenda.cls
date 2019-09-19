VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_FaturaVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe    As FrmFaturaVenda
Attribute mvarMe.VB_VarHelpID = -1
Dim nLenID   As Integer
Dim dToday  As Date
Dim bFaturaPaga As Boolean

Private mvarSys            As Object
Private mvarIDLOJA         As Integer
Private mvarIDVENDA        As Long
Public Property Let IDVENDA(ByVal vData As Long)
    mvarIDVENDA = vData
End Property
Public Property Get IDVENDA() As Long
   IDVENDA = mvarIDVENDA
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
   nLenID = mvarSys.Propriedades("LENIDVENDA")
   On Error Resume Next
   dToday = CDate(Format(mvarSys.xDb.SysDate, "DD/MM/YYYY"))
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Private Sub Class_Initialize()
   Set mvarMe = New FrmFaturaVenda
End Sub
Private Sub Class_Terminate()
   Set mvarMe = Nothing
End Sub
Public Sub Show(Optional pMode As FormShowConstants = vbModal)
   mvarMe.Show pMode
End Sub
Private Sub mvarMe_Activate()
'   Call MontaTela
'   Call PopulaTela
   If GetTag(mvarMe, "1VEZ", "1") = "1" Then
      Call SetTag(mvarMe, "1VEZ", "0")
      Screen.MousePointer = vbDefault
   End If
End Sub
Private Sub mvarMe_CmdCancelClick()
   Unload mvarMe
End Sub
Private Sub mvarMe_CmdExcluirClick()
   Dim sMsg As String
   Dim Queries As Collection
   Dim Sql As String
   Dim nRecibo As Long
   Dim MyRs As Object
      
   Dim TBFatura As Object 'TB_FFATURA
   
   sMsg = "Esta operação irá excluir definitivamente a fatura '" & "'"
   sMsg = sMsg & vbNewLine & vbNewLine
   sMsg = sMsg & "Deseja continuar?"
   If ExibirPergunta(sMsg, , False) = vbYes Then
   
   End If
End Sub
Private Sub mvarMe_CmdOkClick()
   DoEvents
   'If Not VerificaFatura() Then Exit Sub
   If SalvarTela Then
      Unload mvarMe
   End If
End Sub
Private Function SalvarTela() As Boolean

End Function
Private Sub mvarMe_Load()
   mvarMe.Icon = mvarSys.MDI.Icon
   '* Monta Tela
   Call MontaTela
   Call PopulaTela
   
   
   Screen.MousePointer = vbDefault
End Sub
Private Sub MontaTela()
   With mvarMe
      .CmdExcluir.Enabled = False
   End With
End Sub
Private Sub LimparTela()
   With mvarMe
      mvarMe.GrdFaturas.Clear False
   End With
End Sub
Private Sub PopulaTela()

End Sub



