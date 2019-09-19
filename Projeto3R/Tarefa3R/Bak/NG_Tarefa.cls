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
Private mvarSys      As Object

Private mvarBarItem  As Object
Private mvarTLPane   As Object
Private mvarPane     As Object
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
Public Sub Show()
   mvarTLPane.Show
End Sub
Private Sub Class_Initialize()
   Set mvarTLPane = New TL_Tarefas
   Set mvarPane = mvarTLPane.Pane
End Sub
Public Sub ExibirVenda(pCod As String, Optional pCalControl)
   Dim TbVenda As Object
   Dim TBAtend As Object
   Dim TlVenda As Object
   
   Set TbVenda = CriarObjeto("BANCO_3R.TB_CVENDA")
   Set TbVenda.xDb = mvarSys.xDb
   If TbVenda.Pesquisar(Ch_IDLOJA:=mvarSys.Propriedades("IDLOJA"), Ch_IDVENDA:=pCod) Then
      Set TlVenda = CriarObjeto("Calendario3R.TL_Venda")
      With TlVenda
         Set TBAtend = CriarObjeto("BANCO_3R.TB_OATENDIMENTO_VENDA")
         Set TBAtend.xDb = mvarSys.xDb
         If TBAtend.Pesquisar(Ch_IDLOJA:=mvarSys.Propriedades("IDLOJA"), Ch_IDVENDA:=pCod) Then
            .IDATENDIMENTO = TBAtend.IDATENDIMENTO
         End If
         
         Set .Sys = mvarSys
         If Not IsMissing(pCalControl) Then
            Set .CalControl = pCalControl
         End If
         .IDLOJA = TbVenda.IDLOJA
         .IDCLIENTE = TbVenda.IDCLIENTE
         .IDVENDA = TbVenda.IDVENDA
         
         .Show vbModal
      End With
      Set TlVenda = Nothing
      Set TBAtend = Nothing
      Set TbVenda = Nothing
   End If
End Sub

