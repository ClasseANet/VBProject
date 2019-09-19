VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_CADSESTOQUE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private WithEvents mvarMe     As FrmCADSESTOQUE
Attribute mvarMe.VB_VarHelpID = -1
Private mvarSys      As Object
Private mvarIDLOJA   As Integer
Private mvarIDCONTA  As Integer
Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
   mvarIDLOJA = mvarSys.Propriedades("IDLOJA")
   mvarIDCONTA = mvarSys.Propriedades("IDCONTA")
End Property
Public Property Get Sys() As Object
   Set Sys = mvarSys
End Property
Public Property Let IDLOJA(ByVal vData As Integer)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Integer
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDCONTA(ByVal vData As Integer)
   mvarIDCONTA = vData
End Property
Public Property Get IDCONTA() As Integer
   IDCONTA = mvarIDCONTA
End Property
Private Sub Class_Initialize()
   Set mvarMe = New FrmCADSESTOQUE
   mvarIDLOJA = 1
End Sub
Public Sub Show()
   If mvarIDCONTA = 0 Then Call ExibirAviso("Não existe conta corrente definida.")
   
   '* Carregar e Acoplar Form no MDI
   Load mvarMe
   Call AcoplarForm(mvarMe, 2, mvarSys)
   mvarMe.Show
End Sub
Public Sub Hide()
   mvarMe.Hide
End Sub
Private Sub mvarMe_Activate()
   'Call PopulaTela
   If GetTag(mvarMe, "LOAD", 1) = 1 Then
      Call SetTag(mvarMe, "LOAD", 0)
   End If
   Screen.MousePointer = vbDefault
End Sub
