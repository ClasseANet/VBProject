VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_VendaSaldo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe    As FrmVendaSaldo
Attribute mvarMe.VB_VarHelpID = -1

Dim nLenIDVenda   As Integer
Dim dToday        As Date

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
   nLenIDVenda = mvarSys.Propriedades("LENIDVENDA")
   On Error Resume Next
   dToday = CDate(Format(mvarSys.xDb.SysDate, "DD/MM/YYYY"))
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Private Sub Class_Initialize()
   Set mvarMe = New FrmVendaSaldo
End Sub
Private Sub Class_Terminate()
   Set mvarMe = Nothing
End Sub
Public Sub Show(Optional pMode As FormShowConstants = vbModal)
'   If mvarIDLOJA = 0 Then
'      Call ExibirAviso("Código da Loja Inválido! (""" & mvarIDLOJA & """)")
'      Exit Sub
'   End If
   
   mvarMe.Show pMode
End Sub
