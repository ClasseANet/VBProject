VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NG_Suprimento"
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
   Dim oTb As Object
   
   Set mvarSys = vData
   
   If Val(mvarSys.Propriedades("IDCONTA")) = 0 Then
      If DefineContaLAN(oTb, 1, mvarSys) Then
         mvarSys.Propriedades("IDCONTA") = oTb.IDCONTA
      End If
      Set oTb = Nothing
   End If
   If Val(mvarSys.Propriedades("IDEST")) = 0 Then
      Set oTb = CriarObjeto("BANCO_3R.TB_SESTOQUE")
      Set oTb.xdb = mvarSys.xdb
      If oTb.Pesquisar(Ch_IDLOJA:=mvarSys.IDLOJA, Ch_OrderBy:="IDEST") Then
         mvarSys.Propriedades("IDEST") = oTb.IDEST
      End If
      Set oTb = Nothing
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
Public Sub Show(Optional Reload As Boolean = False)
   mvarTLPane.Show Reload
End Sub
Private Sub Class_Initialize()
   Set mvarTLPane = New TL_PaneSupri
   Set mvarPane = mvarTLPane.Pane
End Sub
