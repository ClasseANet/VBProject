VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLS_SALA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mvarSys         As Object
Private mvarIDLOJA      As Integer
Private mvarIDSALA      As Integer
Private mvarNMSALA      As String
Private mvarScheduleID    As Long
Private mvarSelecionada As Boolean
'Private mvarTbSala      As Object
Public Property Let Selecionada(ByVal vData As Boolean)
   mvarSelecionada = vData
End Property
Public Property Get Selecionada() As Boolean
   Selecionada = mvarSelecionada
End Property
Public Property Let ScheduleID(ByVal vData As Long)
   mvarScheduleID = vData
End Property
Public Property Get ScheduleID() As Long
   ScheduleID = mvarScheduleID
End Property
'Public Property Set TbSala(ByVal vData As Object)
'   Set mvarTbSala = vData
'End Property
'Public Property Get TbSala() As Object
'    Set TbSala = mvarTbSala
'End Property
Public Property Let NMSALA(ByVal vData As String)
   mvarNMSALA = vData
End Property
Public Property Get NMSALA() As String
   NMSALA = mvarNMSALA
End Property
Public Property Let IDSALA(ByVal vData As Integer)
'   If mvarTbSala.IDSALA <> vData Then
'      Set mvarTbSala = CriarObjeto("BANCO_3R.TB_OSALA")
'      Set mvarTbSala.xDb = mvarSys.xDb
'      Call mvarTbSala.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDSALA:=vData)
'   End If
   mvarIDSALA = vData
End Property
Public Property Get IDSALA() As Integer
   IDSALA = mvarIDSALA
End Property
Public Property Let IDLOJA(ByVal vData As Integer)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Integer
   IDLOJA = mvarIDLOJA
End Property
Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Private Sub Class_Initialize()
'   Set mvarTbSala = CriarObjeto("BANCO_3R.TB_OSALA")
End Sub
Private Sub Class_Terminate()
'   Set mvarTbSala = Nothing
End Sub
