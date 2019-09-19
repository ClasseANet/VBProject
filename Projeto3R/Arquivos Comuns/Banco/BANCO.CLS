VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "BANCO"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0", "TB_FLOTERPS"
Attribute VB_Ext_KEY = "Member1", "TB_FNOTAFISCAL"
Attribute VB_Ext_KEY = "Member2", "TB_FRECIBO"
Attribute VB_Ext_KEY = "Member3", "TB_OCLIENTE"
Option Explicit 
Private mvarXDb As Object 
Private mvarTB_FNOTAFISCAL As TB_FNOTAFISCAL
Private mvarTB_FRECIBO As TB_FRECIBO
Private mvarTB_OCLIENTE As TB_OCLIENTE
Public Property Get TB_OCLIENTE() As TB_OCLIENTE
   If mvarTB_OCLIENTE Is Nothing Then
      Set mvarTB_OCLIENTE = New TB_OCLIENTE
      mvarTB_OCLIENTE.XDb = mvarXDb
   End If
   Set TB_OCLIENTE = mvarTB_OCLIENTE
End Property
Public Property Set TB_OCLIENTE(vData As TB_OCLIENTE)
   Set mvarTB_OCLIENTE = vData
End Property
Public Property Get TB_FRECIBO() As TB_FRECIBO
   If mvarTB_FRECIBO Is Nothing Then
      Set mvarTB_FRECIBO = New TB_FRECIBO
      mvarTB_FRECIBO.XDb = mvarXDb
   End If
   Set TB_FRECIBO = mvarTB_FRECIBO
End Property
Public Property Set TB_FRECIBO(vData As TB_FRECIBO)
   Set mvarTB_FRECIBO = vData
End Property
Public Property Get TB_FNOTAFISCAL() As TB_FNOTAFISCAL
   If mvarTB_FNOTAFISCAL Is Nothing Then
      Set mvarTB_FNOTAFISCAL = New TB_FNOTAFISCAL
      mvarTB_FNOTAFISCAL.XDb = mvarXDb
   End If
   Set TB_FNOTAFISCAL = mvarTB_FNOTAFISCAL
End Property
Public Property Set TB_FNOTAFISCAL(vData As TB_FNOTAFISCAL)
   Set mvarTB_FNOTAFISCAL = vData
End Property
 
Private mvarTB_FLOTERPS As TB_FLOTERPS
Public Property Get TB_FLOTERPS() As TB_FLOTERPS
   If mvarTB_FLOTERPS Is Nothing Then
      Set mvarTB_FLOTERPS = New TB_FLOTERPS
      mvarTB_FLOTERPS.XDb = mvarXDb
   End If
   Set TB_FLOTERPS = mvarTB_FLOTERPS
End Property
Public Property Set TB_FLOTERPS(vData As TB_FLOTERPS)
   Set mvarTB_FLOTERPS = vData
End Property
Public Property Set XDb(ByVal vData As Object)
   Set mvarXDb = vData
End Property
Public Property Let XDb(ByVal vData As Object)
   Set mvarXDb = vData
End Property
Public Property Get XDb() As Object
   Set XDb = mvarXDb
End Property
Private Sub Class_Terminate()
   Set mvarTB_OCLIENTE = Nothing
   Set mvarTB_FRECIBO = Nothing
   Set mvarTB_FNOTAFISCAL = Nothing
  Set mvarTB_FLOTERPS = Nothing
End Sub
