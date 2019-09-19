VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AreaTemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Private mvarTipoObj As String
Private mvarNomeObj As String
Private mvarTextoObj As String
Enum TpOper
   Copiar
   Recortar
End Enum
Private mvarTextoOper As TpOper
Public Property Let TipoOper(vData As TpOper)
   mvarTipoOper = vData
End Property
Public Property Get TipoOper() As TpOper
   TipoOper = mvarTipoOper
End Property
Public Property Let TipoObj(vData As String)
   mvarTipoObj = vData
End Property
Public Property Get TipoObj() As String
   TipoObj = mvarTipoObj
End Property
Public Property Let NomeObj(vData As String)
   mvarNomeObj = vData
End Property
Public Property Get NomeObj() As String
   NomeObj = mvarNomeObj
End Property
Public Property Let TextoObj(vData As String)
   mvarTextoObj = vData
End Property
Public Property Get TextoObj() As String
   TextoObj = mvarTextoObj
End Property
Private Sub Class_Initialize()
   mvarNomeObj = ""
   mvarTextoObj = ""
   mvarTipoObj = ""
   mvarTipoOper = Copiar
End Sub
