VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DbCampo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarNOME        As String
Private mvarDESCRICAO   As String
Private mvarTIPO        As String
Private mvarTAMANHO     As Long
Private mvarisSys       As Boolean
Private mvarisNull      As Boolean
Public Property Let TAMANHO(ByVal vData As Long)
    mvarTAMANHO = vData
End Property
Public Property Get TAMANHO() As Long
    TAMANHO = mvarTAMANHO
End Property
Public Property Let TIPO(ByVal vData As String)
    mvarTIPO = vData
End Property
Public Property Get TIPO() As String
    TIPO = mvarTIPO
End Property
Public Property Let NOME(ByVal vData As String)
    mvarNOME = vData
End Property
Public Property Get NOME() As String
    NOME = mvarNOME
End Property
Public Property Let DESCRICAO(ByVal vData As String)
    mvarDESCRICAO = vData
End Property
Public Property Get DESCRICAO() As String
    DESCRICAO = mvarDESCRICAO
End Property
Public Property Let isSys(ByVal vData As Boolean)
   mvarisSys = vData
End Property
Public Property Get isSys() As Boolean
   isSys = mvarisSys
End Property
Public Property Let isNull(ByVal vData As Boolean)
   mvarisNull = vData
End Property
Public Property Get isNull() As Boolean
   isNull = mvarisNull
End Property



