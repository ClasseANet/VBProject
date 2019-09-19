VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VARIAVEL"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarTIPO As String
Private mvarNOME As String
Private mvarESCOPO As String
Private mvarVALOR As String
Private mvarPassArg As String
Public Property Let PassArg(ByVal vData As String)
    mvarPassArg = vData
End Property
Public Property Get PassArg() As String
    PassArg = mvarPassArg
End Property
Public Property Let VALOR(ByVal vData As String)
    mvarVALOR = vData
End Property
Public Property Get VALOR() As String
    VALOR = mvarVALOR
End Property
Public Property Let ESCOPO(ByVal vData As String)
    mvarESCOPO = vData
End Property
Public Property Get ESCOPO() As String
    ESCOPO = mvarESCOPO
End Property
Public Property Let NOME(ByVal vData As String)
    mvarNOME = vData
End Property
Public Property Get NOME() As String
    NOME = mvarNOME
End Property
Public Property Let Tipo(ByVal vData As String)
    mvarTIPO = vData
End Property
Public Property Get Tipo() As String
    Tipo = mvarTIPO
End Property
