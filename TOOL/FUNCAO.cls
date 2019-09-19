VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FUNCAO"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Private mvarLINHAS As String 'local copy
Private mvarTIPO As String 'local copy
Private mvarNOME As String 'local copy
Private mvarisFuncao As Boolean 'local copy
Private mvarESCOPO As String 'local copy
Private mvarPARAM As New Collection 'local copy
Private mvarFUNCOES As New Collection 'local copy
Private mvarVARIAVEIS As New Collection 'local copy
Public Property Set VARIAVEIS(ByVal vData As Collection)
    Set mvarVARIAVEIS = vData
End Property
Public Property Get VARIAVEIS() As Collection
    Set VARIAVEIS = mvarVARIAVEIS
End Property
Public Property Set FUNCOES(ByVal vData As Collection)
    Set mvarFUNCOES = vData
End Property
Public Property Get FUNCOES() As Collection
    Set FUNCOES = mvarFUNCOES
End Property
Public Property Set Param(ByVal vData As Object)
    Set mvarPARAM = vData
End Property
Public Property Get Param() As Collection
    Set Param = mvarPARAM
End Property
Public Property Let ESCOPO(ByVal vData As String)
    mvarESCOPO = vData
End Property
Public Property Get ESCOPO() As String
    ESCOPO = mvarESCOPO
End Property
Public Property Let isFuncao(ByVal vData As Boolean)
    mvarisFuncao = vData
End Property
Public Property Get isFuncao() As Boolean
    isFuncao = mvarisFuncao
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
Public Property Let LINHAS(ByVal vData As String)
    mvarLINHAS = vData
End Property
Public Property Get LINHAS() As String
    LINHAS = mvarLINHAS
End Property
