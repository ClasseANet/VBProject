VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFUNCAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"clsVARIAVEL"
Private mvarVARIAVEL As clsVARIAVEL
Private mvarPARAMETROS As New Collection
Private mvarFUNCOES As New Collection
Private mvarVARIAVEIS As New Collection

Private mvarLINHAS As Integer
Private mvarTipo As String
Private mvarNOME As String
Private mvarisFuncao As Boolean
Private mvarEscopo As String
Private mvarPagProp As PROPPAGE
Private mvarStartLine As Integer
Private mvarFileName As String
Public Property Let FileName(ByVal vData As String)
    mvarFileName = vData
End Property
Public Property Get FileName() As String
    FileName = mvarFileName
End Property
Public Property Let StartLine(ByVal vData As Integer)
    mvarStartLine = vData
End Property
Public Property Get StartLine() As Integer
    StartLine = mvarStartLine
End Property
Public Property Get PagProp() As PROPPAGE
    If mvarPagProp Is Nothing Then
        Set mvarPagProp = New PROPPAGE
    End If
    Set PagProp = mvarPagProp
End Property
Public Property Set PagProp(vData As PROPPAGE)
    Set mvarPagProp = vData
End Property
Public Property Get VARIAVEL() As clsVARIAVEL
    If mvarVARIAVEL Is Nothing Then
        Set mvarVARIAVEL = New clsVARIAVEL
        mvarVARIAVEL.FileName = mvarFileName
    End If
    Set VARIAVEL = mvarVARIAVEL
End Property
Public Property Set VARIAVEL(vData As clsVARIAVEL)
    Set mvarVARIAVEL = vData
End Property
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
Public Property Set PARAMETROS(ByVal vData As Object)
    Set mvarPARAMETROS = vData
End Property
Public Property Get PARAMETROS() As Collection
    Set PARAMETROS = mvarPARAMETROS
End Property
Public Property Let Escopo(ByVal vData As String)
    mvarEscopo = vData
End Property
Public Property Get Escopo() As String
    Escopo = mvarEscopo
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
    mvarTipo = vData
End Property
Public Property Get Tipo() As String
    Tipo = mvarTipo
End Property
Public Property Let LINHAS(ByVal vData As Integer)
    mvarLINHAS = vData
End Property
Public Property Get LINHAS() As Integer
    LINHAS = mvarLINHAS
End Property
Private Sub Class_Initialize()
   Set mvarVARIAVEL = New clsVARIAVEL
End Sub
Private Sub Class_Terminate()
   Set mvarVARIAVEL = Nothing
   
   Set mvarPARAMETROS = Nothing
   Set mvarFUNCOES = Nothing
   Set mvarVARIAVEIS = Nothing
End Sub
