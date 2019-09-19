VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMEMBRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"clsCONTROLE"
Attribute VB_Ext_KEY = "Member1" ,"clsEVENTO"
Attribute VB_Ext_KEY = "Member2" ,"clsMETODO"
Attribute VB_Ext_KEY = "Member3" ,"clsPROPRIEDADE"
Attribute VB_Ext_KEY = "Member4" ,"clsVARIAVEL"
Private mvarEVENTO As clsEVENTO
Private mvarMETODO As clsMETODO
Private mvarVARIAVEL As clsVARIAVEL
Private mvarPROPRIEDADE As clsPROPRIEDADE
Private mvarFUNCAO As New clsFUNCAO
Private mvarCONTROLE As clsCONTROLE
Private mvarLIB As clsLIB

Private mvarEVENTOS As New Collection
Private mvarMETODOS As New Collection
Private mvarVARIAVEIS As New Collection
Private mvarPROPRIEDADES As New Collection
Private mvarFUNCOES As New Collection
Private mvarCONTROLES As New Collection
Private mvarLIBS As New Collection

Private mvarLINHAS As Integer
Private mvarNOME As String
Private mvarFileName As String
Private mvarTAMANHO As String
Private mvarPATH As String

Private mvarFlag As Boolean
Private mvarMDI As Boolean

Private mvarCarregado As Boolean
Private mvarPagProp As PROPPAGE
Private mvarStartLine As Integer
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
Public Property Let Carregado(ByVal vData As Boolean)
    mvarCarregado = vData
End Property
Public Property Get Carregado() As Boolean
    Carregado = mvarCarregado
End Property
Public Property Get Propriedade() As clsPROPRIEDADE
    If mvarPROPRIEDADE Is Nothing Then
        Set mvarPROPRIEDADE = New clsPROPRIEDADE
        mvarPROPRIEDADE.FileName = mvarFileName
    End If
    Set Propriedade = mvarPROPRIEDADE
End Property
Public Property Set Propriedade(vData As clsPROPRIEDADE)
    Set mvarPROPRIEDADE = vData
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
Public Property Get LIB() As clsLIB
    If mvarLIB Is Nothing Then
        Set mvarLIB = New clsLIB
    End If
    Set LIB = mvarLIB
End Property
Public Property Set LIB(vData As clsLIB)
    Set mvarLIB = vData
End Property
Public Property Set LIBS(ByVal vData As Object)
    Set mvarLIBS = vData
End Property
Public Property Get LIBS() As Collection
    Set LIBS = mvarLIBS
End Property
Public Property Set METODOS(ByVal vData As Object)
    Set mvarMETODOS = vData
End Property
Public Property Get METODOS() As Collection
    Set METODOS = mvarMETODOS
End Property
Public Property Set EVENTOS(ByVal vData As Object)
    Set mvarEVENTOS = vData
End Property
Public Property Get EVENTOS() As Collection
    Set EVENTOS = mvarEVENTOS
End Property
Public Property Get METODO() As clsMETODO
    Set METODO = mvarMETODO
    mvarMETODO.FileName = mvarFileName
End Property
Public Property Set METODO(vData As clsMETODO)
    Set mvarMETODO = vData
End Property
Public Property Get EVENTO() As clsEVENTO
    Set EVENTO = mvarEVENTO
    mvarEVENTO.FileName = mvarFileName
End Property
Public Property Set EVENTO(vData As clsEVENTO)
    Set mvarEVENTO = vData
End Property
Public Property Set CONTROLES(ByVal vData As Object)
    Set mvarCONTROLES = vData
End Property
Public Property Get CONTROLES() As Collection
    Set CONTROLES = mvarCONTROLES
End Property
Public Property Let Flag(ByVal vData As Boolean)
    mvarFlag = vData
End Property
Public Property Get Flag() As Boolean
    Flag = mvarFlag
End Property
Public Property Let MDI(ByVal vData As Boolean)
    mvarMDI = vData
End Property
Public Property Get MDI() As Boolean
    MDI = mvarMDI
End Property
Public Property Get CONTROLE() As clsCONTROLE
    Set CONTROLE = mvarCONTROLE
    'mvarCONTROLE.FileName = mvarFileName
End Property
Public Property Set CONTROLE(vData As clsCONTROLE)
    Set mvarCONTROLE = vData
End Property
Public Property Set FUNCOES(ByVal vData As Collection)
    Set mvarFUNCOES = vData
End Property
Public Property Get FUNCOES() As Collection
   Set FUNCOES = mvarFUNCOES
End Property
Public Property Set FUNCAO(vData As clsFUNCAO)
    Set mvarFUNCAO = vData
End Property
Public Property Get FUNCAO() As Collection
   Set FUNCAO = mvarFUNCAO
   mvarFUNCAO.FileName = mvarFileName
End Property
Public Property Set VARIAVEIS(ByVal vData As Collection)
   Set mvarVARIAVEIS = vData
End Property
Public Property Get VARIAVEIS() As Collection
   Set VARIAVEIS = mvarVARIAVEIS
End Property
Public Property Set PROPRIEDADES(ByVal vData As Collection)
   Set mvarPROPRIEDADES = vData
End Property
Public Property Get PROPRIEDADES() As Collection
   Set PROPRIEDADES = mvarPROPRIEDADES
End Property
Public Property Let PATH(ByVal vData As String)
   mvarPATH = vData
End Property
Public Property Get PATH() As String
   PATH = mvarPATH
End Property
Public Property Let TAMANHO(ByVal vData As String)
   mvarTAMANHO = vData
End Property
Public Property Get TAMANHO() As String
   TAMANHO = mvarTAMANHO
End Property
Public Property Let FileName(ByVal vData As String)
   mvarFileName = Trim(UCase(vData))
End Property
Public Property Get FileName() As String
   FileName = mvarFileName
End Property
Public Property Let NOME(ByVal vData As String)
   mvarNOME = vData
End Property
Public Property Get NOME() As String
   NOME = mvarNOME
End Property
Public Property Let LINHAS(ByVal vData As Integer)
   mvarLINHAS = vData
End Property
Public Property Get LINHAS() As Integer
   LINHAS = mvarLINHAS
End Property
Private Sub Class_Initialize()
   mvarMDI = False
   Set mvarPROPRIEDADE = New clsPROPRIEDADE
   Set mvarVARIAVEL = New clsVARIAVEL
   Set mvarCONTROLE = New clsCONTROLE
   Set mvarFUNCAO = New clsFUNCAO
   Set mvarEVENTO = New clsEVENTO
   Set mvarMETODO = New clsMETODO
   Set mvarLIB = New clsLIB
End Sub
Private Sub Class_Terminate()
   Set mvarPROPRIEDADE = Nothing
   Set mvarVARIAVEL = Nothing
   Set mvarMETODO = Nothing
   Set mvarFUNCAO = Nothing
   Set mvarEVENTO = Nothing
   Set mvarCONTROLE = Nothing
   
   Set mvarPROPRIEDADES = Nothing
   Set mvarVARIAVEIS = Nothing
   Set mvarMETODOS = Nothing
   Set mvarFUNCAO = Nothing
   Set mvarEVENTOS = Nothing
   Set mvarCONTROLES = Nothing
   Set mvarLIBS = Nothing
End Sub
