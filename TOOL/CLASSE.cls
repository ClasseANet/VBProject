VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLASSE"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Private mvarLINHAS As String
Private mvarFILENAME As String
Private mvarNOME As String
Private mvarTAMANHO As String
Private mvarPATH As String
Private mvarENUMEROS As New Collection
Private mvarPROPRIEDADES As New Collection
Private mvarMETODOS As New Collection
Private mvarEVENTOS As New Collection
Public Property Set ENUMEROS(ByVal vData As Object)
    Set mvarENUM = vData
End Property
Public Property Get ENUMEROS() As Collection
    Set ENUMEROS = mvarENUMEROS
End Property
Public Property Set EVENTOS(ByVal vData As Object)
    Set mvarEVENTOS = vData
End Property
Public Property Get EVENTOS() As Collection
    Set EVENTOS = mvarEVENTOS
End Property
Public Property Set METODOS(ByVal vData As Object)
    Set mvarMETODOS = vData
End Property
Public Property Get METODOS() As Collection
    Set METODOS = mvarMETODOS
End Property
Public Property Set PROPRIEDADES(ByVal vData As Object)
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
Public Property Let NOME(ByVal vData As String)
    mvarNOME = vData
End Property
Public Property Get NOME() As String
    NOME = mvarNOME
End Property
Public Property Let FILENAME(ByVal vData As String)
   mvarFILENAME = Trim(UCase(vData))
End Property
Public Property Get FILENAME() As String
    FILENAME = mvarFILENAME
End Property
Public Property Let LINHAS(ByVal vData As String)
    mvarLINHAS = vData
End Property
Public Property Get LINHAS() As String
    LINHAS = mvarLINHAS
End Property



