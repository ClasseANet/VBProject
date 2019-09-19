VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MODULO"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Private mvarLINHAS As String
Private mvarNOME As String
Private mvarFILENAME As String
Private mvarFUNCOES As New Collection 'local copy
Private mvarPATH As String 'local copy
Private mvarVARIAVEIS As New Collection 'local copy
Private mvarENUMEROS As New Collection 'local copy
Public Property Set ENUMEROS(ByVal vData As Collection)
    Set mvarENUMEROS = vData
End Property
Public Property Get ENUMEROS() As Collection
    Set ENUMEROS = mvarENUMEROS
End Property
Public Property Set VARIAVEIS(ByVal vData As Collection)
    Set mvarVARIAVEIS = vData
End Property
Public Property Get VARIAVEIS() As Collection
    Set VARIAVEIS = mvarVARIAVEIS
End Property
Public Property Let PATH(ByVal vData As String)
    mvarPATH = vData
End Property
Public Property Get PATH() As String
    PATH = mvarPATH
End Property
Public Property Set FUNCOES(ByVal vData As Object)
    Set mvarFUNCOES = vData
End Property
Public Property Get FUNCOES() As Collection
    Set FUNCOES = mvarFUNCOES
End Property
Public Property Let FILENAME(ByVal vData As String)
    mvarFILENAME = Trim(UCase(vData))
End Property
Public Property Get FILENAME() As String
    FILENAME = mvarFILENAME
End Property
Public Property Let NOME(ByVal vData As String)
    mvarNOME = vData
End Property
Public Property Get NOME() As String
    NOME = mvarNOME
End Property
Public Property Let LINHAS(ByVal vData As String)
    mvarLINHAS = vData
End Property
Public Property Get LINHAS() As String
    LINHAS = mvarLINHAS
End Property

