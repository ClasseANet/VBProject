VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FORMULARIO"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"CONTROLE"
Private mvarLINHAS As String
Private mvarNOME As String
Private mvarFILENAME As String
Private mvarTAMANHO As String
Private mvarPATH As String
Private mvarVARIAVEIS As New Collection
Private mvarFUNCOES As New Collection
Private mvarCONTROLE As CONTROLE
Private mvarFlag As Boolean
Private mvarCONTROLES As New Collection
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
Public Property Get CONTROLE() As CONTROLE
    Set CONTROLE = mvarCONTROLE
End Property
Public Property Set CONTROLE(vData As CONTROLE)
    Set mvarCONTROLE = vData
End Property
Public Property Set FUNCOES(ByVal vData As Collection)
    Set mvarFUNCOES = vData
End Property
Public Property Get FUNCOES() As Collection
   Set FUNCOES = mvarFUNCOES
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
Public Property Let TAMANHO(ByVal vData As String)
   mvarTAMANHO = vData
End Property
Public Property Get TAMANHO() As String
   TAMANHO = mvarTAMANHO
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
Private Sub Class_Initialize()
   Set mvarCONTROLE = New CONTROLE
   Set mvarCONTROLES = New Collection
End Sub
Private Sub Class_Terminate()
   Set mvarCONTROLE = Nothing
   Set mvarCONTROLES = Nothing
End Sub

