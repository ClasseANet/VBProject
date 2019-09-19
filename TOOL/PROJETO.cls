VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PROJETO"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"CLASSE"
Attribute VB_Ext_KEY = "Member1" ,"FORMULARIO"
Attribute VB_Ext_KEY = "Member2" ,"FUNCAO"
Attribute VB_Ext_KEY = "Member3" ,"MODULO"
Attribute VB_Ext_KEY = "Member4" ,"Modulo"
Private mvarCLASSES As New Collection
Private mvarFORMS As New Collection
Private mvarMODULOS As New Collection
Private mvarNOME As String
Private mvarFILENAME As String
Private mvarRESFILE As String
Private mvarDESCRI플O As String
Private mvarPATH As String '
Private mvarFORMULARIO As FORMULARIO
Private mvarCLASSE As CLASSE
Private mvarMODULO As Modulo
Private mvarFUNCAO As FUNCAO
Private mvarFUNCOES As New Collection  'local copy
Private mvarLINHAS As String 'local copy
Public Property Let LINHAS(ByVal vData As String)
    mvarLINHAS = vData
End Property
Public Property Get LINHAS() As String
    LINHAS = mvarLINHAS
End Property
Public Property Set FUNCOES(ByVal vData As Collection)
    Set mvarFUNCOES = vData
End Property
Public Property Get FUNCOES() As Collection
    Set FUNCOES = mvarFUNCOES
End Property
Public Property Get FUNCAO() As FUNCAO
    If mvarFUNCAO Is Nothing Then
        Set mvarFUNCAO = New FUNCAO
    End If
    Set FUNCAO = mvarFUNCAO
End Property
Public Property Set FUNCAO(vData As FUNCAO)
    Set mvarFUNCAO = vData
End Property
Public Property Get Modulo() As Modulo
    If mvarMODULO Is Nothing Then
        Set mvarMODULO = New Modulo
    End If
    Set Modulo = mvarMODULO
End Property
Public Property Set Modulo(vData As Modulo)
    Set mvarMODULO = vData
End Property
Public Property Get CLASSE() As CLASSE
    If mvarCLASSE Is Nothing Then
        Set mvarCLASSE = New CLASSE
    End If
    Set CLASSE = mvarCLASSE
End Property
Public Property Set CLASSE(vData As CLASSE)
    Set mvarCLASSE = vData
End Property
Public Property Get FORMULARIO() As FORMULARIO
    If mvarFORMULARIO Is Nothing Then
        Set mvarFORMULARIO = New FORMULARIO
    End If
    Set FORMULARIO = mvarFORMULARIO
End Property
Public Property Set FORMULARIO(vData As FORMULARIO)
    Set mvarFORMULARIO = vData
End Property
Private Sub Class_Terminate()
  Set mvarFUNCAO = Nothing
  Set mvarMODULO = Nothing
  Set mvarFUNCAO = Nothing
  Set mvarCLASSE = Nothing
  Set mvarFORMULARIO = Nothing
End Sub
Public Property Let PATH(ByVal vData As String)
    mvarPATH = vData
End Property
Public Property Get PATH() As String
    PATH = mvarPATH
End Property
Public Property Let DESCRI플O(ByVal vData As String)
    mvarDESCRI플O = vData
End Property
Public Property Get DESCRI플O() As String
    DESCRI플O = mvarDESCRI플O
End Property
Public Property Let RESFILE(ByVal vData As String)
    mvarRESFILE = vData
End Property
Public Property Get RESFILE() As String
    RESFILE = mvarRESFILE
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
Public Property Set MODULOS(ByVal vData As Collection)
    Set mvarMODULOS = vData
End Property
Public Property Get MODULOS() As Collection
    Set MODULOS = mvarMODULOS
End Property
Public Property Set FORMS(ByVal vData As Collection)
    Set mvarFORMS = vData
End Property
Public Property Get FORMS() As Collection
    Set FORMS = mvarFORMS
End Property
Public Property Set CLASSES(ByVal vData As Collection)
    Set mvarCLASSES = vData
End Property
Public Property Get CLASSES() As Collection
    Set CLASSES = mvarCLASSES
End Property
