VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLIB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarLINHAS As Integer
Private mvarNOME As String
Private mvarFileName As String
Private mvarPATH As String
Private mvarMEMBER As clsMEMBRO
Private mvarFUNCAO As clsFUNCAO
Private mvarFUNCOES As New Collection
Private mvarPagProp As PROPPAGE
Private mvarStartLine As Integer
Public Property Let StartLine(ByVal vData As Integer)
    mvarStartLine = vData
End Property
Public Property Get StartLine() As Integer
    StartLine = mvarStartLine
End Property
Public Property Get FUNCAO(NOME$) As clsFUNCAO
    If Trim(NOME$) <> "" Then
       Set mvarFUNCAO = mvarFUNCOES(NOME$)
       mvarFUNCAO.FileName = mvarFileName
    End If
    Set FUNCAO = mvarFUNCAO
End Property
Public Property Set FUNCOES(ByVal vData As Collection)
    Set mvarFUNCOES = vData
End Property
Public Property Get FUNCOES() As Collection
    Set FUNCOES = mvarFUNCOES
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
Public Property Get MEMBER() As clsMEMBRO
    If mvarMEMBER Is Nothing Then
        Set mvarMEMBER = New clsMEMBRO
        mvarMEMBER.FileName = mvarFileName
    End If
    Set MEMBER = mvarMEMBER
End Property
Public Property Set MEMBER(vData As clsMEMBRO)
    Set mvarMEMBER = vData
End Property
Public Property Get PATH() As String
    PATH = mvarPATH
End Property
Public Property Let PATH(ByVal vData As String)
    mvarPATH = vData
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
   Set mvarFUNCAO = New clsFUNCAO
End Sub

Private Sub Class_Terminate()
  Set mvarMEMBER = Nothing
  Set mvarFUNCAO = Nothing
End Sub


