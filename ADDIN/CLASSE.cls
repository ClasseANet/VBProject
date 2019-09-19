VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCLASSE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"clsMEMBRO"
Private mvarLINHAS As Integer
Private mvarFileName As String
Private mvarNOME As String
Private mvarTAMANHO As String
Private mvarPATH As String
Private mvarPAI As String
Private mvarOCCUR As String
Private mvarMEMBER As clsMEMBRO
Private mvarPagProp As PROPPAGE

Private mvarSubClasses As New Collection
'Private mvarSubClasse As CLASSE
'Public Property Get SubClasse() As CLASSE
'    If mvarSubClasse Is Nothing Then
'        Set mvarSubClasse = New CLASSE
'    End If
'    Set SubClasse = mvarSubClasse
'End Property
'Public Property Set SubClasse(vData As CLASSE)
'    Set mvarSubClasse = vData
'End Property
Public Property Set SubClasses(ByVal vData As Collection)
   Set mvarSubClasses = vData
End Property
Public Property Get SubClasses() As Collection
   Set SubClasses = mvarSubClasses
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
Public Property Get PagProp() As PROPPAGE
   If mvarPagProp Is Nothing Then
      Set mvarPagProp = New PROPPAGE
   End If
   Set PagProp = mvarPagProp
End Property
Public Property Set PagProp(vData As PROPPAGE)
   Set mvarPagProp = vData
End Property
Private Sub Class_Terminate()
   Set mvarMEMBER = Nothing
   Set mvarPagProp = Nothing
End Sub
Public Property Let OCCUR(ByVal vData As String)
   mvarOCCUR = vData
End Property
Public Property Get OCCUR() As String
   OCCUR = mvarOCCUR
End Property
Public Property Let PAI(ByVal vData As String)
   mvarPAI = vData
End Property
Public Property Get PAI() As String
   PAI = mvarPAI
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
Public Property Let FileName(ByVal vData As String)
   mvarFileName = Trim(UCase(vData))
End Property
Public Property Get FileName() As String
   FileName = mvarFileName
End Property
Public Property Let LINHAS(ByVal vData As Integer)
   mvarLINHAS = vData
End Property
Public Property Get LINHAS() As Integer
   LINHAS = mvarLINHAS
End Property
