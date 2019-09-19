VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFORMULARIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"clsMEMBRO"
Private mvarLINHAS As Integer
Private mvarNOME As String
Private mvarFileName As String
Private mvarTAMANHO As String
Private mvarPATH As String

Private mvarFlag As Boolean
Private mvarMDI As Boolean
Private mvarMEMBER As clsMEMBRO
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
End Sub
Private Sub Class_Terminate()
  Set mvarMEMBER = Nothing
End Sub
