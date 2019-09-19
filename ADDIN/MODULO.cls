VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMODULO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"clsMEMBRO"
Private mvarLINHAS As Integer
Private mvarNOME As String
Private mvarFILENAME As String
Private mvarPATH As String
Private mvarMEMBER As clsMEMBRO
Private mvarPagProp As PROPPAGE
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
Public Property Let LINHAS(ByVal vData As Integer)
    mvarLINHAS = vData
End Property
Public Property Get LINHAS() As Integer
    LINHAS = mvarLINHAS
End Property
Private Sub Class_Terminate()
  Set mvarMEMBER = Nothing
End Sub

