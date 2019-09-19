VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPROJETO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"clsCLASSE"
Attribute VB_Ext_KEY = "Member1" ,"clsFORMULARIO"
Attribute VB_Ext_KEY = "Member2" ,"clsMODULO"
Attribute VB_Ext_KEY = "Member3" ,"clsRELDOC"
Private mvarFORMS As New Collection
Private mvarCLASSES As New Collection
Private mvarMODULOS As New Collection
Private mvarFUNCOES As New Collection
Private mvarRELDOCS As New Collection
Private mvarLIBS As New Collection

Private mvarFORMULARIO As clsFORMULARIO
Private mvarCLASSE As clsCLASSE
Private mvarMODULO As clsMODULO
Private mvarFUNCAO As clsFUNCAO
Private mvarRELDOC As clsRELDOC
Private mvarLIB As clsLIB

Private mvarNOME As String
Private mvarFileName As String
Private mvarRESFILE As String
Private mvarDESCRI플O As String
Private mvarPATH As String '

Private mvarLINHAS As Integer
Public Property Let LINHAS(ByVal vData As Integer)
    mvarLINHAS = vData
End Property
Public Property Get LINHAS() As Integer
    LINHAS = mvarLINHAS
End Property
Public Property Set FUNCOES(ByVal vData As Collection)
    Set mvarFUNCOES = vData
End Property
Public Property Get FUNCOES() As Collection
    Set FUNCOES = mvarFUNCOES
End Property
Public Property Set RELDOCS(ByVal vData As Collection)
    Set mvarRELDOCS = vData
End Property
Public Property Get RELDOCS() As Collection
    Set RELDOCS = mvarRELDOCS
End Property
Public Property Get FUNCAO(NOME$) As clsFUNCAO
    If Trim(NOME$) <> "" Then
       If vbKey0 <= Asc(Mid(NOME$, 1, 1)) And Asc(Mid(NOME$, 1, 1)) <= vbKey9 Then
          Set mvarFUNCAO = mvarFUNCOES(CInt(NOME$))
       Else
          Set mvarFUNCAO = mvarFUNCOES(NOME$)
       End If
    End If
    Set FUNCAO = mvarFUNCAO
End Property
Public Property Get MODULO(NOME$) As clsMODULO
   On Error Resume Next
    If mvarMODULO Is Nothing Then
        Set mvarMODULO = New clsMODULO
    End If
    If Trim(NOME$) <> "" Then
      If vbKey0 <= Asc(Mid(NOME$, 1, 1)) And Asc(Mid(NOME$, 1, 1)) <= vbKey9 Then
         Set mvarMODULO = mvarMODULOS(CInt(NOME$))
      Else
         Set mvarMODULO = mvarMODULOS(NOME$)
      End If
    End If
    Set MODULO = mvarMODULO
End Property
Public Property Get CLASSE(NOME$) As clsCLASSE
    If mvarCLASSE Is Nothing Then
        Set mvarCLASSE = New clsCLASSE
    End If
    If Trim(NOME$) <> "" Then
      If vbKey0 <= Asc(Mid(NOME$, 1, 1)) And Asc(Mid(NOME$, 1, 1)) <= vbKey9 Then
         Set mvarCLASSE = mvarCLASSES(CInt(NOME$))
      Else
         Set mvarCLASSE = mvarCLASSES(NOME$)
      End If
    End If
    Set CLASSE = mvarCLASSE
End Property
'Public Property Set CLASSE(vData As CLASSE)
'    Set mvarCLASSE = vData
'End Property
Public Property Get FORMULARIO(NOME$) As clsFORMULARIO
   On Error Resume Next
   If mvarFORMULARIO Is Nothing Then
      Set mvarFORMULARIO = New clsFORMULARIO
   End If
   If Trim(NOME$) <> "" Then
      If vbKey0 <= Asc(Mid(NOME$, 1, 1)) And Asc(Mid(NOME$, 1, 1)) <= vbKey9 Then
         Set mvarFORMULARIO = mvarFORMS(CInt(NOME$))
      Else
         Set mvarFORMULARIO = mvarFORMS(NOME$)
      End If
   End If
   Set FORMULARIO = mvarFORMULARIO
End Property
Public Property Get RELDOC(NOME$) As clsRELDOC
    If Trim(NOME$) <> "" Then
       If vbKey0 <= Asc(Mid(NOME$, 1, 1)) And Asc(Mid(NOME$, 1, 1)) <= vbKey9 Then
          Set mvarFRELDOC = mvarRELDOCS(CInt(NOME$))
       Else
          Set mvarFRELDOC = mvarRELDOCS(NOME$)
       End If
    End If
    Set RELDOC = mvarRELDOC
End Property
Public Property Get LIB(NOME$) As clsLIB
    If Trim(NOME$) <> "" Then
       If vbKey0 <= Asc(Mid(NOME$, 1, 1)) And Asc(Mid(NOME$, 1, 1)) <= vbKey9 Then
          Set mvarLIB = mvarLIBS(CInt(NOME$))
       Else
          Set mvarLIB = mvarLIBS(NOME$)
       End If
    End If
    Set LIB = mvarLIB
End Property
Private Sub Class_Initialize()
   Set mvarFORMULARIO = New clsFORMULARIO
   Set mvarCLASSE = New clsCLASSE
   Set mvarMODULO = New clsMODULO
   Set mvarFUNCAO = New clsFUNCAO
   Set mvarRELDOC = New clsRELDOC
   Set mvarLIB = New clsLIB
End Sub
Private Sub Class_Terminate()
   Set mvarRELDOC = Nothing
   Set mvarFUNCAO = Nothing
   Set mvarMODULO = Nothing
   Set mvarFUNCAO = Nothing
   Set mvarCLASSE = Nothing
   Set mvarFORMULARIO = Nothing
   Set mvarLIB = Nothing
  
   Set mvarFORMS = Nothing
   Set mvarCLASSES = Nothing
   Set mvarMODULOS = Nothing
   Set mvarFUNCOES = Nothing
   Set mvarRELDOCS = Nothing
   Set mvaLIBS = Nothing
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
Public Property Set LIBS(ByVal vData As Collection)
    Set mvarLIBS = vData
End Property
Public Property Get LIBS() As Collection
    Set LIBS = mvarLIBS
End Property

