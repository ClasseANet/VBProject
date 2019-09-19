VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Set_Edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Private mvarAppName As String
Private mvarDesenvolvedor As String
Private mvarWebSite As String
Private mvareMail As String
Private mvarTelefone As String

Private mvarErrorLabel As String
Private mvarErrorFunction As String
Private mvarSaidaLabel As String
Private mvarSaidaFunction As String
Private mvarSpcIndent As Integer
Private mvarTemplates As New Collection
Private mvarTemplate As String
Private mvarUserComment As String
Private mvarCharComment As String
Private mvarIndentFunction As Boolean
Private mvarIndentComment As Boolean
Private mvarIndentSelect As Boolean
Private mvarIdioma As Integer
Private mvarLineBlankBefore As Boolean
Private mvarExeAuxiliar As String
Public Property Let ExeAuxiliar(ByVal vData As String)
   mvarExeAuxiliar = vData
End Property
Public Property Get ExeAuxiliar() As String
   ExeAuxiliar = mvarExeAuxiliar
End Property
Public Property Let LineBlankBefore(ByVal vData As Boolean)
   mvarLineBlankBefore = vData
End Property
Public Property Get LineBlankBefore() As Boolean
   LineBlankBefore = mvarLineBlankBefore
End Property
Public Property Let Idioma(ByVal vData As Integer)
   mvarIdioma = vData
End Property
Public Property Get Idioma() As Integer
   Idioma = mvarIdioma
End Property
Public Property Let IndentSelect(ByVal vData As Boolean)
   mvarIndentSelect = vData
End Property
Public Property Get IndentSelect() As Boolean
   IndentSelect = mvarIndentSelect
End Property
Public Property Let IndentComment(ByVal vData As Boolean)
   mvarIndentComment = vData
End Property
Public Property Get IndentComment() As Boolean
   IndentComment = mvarIndentComment
End Property
Public Property Let IndentFunction(ByVal vData As Boolean)
   mvarIndentFunction = vData
End Property
Public Property Get IndentFunction() As Boolean
   IndentFunction = mvarIndentFunction
End Property
Public Property Let CharComment(ByVal vData As String)
   mvarCharComment = vData
End Property
Public Property Get CharComment() As String
   CharComment = mvarCharComment
End Property
Public Property Let UserComment(ByVal vData As String)
   mvarUserComment = vData
End Property
Public Property Get UserComment() As String
   UserComment = mvarUserComment
End Property
Public Property Let Template(ByVal vData As String)
   mvarTemplate = vData
End Property
Public Property Get Template() As String
   Template = mvarTemplate
End Property
Public Property Get Templates() As Collection
   If mvarTemplates Is Nothing Then
      Set mvarTemplates = New Collection
   End If
   Set Templates = mvarTemplates
End Property
Public Property Set Templates(ByVal vData As Collection)
   Set mvarTemplates = vData
End Property
Public Property Let SpcIndent(ByVal vData As Integer)
   mvarSpcIndent = vData
End Property
Public Property Get SpcIndent() As Integer
   SpcIndent = mvarSpcIndent
End Property
Public Property Let SaidaLabel(ByVal vData As String)
   mvarSaidaLabel = vData
End Property
Public Property Get SaidaLabel() As String
   SaidaLabel = mvarSaidaLabel
End Property
Public Property Let SaidaFunction(ByVal vData As String)
   mvarSaidaFunction = vData
End Property
Public Property Get SaidaFunction() As String
   SaidaFunction = mvarSaidaFunction
End Property
Public Property Let ErrorLabel(ByVal vData As String)
   mvarErrorLabel = vData
End Property
Public Property Get ErrorLabel() As String
   ErrorLabel = mvarErrorLabel
End Property
Public Property Let ErrorFunction(ByVal vData As String)
   mvarErrorFunction = vData
End Property
Public Property Get ErrorFunction() As String
   ErrorFunction = mvarErrorFunction
End Property
Public Property Let Desenvolvedor(ByVal vData As String)
   mvarDesenvolvedor = vData
End Property
Public Property Get Desenvolvedor() As String
   Desenvolvedor = mvarDesenvolvedor
End Property
Public Property Let WebSite(ByVal vData As String)
   mvarWebSite = vData
End Property
Public Property Get WebSite() As String
   WebSite = mvarWebSite
End Property
Public Property Let eMail(ByVal vData As String)
   mvareMail = vData
End Property
Public Property Get eMail() As String
   eMail = mvareMail
End Property
Public Property Let Telefone(ByVal vData As String)
   mvarTelefone = vData
End Property
Public Property Get Telefone() As String
   Telefone = mvarTelefone
End Property
Public Property Get AppName() As String
   AppName = "VbEditorUtil Editor"
End Property

Private Sub Class_Initialize()
   mvarIndentFunction = True
   mvarIndentComment = True
   mvarIndentSelect = True
   mvarIdioma = 5000 '* Português
   mvarLineBlankBefore = False
End Sub

Private Sub Class_Terminate()
   Set mvarTemplates = Nothing
End Sub
