VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAD"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"

'Private WithEvents mvarMeForm As FrmLov
Private mvarRS As Object
Private mvarDBase As Object
Private mvarMeForm As Object
Private mvarIdField As Variant
Private mvarCaption$

Private mvarAcesso As String
Private mvarDBaseName As String

'Private mvarWs As Workspace
'Private mvarDBaseName As object

Private mvarTable$
Private mvarCab As Variant
Private mvarQuery As String 'Variant

Private mvarFrmCad As Object

Private mvarGrd As Object
Private mvarData As Object
Private mvarAplic As Object
Private mvarIdioma As Integer
Public Property Let Idioma(vData As Integer)
    mvarIdioma = vData
End Property
Public Property Get Idioma() As Integer
    Idioma = mvarIdioma
End Property
Public Sub Show()
Attribute Show.VB_Description = "Exibe Tela de Seleção para Cadastro."
'   Set mvarMeForm = FrmCad
   On Error GoTo TError
   
'   With mvarMeForm.ClsCad
'      .DBase = mvarDBaseName
'      .Table = mvarTable
'      .Aplic = mvarAplic
'      .Idioma = mvarIdioma
'      .Cab = mvarCab
'      .Caption = mvarCaption
'      .IdField = mvarIdField
'      .Query = mvarQuery$
'      .FrmCad = mvarFrmCad
'   End With
'   'mvarMeForm.MDIChild = True
   
   Call ClsCtrl.SetTag(mvarMeForm, "ACESSO", mvarAcesso)
   mvarMeForm.Show vbModeless
   Exit Sub
TError:
   Call ClsMsg.ShowError
End Sub
Public Property Get MeForm() As Object
    Set MeForm = mvarMeForm
End Property
Public Property Let MeForm(Frm As Object)
    Set mvarMeForm = Frm
End Property
Public Property Let Acesso(ByVal vData As String)
Attribute Acesso.VB_Description = "Define o nível de acesso ao módulo."
    mvarAcesso = vData
End Property
Public Property Get Acesso() As String
    Acesso = mvarAcesso
End Property
Public Property Get Table() As String
    Table = mvarTable
End Property
Public Property Let Table(Txt$)
    mvarTable = Txt$
End Property
Public Property Get Cab() As Variant
Attribute Cab.VB_Description = "Vetor contendo o Caveeçalho do Grid de exibição. Arr(Campo,Descrição, Tamanho, Alinhamento)"
    Cab = mvarCab
End Property
Public Property Let Cab(Vet As Variant)
    mvarCab = Vet
End Property
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Define o título da Tela."
    Caption = mvarCaption
End Property
Public Property Let Caption(Txt$)
    mvarCaption = Txt$
End Property
Public Property Get Query() As Variant
    Query = mvarQuery
End Property
Public Property Let Query(Txt As Variant)
    mvarQuery = Txt
End Property
Public Property Get IdField() As Variant
    IdField = mvarIdField
End Property
Public Property Let IdField(Vet As Variant)
    mvarIdField = Vet
End Property
Public Property Get FrmCad() As Object
    Set FrmCad = mvarFrmCad
End Property
Public Property Let FrmCad(Frm As Object)
    Set mvarFrmCad = Frm
End Property
Public Property Get Grd() As Object
    Set Grd = mvarGrd
End Property
Public Property Let Grd(Grd As Object)
    Set mvarGrd = Grd
End Property
Public Property Get Data() As Object
Attribute Data.VB_Description = "Objeto ""DataSource"" gerenciador dos dados a serem exibidos."
    Set Data = mvarData
End Property
Public Property Let Data(Dat As Object)
    Set mvarData = Dat
End Property
Public Property Get Aplic() As Object
Attribute Aplic.VB_Description = "Objeto ""App"" relacionada a sua Aplicação."
    Set Aplic = mvarAplic
End Property
Public Property Let Aplic(vData As Object)
    Set mvarAplic = vData
End Property
Public Property Get Dbase() As Object
    Set Dbase = mvarDBase
End Property
Public Property Let Dbase(vData As Object)
    Set mvarDBase = vData
End Property
Public Property Get RS() As Object
    Set RS = mvarRS
End Property
Public Property Let RS(vData As Object)
    Set mvarRS = vData
End Property
Public Property Set RS(vData As Object)
    Set mvarRS = vData
End Property
Public Property Set Dbase(vData As Object)
    Set mvarDBase = vData
End Property
Public Property Get DBaseName() As String
    DBaseName = mvarDBaseName
End Property
Public Property Let DBaseName(vData As String)
    mvarDBaseName = vData
End Property
Private Sub Class_Initialize()
   mvarTable = ""
   mvarCaption = "Cadastro"
   mvarQuery = ""
   mvarIdField = Array("")
   mvarCab = Array("")
   Set mvarFrmCad = Nothing
   mvarAcesso = ""
   mvarDBaseName = ""
   mvarIdioma = 5000 'Português
End Sub
Private Sub Class_Terminate()
   On Error Resume Next
   Set mvarFrmCad = Nothing
   Set mvarMeForm = Nothing
   Set mvarRS = Nothing
   Set mvarDBase = Nothing
End Sub

