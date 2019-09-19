VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Detalhes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarNome     As String
Private mvarTipo     As String
Private mvarTamanho  As String
Private mvarCampo    As String
Private mvarTabela   As String
Private mvarCombo    As String
Private mvarTPCmb    As String
Private mvarItem01   As Double
Private mvarColecao  As Collection
Public Property Set colecao(vdados As Collection)
    Set mvarColecao = vdados
End Property
Public Property Get colecao() As Collection
    If mvarColecao Is Nothing Then
        Set mvarColecao = New Collection
    End If
    Set colecao = mvarColecao
End Property

Public Property Let Nome(vdados As String)
    mvarNome = vdados
End Property
Public Property Get Nome() As String
    Nome = mvarNome
End Property
Public Property Let Tipo(vdados As String)
    mvarTipo = vdados
End Property
Public Property Get Tipo() As String
    Tipo = mvarTipo
End Property

Public Property Let Tamanho(vdados As String)
    mvarTamanho = vdados
End Property
Public Property Get Tamanho() As String
    Tamanho = mvarTamanho
End Property
Public Property Let Campo(vdados As String)
    mvarCampo = vdados
End Property
Public Property Get Campo() As String
    Campo = mvarCampo
End Property
Public Property Let Tabela(vdados As String)
    mvarTabela = vdados
End Property
Public Property Get Tabela() As String
    Tabela = mvarTabela
End Property
Public Property Let Combo(vdados As String)
    mvarCombo = vdados
End Property
Public Property Get Combo() As String
    Combo = mvarCombo
End Property
Public Property Let TpCmb(vdados As String)
    mvarTPCmb = vdados
End Property
Public Property Get TpCmb() As String
    TpCmb = mvarTPCmb
End Property
Public Property Let Item01(vdados As String)
    mvarItem01 = vdados
End Property
Public Property Get Item01() As String
    Item01 = mvarItem01
End Property

Private Sub Class_Initialize()
   mvarNome = ""
   mvarTipo = ""
   mvarTamanho = ""
   mvarItem01 = 0
   mvarCampo = ""
   mvarTabela = ""
   mvarCombo = ""
   mvarTPCmb = ""
End Sub
