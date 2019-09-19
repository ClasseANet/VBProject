VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Titulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Enum eAlin
   Esquerda
   Direita
   Centro
End Enum

Private mvarTexto       As String
Private mvarTamanho     As Long
Private mvarAlinhamento As eAlin
Private mvarCampo       As String

Private mvarItem        As Collection
Public Property Let Item(ByVal vData As Collection)
    Set mvarItem = vData
End Property
Public Property Get Item() As Collection
    Set Item = mvarItem
End Property
Public Property Let Texto(ByVal vData As String)
    mvarTexto = vData
End Property
Public Property Get Texto() As String
    Texto = mvarTexto
End Property
Public Property Let Tamanho(ByVal vData As Long)
    mvarTamanho = vData
End Property
Public Property Get Tamanho() As Long
    Tamanho = mvarTamanho
End Property
Public Property Let Alinhamento(ByVal vData As eAlin)
    mvarAlinhamento = vData
End Property
Public Property Get Alinhamento() As eAlin
    Alinhamento = mvarAlinhamento
End Property
Public Property Let Campo(ByVal vData As String)
    mvarCampo = vData
End Property
Public Property Get Campo() As String
    Campo = mvarCampo
End Property
Public Sub Add(Optional Campo As String, Optional Texto As String, Optional Tamanho As Long, Optional Alinhamento As eAlin)
   Dim MyTitulo As Titulo
   Dim sCod     As String
   
   Set MyTitulo = New Titulo
   With MyTitulo
      .Alinhamento = Alinhamento
      .Campo = Campo
      .Tamanho = Tamanho
      .Texto = Texto
   End With
   If mvarItem.Count = 1 Then
      If Trim(mvarItem(1).Texto & mvarItem(1).Campo) = "" Then
         'mvarItem.Remove 1
         Set mvarItem = New Collection
      End If
   End If
   If Not ClsDsr.ExisteItem(mvarItem, MyTitulo.Campo) Then
      mvarItem.Add MyTitulo, MyTitulo.Campo
   End If
   sCod = MyTitulo.Campo
   Set MyTitulo = Nothing
   
   On Error Resume Next
   mvarItem(sCod).Alinhamento = Alinhamento
   mvarItem(sCod).Campo = Campo
   mvarItem(sCod).Tamanho = Tamanho
   mvarItem(sCod).Texto = Texto
End Sub
Public Function Count() As Long
    Count = mvarItem.Count
End Function
Private Sub Class_Initialize()
   Set mvarItem = New Collection
End Sub
Private Sub Class_Terminate()
   Set mvarItem = Nothing
End Sub
