VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cartao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarNumeros As Collection
Private mvarSomas As Collection
Private mvarValido As Boolean
Private mvarVerificado As Boolean
Public Property Let Valido(ByVal vData As Boolean)
    mvarValido = vData
End Property
Public Property Get Valido() As Boolean
    Valido = mvarValido
End Property
Public Property Let Verificado(ByVal vData As Boolean)
    mvarVerificado = vData
End Property
Public Property Get Verificado() As Boolean
    Verificado = mvarVerificado
End Property
Public Property Set Somas(ByVal vData As Collection)
    Set mvarSomas = vData
End Property
Public Property Get Somas() As Collection
    If mvarSomas Is Nothing Then
       Set mvarSomas = New Collection
    End If
    Set Somas = mvarSomas
End Property
Public Property Set Numeros(ByVal vData As Collection)
    Set mvarNumeros = vData
End Property
Public Property Get Numeros() As Collection
    If mvarNumeros Is Nothing Then
       Set mvarNumeros = New Collection
    End If
    Set Numeros = mvarNumeros
End Property
Public Sub MontarSomas(GrpSoma As Integer)
   Dim i As Integer
   Dim n As Variant
   Dim nAux As Double
   Dim Pos As Collection
   Dim PosA()
   
   If mvarSomas Is Nothing Then Set mvarSomas = New Collection
   If mvarNumeros Is Nothing Then Set mvarNumeros = New Collection
   
   Set Pos = New Collection
   ReDim PosA(GrpSoma)
   For i = 1 To GrpSoma
      Pos.Add i
      PosA(i) = 0
   Next
   k = 0
   While PosA(1) <= mvarNumeros.Count - GrpSoma + 1
      nAux = 0
      For i = 1 To GrpSoma
         nAux = nAux + mvarNumeros(Pos(i) + PosA(i))
      Next
      mvarSomas.Add nAux
      Call SomaPos(PosA, GrpSoma, mvarNumeros.Count - GrpSoma)
   Wend
End Sub
Public Sub SomaPos(ByRef PosA, Item, MaxA)
   If Item = 0 Then
      PosA(1) = MaxA + 2
   Else
      If PosA(Item) < MaxA Then
         PosA(Item) = PosA(Item) + 1
      Else
         Call SomaPos(PosA, (Item - 1), MaxA)
         For i = Item To UBound(PosA)
            PosA(i) = PosA(i - 1)
         Next
      End If
   End If
   PosA(0) = PosA(1)
End Sub
Public Function SomaSeg(Pos1, Pos2, QtdPos) As Double
   Dim i As Integer
   Dim nAux As Double
   Dim PosFim As Integer
   
   If QtdPos = 1 Then Exit Function
   
   PosFim = Pos2 + QtdPos - 1
   'On Error Resume Next
   
   nAux = mvarNumeros(Pos1)
   For i = Pos2 To PosFim
      nAux = nAux + SomaSeg(i, i + 1, QtdPos - 1)
   Next
   SomaSeg = nAux
End Function
Private Sub Class_Initialize()
   mvarValido = True
End Sub

Private Sub Class_Terminate()
   Set mvarSomas = Nothing
   Set mvarNumeros = Nothing
End Sub
