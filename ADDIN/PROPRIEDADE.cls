VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPROPRIEDADE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"

Private mvarNOME As String
Private mvarFileName As String
Private mvarPagProp As PROPPAGE
Private mvarCom_Get  As Boolean
Private mvarCom_Let  As Boolean
Private mvarCom_Set  As Boolean
Private mvarValorPadrao As String
Private mvarEscopo As String
Private mvarTIPO As String
Public Property Let ValorPadrao(ByVal vData As String)
    mvarValorPadrao = vData
End Property
Public Property Get ValorPadrao() As String
    ValorPadrao = mvarValorPadrao
End Property
Public Property Let Escopo(ByVal vData As String)
'     Case 1: ItemX.SubItems(1) = "'Private'"
'     Case 2: ItemX.SubItems(1) = "'Public'"
'     Case 3: ItemX.SubItems(1) = "'Friend'"
'     Case Else: ItemX.SubItems(1) = "'Outro'"
   Select Case vData
      Case "1", "Private":  vData = "Private"
      Case "2", "Public":  vData = "Public"
      Case "3", "Friend":  vData = "Friend"
      Case Else
         If Len(vData) <= 2 Then vData = "Outro(" & vData & ")"
   End Select
   mvarEscopo = vData
End Property
Public Property Get Escopo() As String
    Escopo = mvarEscopo
End Property
Public Property Let TIPO(ByVal vData As String)
   mvarTIPO = vData
End Property
Public Property Get TIPO() As String
   TIPO = mvarTIPO
End Property
Public Property Let Com_Set(ByVal vData As String)
    mvarCom_Set = vData
End Property
Public Property Get Com_Set() As String
    Com_Set = mvarCom_Set
End Property
Public Property Let Com_Let(ByVal vData As String)
    mvarCom_Let = vData
End Property
Public Property Get Com_Let() As String
    Com_Let = mvarCom_Let
End Property
Public Property Let Com_Get(ByVal vData As String)
    mvarCom_Get = vData
End Property
Public Property Get Com_Get() As String
    Com_Get = mvarCom_Get
End Property
Public Property Get PagProp() As PROPPAGE
    If mvarPagProp Is Nothing Then
        Set mvarPagProp = New PROPPAGE
        mvarPagProp.TipoPagina = tpPropriedade
        mvarPagProp.NOME = mvarNOME
        
        mvarPagProp.TIPO = mvarTIPO
        mvarPagProp.Escopo = mvarEscopo
        mvarPagProp.ValorPadrao = mvarVALOR
'        mvarPagProp.isConst = mvarisConst
'        mvarPagProp.isEnum = (mvarTipo = "Enum")
'        mvarPagProp.isType = (mvarTipo = "Type")
    End If
    Set PagProp = mvarPagProp
End Property
Public Property Set PagProp(vData As PROPPAGE)
    Set mvarPagProp = vData
End Property
Public Property Let FileName(ByVal vData As String)
    mvarFileName = vData
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
