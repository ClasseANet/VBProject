VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVARIAVEL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Private mvarTIPO            As String
Private mvarNOME            As String
Private mvarEscopo          As String
Private mvarVALOR           As String
Private mvarPassArg         As String
Private mvarisConst         As Boolean
Private mvarisObjWithEvents As Boolean
Private mvarPagProp         As PROPPAGE
Private mvarStartLine       As Long
Private mvarFileName As String
Public Property Let FileName(ByVal vData As String)
    mvarFileName = vData
End Property
Public Property Get FileName() As String
    FileName = mvarFileName
End Property
Public Property Let StartLine(vData As Long)
   mvarStartLine = vData
End Property
Public Property Get StartLine() As Long
   mvarStartLine = GetStartLine(mvarFileName, mvarNOME)
   StartLine = mvarStartLine
End Property
Public Property Get PagProp() As PROPPAGE
    If mvarPagProp Is Nothing Then
        Set mvarPagProp = New PROPPAGE
        mvarPagProp.TipoPagina = tpVariavel
        mvarPagProp.NOME = mvarNOME
        mvarPagProp.TIPO = mvarTIPO
        mvarPagProp.Escopo = mvarEscopo
        mvarPagProp.ValorPadrao = mvarVALOR
        mvarPagProp.isConst = mvarisConst
        mvarPagProp.isEnum = (mvarTIPO = "Enum")
        mvarPagProp.isType = (mvarTIPO = "Type")
    End If
    Set PagProp = mvarPagProp
End Property
Public Property Set PagProp(vData As PROPPAGE)
    Set mvarPagProp = vData
End Property
Public Property Let isConst(ByVal vData As Boolean)
    mvarisConst = vData
End Property
Public Property Get isConst() As Boolean
    isConst = mvarisConst
End Property
Public Property Let isObjWithEvents(ByVal vData As Boolean)
    mvarisObjWithEvents = vData
End Property
Public Property Get isObjWithEvents() As Boolean
    isObjWithEvents = mvarisObjWithEvents
End Property
Public Property Let PassArg(ByVal vData As String)
    mvarPassArg = vData
End Property
Public Property Get PassArg() As String
    PassArg = mvarPassArg
End Property
Public Property Let VALOR(ByVal vData As String)
    mvarVALOR = vData
End Property
Public Property Get VALOR() As String
    VALOR = mvarVALOR
End Property
Public Property Let Escopo(ByVal vData As String)
'     Case 1: ItemX.SubItems(1) = "'Private'"
'     Case 2: ItemX.SubItems(1) = "'Public'"
'     Case 3: ItemX.SubItems(1) = "'Friend'"
'     Case Else: ItemX.SubItems(1) = "'Outro'"
   Select Case vData
      Case "1", "Private": vData = "Private"
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
Public Property Let NOME(ByVal vData As String)
    mvarNOME = vData
End Property
Public Property Get NOME() As String
    NOME = mvarNOME
End Property
Public Property Let TIPO(ByVal vData As String)
   If InStr(" " & Trim(LCase(vData)), " new ") <> 0 Then
      vData = Mid(vData, InStr(" " & Trim(LCase(vData)), " new ") + 4)
   End If
   mvarTIPO = vData
End Property
Public Property Get TIPO() As String
    TIPO = mvarTIPO
End Property
Private Sub Class_Initialize()
   mvarisConst = False
End Sub
