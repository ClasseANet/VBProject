VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMETODO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"clsVARIAVEL"
Private mvarNOME As String
Private mvarVARIAVEL As clsVARIAVEL
Private mvarPagProp As PROPPAGE
Private mvarTipo As String
Private mvarStartLine As Integer
Private mvarFileName As String
Public Property Let FileName(ByVal vData As String)
    mvarFileName = vData
End Property
Public Property Get FileName() As String
    FileName = mvarFileName
End Property
Public Property Let StartLine(ByVal vData As Integer)
    mvarStartLine = vData
End Property
Public Property Get StartLine() As Integer
    Dim nAux As Long
    Dim sAux As String
    If Not mvarPagProp Is Nothing Then
       sAux = IIf(mvarPagProp.Retorno = "", "Sub ", "Function ") & mvarNOME
       nAux = GetStartLine(mvarFileName, sAux)
    Else
       nAux = GetStartLine(mvarFileName, "Sub " & mvarNOME)
       If nAux = 0 Then nAux = GetStartLine(mvarFileName, "Function " & mvarNOME)
    End If
    nAux = GetStartLine(mvarFileName, CStr(sAux))
    If nAux <> mvarStartLine And mvarStartLine <> 0 And nAux <> 0 Then
       MsgBox "Linha Calculada Difere da Atual."
       nAux = nAux
    End If
    StartLine = IIf(nAxu = 0, mvarStartLine, nAux)
End Property
Public Property Let Tipo(ByVal vData As String)
   mvarTipo = vData
End Property
Public Property Get Tipo() As String
   Tipo = mvarTipo
End Property
Public Property Get PagProp() As PROPPAGE
    If Not mvarPagProp Is Nothing Then
       Set PagProp = mvarPagProp
       'Set mvarPagProp = New PROPPAGE
       mvarPagProp.TipoPagina = tpMetodo
       mvarPagProp.NOME = mvarNOME
    End If
End Property
Public Property Set PagProp(vData As PROPPAGE)
    Set mvarPagProp = vData
End Property
Public Property Get VARIAVEL() As clsVARIAVEL
    If mvarVARIAVEL Is Nothing Then
        Set mvarVARIAVEL = New clsVARIAVEL
    End If
    Set VARIAVEL = mvarVARIAVEL
End Property
Public Property Set VARIAVEL(vData As clsVARIAVEL)
    Set mvarVARIAVEL = vData
End Property
Public Property Let NOME(ByVal vData As String)
   Select Case Right(vData, 1)
      Case "!":      vData = Mid(vData, 1, Len(vData) - 1): mvarTipo = "Single"
      Case "#":      vData = Mid(vData, 1, Len(vData) - 1): mvarTipo = "Double"
      Case "$":      vData = Mid(vData, 1, Len(vData) - 1): mvarTipo = "String"
      Case "%":      vData = Mid(vData, 1, Len(vData) - 1): mvarTipo = "Integer"
      Case "&":      vData = Mid(vData, 1, Len(vData) - 1): mvarTipo = "Long"
      Case "(", ")": vData = Mid(vData, 1, Len(vData) - 1): mvarTipo = "Variant"
   End Select
   mvarNOME = vData
End Property
Public Property Get NOME() As String
   NOME = mvarNOME
End Property
Private Sub Class_Initialize()
   Set mvarVARIAVEL = New clsVARIAVEL
   mvarTipo = "Variant"
End Sub
Private Sub Class_Terminate()
    Set mvarVARIAVEL = Nothing
End Sub
