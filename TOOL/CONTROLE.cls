VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CONTROLE"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Private mvarTIPO As String
Private mvarNOME As String
Private mvarFlag As Boolean
Private mvarNotNull As Boolean
Public Property Let NotNull(ByVal vData As Boolean)
    mvarNotNull = vData
End Property
Public Property Get NotNull() As Boolean
    NotNull = mvarNotNull
End Property
Public Property Let Flag(ByVal vData As Boolean)
    mvarFlag = vData
End Property
Public Property Get Flag() As Boolean
    Flag = mvarFlag
End Property
Public Property Let NOME(ByVal vData As String)
    mvarNOME = vData
End Property
Public Property Get NOME() As String
    NOME = mvarNOME
End Property
Public Property Let Tipo(ByVal vData As String)
    mvarTIPO = vData
    Select Case mvarTIPO
       Case "TextBox":  mvarNOME = "Txt" & Mid(mvarNOME, 4)
       Case "MaskEdit": mvarNOME = "Msk" & Mid(mvarNOME, 4)
       Case "ComboBox": mvarNOME = "Cmb" & Mid(mvarNOME, 4)
       Case "CheckBox": mvarNOME = "Chk" & Mid(mvarNOME, 4)
       Case Else:  mvarNOME = "Txt" & Mid(mvarNOME, 4)
    End Select
End Property
Public Property Get Tipo() As String
    Tipo = mvarTIPO
End Property

Private Sub Class_Initialize()
    mvarTIPO = "TextBox"
End Sub
