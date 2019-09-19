VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Set_Constru"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Private mvarAppName As Boolean
Private mvarLoadIni As Boolean
Private mvarExibeSubPasta As Boolean
Private mvarSalvarOnLine As Boolean
Public Property Let SalvarOnLine(ByVal vData As Boolean)
    mvarSalvarOnLine = vData
End Property
Public Property Get SalvarOnLine() As Boolean
    SalvarOnLine = mvarSalvarOnLine
End Property
Public Property Let ExibeSubPasta(ByVal vData As Boolean)
    mvarExibeSubPasta = vData
End Property
Public Property Get ExibeSubPasta() As Boolean
    ExibeSubPasta = mvarExibeSubPasta
End Property
Public Property Let LoadIni(ByVal vData As Boolean)
    mvarLoadIni = vData
End Property
Public Property Get LoadIni() As Boolean
    LoadIni = mvarLoadIni
End Property
Public Property Get AppName() As String
    AppName = "VbEditorUtil Construtor"
End Property

