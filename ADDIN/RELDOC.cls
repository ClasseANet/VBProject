VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRELDOC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"clsMEMBRO"
Private mvarNOME As String
Private mvarMEMBER As clsMEMBRO
Private mvarFileName As String
Public Property Let FileName(ByVal vData As String)
    mvarFileName = vData
End Property
Public Property Get FileName() As String
    FileName = mvarFileName
End Property
Public Property Get MEMBER() As clsMEMBRO
    If mvarMEMBER Is Nothing Then
        Set mvarMEMBER = New clsMEMBRO
    End If
    Set MEMBER = mvarMEMBER
End Property
Public Property Set MEMBER(vData As clsMEMBRO)
    Set mvarMEMBER = vData
End Property
Private Sub Class_Terminate()
    Set mvarMEMBER = Nothing
End Sub
Public Property Let NOME(ByVal vData As String)
   mvarNOME = vData
End Property
Public Property Get NOME() As String
   NOME = mvarNOME
End Property



