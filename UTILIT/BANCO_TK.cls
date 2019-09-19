VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BANCO_TK"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"TB_CRG_TERMO"
Attribute VB_Ext_KEY = "Member1" ,"TB_DAI"
Attribute VB_Ext_KEY = "Member2" ,"TB_PCG_IMP"
Option Explicit
Private mvardBase As Object
Private mvardBases As New Collection
Private mvarIDPROJ As String
Private mvarSys_Idioma As Integer

Private mvarTB_CRG_TERMO As TB_CRG_TERMO
Private mvarTB_DAI As TB_DAI
Private mvarTB_PCG_IMP As TB_PCG_IMP
Public Property Get TB_PCG_IMP() As TB_PCG_IMP
    If mvarTB_PCG_IMP Is Nothing Then
        Set mvarTB_PCG_IMP = New TB_PCG_IMP
        mvarTB_PCG_IMP.dBase = mvardBase
    End If
    Set TB_PCG_IMP = mvarTB_PCG_IMP
End Property
Public Property Set TB_PCG_IMP(vData As TB_PCG_IMP)
    Set mvarTB_PCG_IMP = vData
End Property
Public Property Get TB_DAI() As TB_DAI
    If mvarTB_DAI Is Nothing Then
        Set mvarTB_DAI = New TB_DAI
        mvarTB_DAI.dBase = mvardBase
    End If
    Set TB_DAI = mvarTB_DAI
End Property
Public Property Set TB_DAI(vData As TB_DAI)
    Set mvarTB_DAI = vData
End Property
Public Property Get TB_CRG_TERMO() As TB_CRG_TERMO
   If mvarTB_CRG_TERMO Is Nothing Then
      Set mvarTB_CRG_TERMO = New TB_CRG_TERMO
      mvarTB_CRG_TERMO.dBase = mvardBase
   End If
   Set TB_CRG_TERMO = mvarTB_CRG_TERMO
End Property
Public Property Set TB_CRG_TERMO(vData As TB_CRG_TERMO)
    Set mvarTB_CRG_TERMO = vData
End Property
Public Property Let dBase(ByVal vData As Object)
   Set mvardBase = vData
End Property
Public Property Get dBase() As Object
   Set dBase = mvardBase
End Property
'Public Property Set dBases(ByVal vData As Object)
'    Set mvardBases = vData
'End Property
'Public Property Get dBases() As Collection
'    Set dBases = mvardBases
'End Property
Public Property Let Sys_Idioma(ByVal vData As Integer)
   mvarSys_Idioma = vData
End Property
Public Property Get Sys_Idioma() As Integer
   Sys_Idioma = mvarSys_Idioma
End Property
Private Sub Class_Terminate()
  Set mvarTB_PCG_IMP = Nothing
  Set mvarTB_DAI = Nothing
  Set mvarTB_CRG_TERMO = Nothing
End Sub
