VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_Consulta3R"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private WithEvents mvarMe  As FrmConsulta3R
Attribute mvarMe.VB_VarHelpID = -1
Dim mvarPane  As TL_ShortBar

Private mvarsys      As Object
Private mvarItem     As Integer
Private mvarIDMODU   As String
Private mvarTLOnFocus  As Object
Public Property Set TLOnFocus(ByVal vData As Object)
   Set mvarTLOnFocus = vData
End Property
Public Property Get TLOnFocus() As Object
   Set TLOnFocus = mvarTLOnFocus
End Property
Public Property Let Item(ByVal vData As Integer)
   mvarItem = vData
End Property
Public Property Get Item() As Integer
   Item = mvarItem
End Property
Public Property Let IDMODU(ByVal vData As String)
   mvarIDMODU = vData
End Property
Public Property Get IDMODU() As String
   IDMODU = mvarIDMODU
End Property

Public Property Set Sys(ByVal vData As Object)
    Set mvarsys = vData
End Property
Public Property Get Sys() As Object
    Set Sys = mvarsys
End Property
Public Sub Show()
   '* Carregar Form
   mvarMe.Show vbModeless
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmConsulta3R
End Sub
Private Sub Class_Terminate()
   On Error Resume Next
   'Unload mvarPane
   'Unload mvarTLOnFocus
   'Unload mvarMe
   'Set mvarPane = Nothing
   'Set mvarTLOnFocus = Nothing
   'Set mvarMe = Nothing
End Sub
Private Sub mvarMe_Activate()
   mvarMe_Resize
   If GetTag(mvarMe, "1Vez", 0) = 0 Then
      Call SetTag(mvarMe, "1Vez", 1)
      Call PosicionaForm
   End If
  
   Screen.MousePointer = vbDefault
End Sub
Private Sub mvarMe_CmdSairClick()
   Unload mvarMe
End Sub
Private Sub mvarMe_Load()
'   Call SetMDI(mvarMe.hwnd, mvarSys.MDI.hwnd)


'   Set mvarSys.ObjAux = mvarMe
'   If Not mvarPane Is Nothing Then
'      mvarItem = mvarPane.Item
'   End If
'   Set mvarPane = New TL_ShortBar
'   With mvarPane
'      Set .Sys = mvarSys
'      Set .MDI = mvarMe
'      .Item = mvarItem
'      .Show
'   End With
'   Call PosicionaForm
   Screen.MousePointer = vbDefault
End Sub
Private Sub PosicionaForm()
   Dim Left As Long
   Dim Top As Long
   Dim Right As Long
   Dim Bottom As Long
   On Error Resume Next
   mvarsys.MDI.CommandBars.GetClientRect Left, Top, Right, Bottom
   mvarsys.MDI.CommandBars(1).GetWindowRect Left, Top, Right, Bottom
      
   mvarMe.Top = Bottom
   
   mvarMe.Left = mvarsys.MDI.Left + 60
   'mvarMe.Top = mvarSys.MDI.Top + 780
   mvarMe.Width = mvarsys.MDI.Width - mvarMe.Left - 120
   mvarMe.Height = mvarsys.MDI.Height - mvarMe.Top - 420
End Sub

Private Sub mvarMe_Resize()
   mvarMe.PictBottom.BorderStyle = 0
   mvarMe.PictBottom.Height = 675

   mvarMe.GrpBoxBottom.Top = 0
   mvarMe.GrpBoxBottom.Left = 0
   mvarMe.GrpBoxBottom.Height = mvarMe.PictBottom.Height
   mvarMe.GrpBoxBottom.Width = mvarMe.PictBottom.Width

   mvarMe.TabPgBotton.Visible = True
   mvarMe.TabPgBotton.Top = 120
   mvarMe.TabPgBotton.Left = 0
   mvarMe.TabPgBotton.Height = mvarMe.PictBottom.Height
   mvarMe.TabPgBotton.Width = mvarMe.PictBottom.Width

   mvarMe.CmdSair.Left = mvarMe.Width - mvarMe.CmdSair.Width - 360
End Sub

