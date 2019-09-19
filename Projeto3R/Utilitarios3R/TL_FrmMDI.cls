VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_FrmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe  As FrmMDI
Attribute mvarMe.VB_VarHelpID = -1
Dim mvarPane  As TL_PaneParam

Dim TLDefault As Object
Private mvarSys      As Object
Public Property Set Sys(ByVal vData As Object)
    Set mvarSys = vData
    Set TLDefault.Sys = mvarSys
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Public Sub Show()
   '* Carregar Form
   mvarMe.Show vbModal
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmMDI
   Set TLDefault = New TL_ParamCom
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
   'Call SetMDI(mvarMe.hWnd, mvarSys.MDI.hWnd)
   Set mvarPane = New TL_PaneParam
   Set mvarPane.Sys = mvarSys
   Set mvarPane.MDI = mvarMe
   Set mvarSys.ObjAux = mvarMe
   mvarPane.Show
   
   Call PosicionaForm
   Screen.MousePointer = vbDefault
End Sub
Private Sub PosicionaForm()
   Dim Left As Long
   Dim Top As Long
   Dim Right As Long
   Dim Bottom As Long
   
   mvarSys.MDI.CommandBars.GetClientRect Left, Top, Right, Bottom
   mvarSys.MDI.CommandBars(1).GetWindowRect Left, Top, Right, Bottom
      
   'mvarMe.Left = mvarSys.MDI.Left + 60
   mvarMe.Top = Bottom
   'mvarMe.Width = mvarSys.MDI.Width - mvarMe.Left - 120
   'mvarMe.Height = mvarSys.MDI.Height - mvarMe.Top - 420
   
   
   mvarMe.Left = mvarSys.MDI.Left + 60
   'mvarMe.Top = mvarSys.MDI.Top + 780
   mvarMe.Width = mvarSys.MDI.Width - mvarMe.Left - 120
   mvarMe.Height = mvarSys.MDI.Height - mvarMe.Top - 420
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
   
   mvarMe.CmdSair.Left = (mvarMe.Width / 2) + ((mvarMe.CmdSair.Width) * 4 / 2)  ' 10000
   mvarMe.CmdPadrao.Left = mvarMe.CmdSair.Left - mvarMe.CmdSair.Width - 120
   mvarMe.CmdCancelar.Left = mvarMe.CmdPadrao.Left - mvarMe.CmdPadrao.Width - 120
   mvarMe.CmdOk.Left = mvarMe.CmdCancelar.Left - mvarMe.CmdCancelar.Width - 270
End Sub
