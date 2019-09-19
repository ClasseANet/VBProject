VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_FIN000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mvarMe  As FrmFIN00
Attribute mvarMe.VB_VarHelpID = -1
Private mvarSys      As Object
Public Property Set Sys(ByVal vData As Object)
    Set mvarSys = vData
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Public Sub Show(Optional Reload As Boolean = False)
   mvarMe.Show
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmFIN000
End Sub
Private Sub Class_Terminate()
   On Error Resume Next
   Unload mvarMe
   Set mvarMe = Nothing
End Sub
Private Sub mvarMe_Activate()
   mvarMe.CommandBars.DeleteAll
End Sub
Private Sub mvarMe_Load()
   Call SetMDI(mvarMe.hwnd, mvarSys.MDI.hwnd)
   With mvarMe
      .CommandBars.DeleteAll
   End With
   Call AcoplarForm(mvarMe, 2, mvarSys)
End Sub
Private Sub mvarMe_Resize()
   With mvarMe
   
   End With
End Sub
