VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mvarMe  As Frm000
Attribute mvarMe.VB_VarHelpID = -1
Private mvarSys      As Object
Public Property Set Sys(ByVal vData As Object)
    Set mvarSys = vData
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Public Sub Show(Optional Reload As Boolean = False)
   Load mvarMe
   Call AcoplarForm(mvarMe, 2, mvarSys)
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New Frm000
End Sub
Private Sub Class_Terminate()
   On Error Resume Next
   Unload mvarMe
   Set mvarMe = Nothing
End Sub
Private Sub mvarMe_Load()
   Call SetMDI(mvarMe.hwnd, mvarSys.MDI.hwnd)
   With mvarMe
      
   End With
End Sub
Private Sub mvarMe_Resize()
   With mvarMe
   
   End With
End Sub
