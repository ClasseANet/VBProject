VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPanels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public PanelText            As String
Public ToolTipTxt           As String
Public ClientLeft           As Long
Public ClientTop            As Long
Public ClientWidth          As Long
Public ClientHeight         As Long
Public pEnabled             As Boolean
Public PanelPicture         As StdPicture

Public Property Get PanelIcon() As StdPicture
    Set PanelIcon = PanelPicture
End Property

Public Property Set PanelIcon(ByVal NewPanelPicture As StdPicture)
    Set PanelPicture = NewPanelPicture
End Property

Private Sub Class_Terminate()
    Set PanelPicture = Nothing
End Sub
