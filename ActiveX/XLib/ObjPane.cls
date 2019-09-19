VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObjPane"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sub AcoplarForm(pForm As Object, nPane As Integer, pSys As Object, Optional bDefineFoco As Boolean = True, Optional pMDI As Object)
   With pForm
      .BorderStyle = vbBSNone
      .ClipControls = False
      .WindowState = vbMaximized
      '.MaxButton = False
      '.MinButton = False
      '.ShowInTaskbar = False
   End With
   If IsMissing(pMDI) Or pMDI Is Nothing Then
      With pSys.MDI.DockingPaneManager
         If .Panes(nPane).Handle <> pForm.hwnd Then
            Call ClsAmbiente.SetMDI(pForm.hwnd, pSys.MDI.hwnd)
            .Panes(nPane).Handle = pForm.hwnd
         End If
      End With
   Else
      With pMDI.DockingPaneManager
         If .Panes(nPane).Handle <> pForm.hwnd Then
            'Call SetMDI(pForm.hWnd, pSys.MDI.hWnd)
            .Panes(nPane).Handle = pForm.hwnd
         End If
      End With
   End If
   
   '* Definir foco
   With pForm
      If bDefineFoco Then
         On Error Resume Next
         Dim i As Integer
         Dim iTab As Integer
         Dim bAchou As Boolean
         
         For iTab = 0 To .Controls.Count - 1
            bAchou = False
            For i = 0 To .Controls.Count - 1
               If .Controls(i).TabIndex = iTab Then
                  bAchou = .Controls(i).Visible
                  bAchou = bAchou And .Controls(i).Enabled
                  bAchou = bAchou And (Err = 0)
                  If bAchou Then .Controls(i).SetFocus
                  bAchou = bAchou And (Err = 0)
                  If bAchou Then
                     .Controls(i).SetFocus
                     iTab = .Controls.Count
                  End If
                  Exit For
               End If
               Err = 0
            Next
         Next
      End If
   End With
End Sub

