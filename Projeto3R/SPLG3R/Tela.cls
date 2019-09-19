VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMe As frmSplash
Attribute mvarMe.VB_VarHelpID = -1
Dim bWait As Boolean
Public Sub Show()
   mvarMe.Show vbModeless
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New frmSplash
End Sub
Private Sub Class_Terminate()
   Unload mvarMe
End Sub
Private Sub mvarMe_Activate()
   Dim sFile As String
   If GetTag(mvarMe, "1VEZ", 1) = 1 Then
      mvarMe.Refresh
      Call SetTag(mvarMe, "1VEZ", 0)
      mvarMe.ZOrder 0
      Call SetTopMostWindow(mvarMe.hWnd, True)
      bWait = True
      
      sFile = App.Path & "\Loader.gif"
      Call ExtractResData("LOADER", "GIF", sFile)
      mvarMe.AniGif1.LoadFile sFile, False
      mvarMe.AniGif1.Visible = True
      Call ExcluirArquivo(App.Path & "\Loader.gif")
      Wait 1
      'While bWait
      '   DoEvents
      'Wend
   End If
End Sub
Private Sub mvarMe_KeyPress(KeyAscii As Integer)
'   Unload mvarMe
End Sub
Private Sub mvarMe_LblNMPROJClick()
   bWait = False
End Sub
Private Sub mvarMe_Load()
   Call mvarMe_Resize
   
   'Call MakeTransparent(mvarMe.hWnd, 100)
   Dim i As Integer
   'Ex: all transparent at ratio 140/255
   'ActiveTransparency Me, True, False, 140, Me.BackColor
   'Ex: Form transparent, visible component at ratio 140/255
   ActiveTransparency mvarMe, True, True, 140, mvarMe.BackColor
    
   'Example display the form transparency degradation
   'ActiveTransparency Me, True, False, 0
   'Me.Show
   'For i = 0 To 255 Step 3
   '   ActiveTransparency Me, True, False, i
   '   Me.Refresh
   'Next i
   
   
   
   'mvarMe.lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
   'mvarMe.lblProductName.Caption = App.Title
End Sub

Private Sub mvarMe_Resize()
   With mvarMe
      .Top = (Screen.Height / 4) - (mvarMe.Height / 2)
      .Left = (Screen.Width / 2) - (mvarMe.Width / 2)
   End With
End Sub
