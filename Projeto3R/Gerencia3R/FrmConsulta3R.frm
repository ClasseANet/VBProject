VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.Controls.v11.2.2.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#11.2#0"; "Codejock.DockingPane.v11.2.2.ocx"
Begin VB.Form FrmConsulta3R 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Consultas Gerenciais"
   ClientHeight    =   7215
   ClientLeft      =   3690
   ClientTop       =   1155
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PictBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   10170
      TabIndex        =   0
      Top             =   6540
      Width           =   10200
      Begin XtremeSuiteControls.GroupBox GrpBoxBottom 
         Height          =   975
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   8175
         _Version        =   720898
         _ExtentX        =   14420
         _ExtentY        =   1720
         _StockProps     =   79
         Transparent     =   -1  'True
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton CmdSair 
            Height          =   375
            Left            =   5880
            TabIndex        =   3
            Top             =   240
            Width           =   1455
            _Version        =   720898
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Sai&r"
            ForeColor       =   0
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.TabControlPage TabPgBotton 
            Height          =   855
            Left            =   -960
            TabIndex        =   2
            Top             =   480
            Width           =   8895
            _Version        =   720898
            _ExtentX        =   15690
            _ExtentY        =   1508
            _StockProps     =   1
         End
      End
   End
   Begin XtremeDockingPane.DockingPane DockingPaneManager 
      Left            =   840
      Top             =   0
      _Version        =   720898
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "FrmConsulta3R"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public nDoc As Long
Event CmdSairClick()
Event Activate()
Event Load()
Event Resize()
Private Sub CmdSair_Click()
   RaiseEvent CmdSairClick
End Sub
Private Sub Form_Activate()
   'Call MontarPanes
   RaiseEvent Activate
End Sub
Private Sub Form_Load()
   Call MontarPanes
   RaiseEvent Load
   
   Me.Caption = "Consulta " & nDoc
End Sub
Private Sub MontarPanes()
   Dim xPane As Pane
   Dim A As Pane
   Dim B As Pane
   Dim C As Pane
   
   Dim gPanes  As Integer
   gPanes = 2
Exit Sub
   
   With Me.DockingPaneManager
      .DestroyAll
      If gPanes = 1 Then
         Set A = .CreatePane(1, 200, 40, DockTopOf)
         A.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable

      ElseIf gPanes = 2 Then
            Set A = .CreatePane(1, 200, 120, DockLeftOf, Nothing)
            A.Tag = 1
            A.TabColor = vbRed
            A.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable
            
            Set B = .CreatePane(2, 700, 400, DockRightOf, A)
            B.Tag = 2
            B.TabColor = vbBlue
            B.Options = PaneNoCloseable + PaneNoFloatable + PaneNoHideable
            
      ElseIf gPanes = 3 Then
         Set A = .CreatePane(1, 200, 120, DockLeftOf, Nothing)
         A.Tag = 1
         
         Set B = .CreatePane(2, 700, 400, DockRightOf, A)
         B.Tag = 2
         
         Set C = .CreatePane(3, 400, 10, DockBottomOf, B)
         C.Tag = 3
         
      ElseIf gPanes = 4 Then
      
      End If
      .Options.HideClient = True
      .PaintManager.ShowCaption = False
   End With
End Sub
Private Sub Form_Resize()
   RaiseEvent Resize
End Sub

