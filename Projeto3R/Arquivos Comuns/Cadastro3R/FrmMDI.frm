VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "COD063~1.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#11.2#0"; "Codejock.DockingPane.v11.2.2.ocx"
Begin VB.Form FrmMDI 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cadastro do Sistema"
   ClientHeight    =   7995
   ClientLeft      =   330
   ClientTop       =   1815
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PictBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   8700
      TabIndex        =   4
      Top             =   7320
      Width           =   8730
      Begin XtremeSuiteControls.GroupBox GrpBoxBottom 
         Height          =   975
         Left            =   0
         TabIndex        =   5
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
         Begin XtremeSuiteControls.PushButton CmdExcluir 
            Height          =   375
            Left            =   2640
            TabIndex        =   1
            Top             =   240
            Width           =   1335
            _Version        =   720898
            _ExtentX        =   2355
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "&Excluir"
            ForeColor       =   192
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton CmdNovo 
            Height          =   375
            Left            =   840
            TabIndex        =   0
            Top             =   240
            Width           =   1335
            _Version        =   720898
            _ExtentX        =   2355
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "&Novo"
            ForeColor       =   12582912
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton CmdEditar 
            Height          =   375
            Left            =   4200
            TabIndex        =   2
            Top             =   240
            Width           =   1455
            _Version        =   720898
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "&Editar"
            ForeColor       =   4210752
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.TabControlPage TabPgBotton 
            Height          =   855
            Left            =   -1200
            TabIndex        =   6
            Top             =   120
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
      Top             =   480
      _Version        =   720898
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "FrmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event CmdExcluirClick()
Public Event CmdEditarClick()
Event CmdNovoClick()
Event CmdSairClick()
Event Activate()
Event Load()
Event Resize()
Public Sub Editar()
   RaiseEvent CmdEditarClick
End Sub
Public Sub Excluir()
   RaiseEvent CmdExcluirClick
End Sub
Public Sub Novo()
   RaiseEvent CmdNovoClick
End Sub
Private Sub CmdExcluir_Click()
   RaiseEvent CmdExcluirClick
End Sub
Private Sub CmdNovo_Click()
   RaiseEvent CmdNovoClick
End Sub
Private Sub CmdEditar_Click()
   RaiseEvent CmdEditarClick
End Sub
Private Sub CmdSair_Click()
   RaiseEvent CmdSairClick
End Sub
Private Sub Form_Activate()
'   Call MontarPanes
   RaiseEvent Activate
   
End Sub
Private Sub Form_Load()
   Call MontarPanes
   RaiseEvent Load
End Sub
Private Sub MontarPanes()
   Dim xPane As Pane
   Dim A As Pane
   Dim B As Pane
   Dim C As Pane
   
   Dim gPanes  As Integer
   gPanes = 2
   
   
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
