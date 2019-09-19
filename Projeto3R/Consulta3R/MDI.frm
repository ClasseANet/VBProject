VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "COD063~1.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#11.2#0"; "Codejock.CommandBars.v11.2.2.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#11.2#0"; "Codejock.DockingPane.v11.2.2.ocx"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#11.2#0"; "Codejock.SkinFramework.v11.2.2.ocx"
Begin VB.MDIForm MDI 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000F&
   Caption         =   "Módulo Gerencial"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "MDI.frx":0000
   LinkTopic       =   "MDI"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picHolder 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   4680
      TabIndex        =   0
      Top             =   2550
      Width           =   4680
      Begin VB.Timer Timer 
         Interval        =   500
         Left            =   2640
         Top             =   0
      End
      Begin XtremeSuiteControls.ProgressBar ProgressBar 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1815
         _Version        =   720898
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   93
         Text            =   "Loading..."
         ForeColor       =   12632256
         BackColor       =   -2147483636
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Scrolling       =   1
         Appearance      =   4
         FlatStyle       =   -1  'True
         BarColor        =   -2147483636
      End
      Begin XtremeSuiteControls.ProgressBar ProgressBar2 
         Height          =   255
         Left            =   7920
         TabIndex        =   2
         Top             =   120
         Width           =   1815
         _Version        =   720898
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   93
         Text            =   "Loading..."
         ForeColor       =   12632256
         BackColor       =   -2147483636
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Scrolling       =   2
         Appearance      =   4
         UseVisualStyle  =   -1  'True
         FlatStyle       =   -1  'True
         BarColor        =   -2147483636
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Left            =   3720
         TabIndex        =   4
         Top             =   120
         Width           =   1815
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label LblPercentual 
         Height          =   255
         Left            =   2040
         TabIndex        =   3
         Top             =   120
         Width           =   585
         _Version        =   720898
         _ExtentX        =   1032
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "100%"
         ForeColor       =   12632256
         BackColor       =   -2147483636
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Transparent     =   -1  'True
      End
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework 
      Left            =   2880
      Top             =   120
      _Version        =   720898
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane DockingPaneManager 
      Left            =   0
      Top             =   0
      _Version        =   720898
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   360
      Top             =   0
      _Version        =   720898
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "MDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Event Activate()
Event Load()
Private Sub MDIForm_Activate()
   RaiseEvent Activate
End Sub
Private Sub MDIForm_Initialize()
   On Error GoTo TrataErro
   Call InitCommonControls
   Exit Sub
TrataErro:
   MsgBox Err & " - " & Error, vbOKOnly + vbCritical, "Atenção!"
End Sub
Private Sub MDIForm_Load()
   RaiseEvent Load
End Sub
