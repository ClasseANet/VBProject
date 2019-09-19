VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "COD063~1.OCX"
Begin VB.Form FrmExportRPS 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Exportar RPS"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton CmdSair 
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Sai&r"
      ForeColor       =   0
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdOk 
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   1440
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Ok"
      ForeColor       =   12582912
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GrpFormato 
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   5415
      _Version        =   720898
      _ExtentX        =   9551
      _ExtentY        =   2143
      _StockProps     =   79
      Caption         =   " Período de Emissão"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.DateTimePicker CmbDTINI 
         Height          =   345
         Left            =   480
         TabIndex        =   3
         Top             =   600
         Width           =   1575
         _Version        =   720898
         _ExtentX        =   2778
         _ExtentY        =   609
         _StockProps     =   68
         Format          =   1
         CurrentDate     =   40356.1843055556
      End
      Begin XtremeSuiteControls.DateTimePicker CmbDTFIM 
         Height          =   345
         Left            =   3240
         TabIndex        =   4
         Top             =   600
         Width           =   1575
         _Version        =   720898
         _ExtentX        =   2778
         _ExtentY        =   609
         _StockProps     =   68
         Format          =   1
         CurrentDate     =   40356.1843055556
      End
      Begin VB.Label LblDTATEND 
         AutoSize        =   -1  'True
         Caption         =   "Data Início"
         Height          =   195
         Left            =   480
         TabIndex        =   6
         Top             =   360
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Fim"
         Height          =   195
         Left            =   3240
         TabIndex        =   5
         Top             =   360
         Width           =   630
      End
   End
End
Attribute VB_Name = "FrmExportRPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Activate()
Event Load()
Event CmdOkClick()
Event CmdSairClick()
Private Sub CmdSair_Click()
   RaiseEvent CmdSairClick
End Sub
Private Sub CmdOk_Click()
   RaiseEvent CmdOkClick
End Sub
Private Sub Form_Activate()
   RaiseEvent Activate
End Sub
Private Sub Form_Load()
   RaiseEvent Load
End Sub
