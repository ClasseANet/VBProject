VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.Controls.v11.2.2.ocx"
Begin VB.Form FrmExportMov 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Exportar Movimento"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4080
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.CheckBox ChkNDOC 
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   3480
      Width           =   3615
      _Version        =   720898
      _ExtentX        =   6376
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Exporta Nº documento"
      UseVisualStyle  =   -1  'True
      Value           =   1
   End
   Begin XtremeSuiteControls.GroupBox GrpFormato 
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   3735
      _Version        =   720898
      _ExtentX        =   6588
      _ExtentY        =   1508
      _StockProps     =   79
      Caption         =   " Formato "
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.RadioButton OptFormato 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1095
         _Version        =   720898
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   ".QIF"
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton OptFormato 
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   8
         Top             =   360
         Width           =   1095
         _Version        =   720898
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   ".OFC"
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.PushButton CmdExportar 
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   3960
      Width           =   1335
      _Version        =   720898
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Exportar"
      ForeColor       =   16711680
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ListBox LstContas 
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3855
      _Version        =   720898
      _ExtentX        =   6800
      _ExtentY        =   2355
      _StockProps     =   77
      BackColor       =   -2147483643
      Sorted          =   -1  'True
      Appearance      =   4
      UseVisualStyle  =   -1  'True
      Style           =   1
   End
   Begin XtremeSuiteControls.DateTimePicker CmbDTINI 
      Height          =   345
      Left            =   120
      TabIndex        =   3
      Top             =   2040
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
      Left            =   2400
      TabIndex        =   5
      Top             =   2040
      Width           =   1575
      _Version        =   720898
      _ExtentX        =   2778
      _ExtentY        =   609
      _StockProps     =   68
      Format          =   1
      CurrentDate     =   40356.1843055556
   End
   Begin XtremeSuiteControls.PushButton CmdSair 
      Height          =   375
      Left            =   2400
      TabIndex        =   11
      Top             =   3960
      Width           =   1335
      _Version        =   720898
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Sair"
      ForeColor       =   128
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.CommonDialog CommonDialog1 
      Left            =   1920
      Top             =   4200
      _Version        =   720898
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   4
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Contas"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Data Fim"
      Height          =   195
      Left            =   2400
      TabIndex        =   4
      Top             =   1800
      Width           =   630
   End
   Begin VB.Label LblDTATEND 
      AutoSize        =   -1  'True
      Caption         =   "Data Início"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   795
   End
End
Attribute VB_Name = "FrmExportMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Load()
Event Activate()
Event CmdExportarClick()
Event CmdSairClick()
Event LstContasItemCheck(ByVal Item As Long)
Private Sub CmdExportar_Click()
   RaiseEvent CmdExportarClick
End Sub
Private Sub CmdSair_Click()
   RaiseEvent CmdSairClick
End Sub
Private Sub Form_Activate()
   RaiseEvent Activate
End Sub
Private Sub Form_Load()
   RaiseEvent Load
End Sub

Private Sub LstContas_Click()
   Me.LstContas.Checked(Me.LstContas.ListIndex) = True
   'RaiseEvent LstContasItemCheck(Me.LstContas.ListIndex)
End Sub

Private Sub LstContas_ItemCheck(ByVal Item As Long)
   RaiseEvent LstContasItemCheck(Item)
End Sub
