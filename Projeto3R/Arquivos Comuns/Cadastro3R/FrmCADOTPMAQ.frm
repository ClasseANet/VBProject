VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "COD063~1.OCX"
Object = "{3F666E72-7F79-447A-BCFF-C42C44E3DBE7}#1.0#0"; "MCIGrid.ocx"
Begin VB.Form FrmCADOTPMAQ 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Tipo de Máquina"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   2175
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   5775
      _Version        =   720898
      _ExtentX        =   10186
      _ExtentY        =   3836
      _StockProps     =   79
      Caption         =   " Numeração "
      UseVisualStyle  =   -1  'True
      Appearance      =   4
      Begin iGrid251_75B4A91C.iGrid GrdTrat 
         Height          =   1455
         Left            =   2280
         TabIndex        =   17
         Top             =   600
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   2566
         Appearance      =   0
      End
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   855
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   1695
         _Version        =   720898
         _ExtentX        =   2990
         _ExtentY        =   1508
         _StockProps     =   79
         Caption         =   "Ordenação "
         UseVisualStyle  =   -1  'True
         Appearance      =   4
         Begin XtremeSuiteControls.RadioButton OptNumeracao 
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   15
            Top             =   240
            Width           =   1215
            _Version        =   720898
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Crescente"
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton OptNumeracao 
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   16
            Top             =   480
            Width           =   1335
            _Version        =   720898
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Decrescente"
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
         End
      End
      Begin XtremeSuiteControls.CheckBox ChkTPMANIPULO 
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   1935
         _Version        =   720898
         _ExtentX        =   3413
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Por Tipo de Manípulo"
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Value           =   1
      End
      Begin XtremeSuiteControls.FlatEdit TxtNUMTRATAMENTO 
         Height          =   1170
         Left            =   2280
         TabIndex        =   12
         Top             =   600
         Width           =   3345
         _Version        =   720898
         _ExtentX        =   5900
         _ExtentY        =   2064
         _StockProps     =   77
         BackColor       =   -2147483643
         MaxLength       =   50
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   330
         Left            =   2400
         TabIndex        =   13
         Top             =   240
         Width           =   1335
         _Version        =   720898
         _ExtentX        =   2355
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Por Tratamento"
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   630
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   5775
      _Version        =   720898
      _ExtentX        =   10186
      _ExtentY        =   1111
      _StockProps     =   79
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   1
      Begin XtremeSuiteControls.PushButton CmdExcluir 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Width           =   1095
         _Version        =   720898
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "&Excluir"
         ForeColor       =   64
         UseVisualStyle  =   -1  'True
         Picture         =   "FrmCADOTPMAQ.frx":0000
      End
      Begin XtremeSuiteControls.PushButton CmdSair 
         Height          =   375
         Left            =   4440
         TabIndex        =   2
         Top             =   180
         Width           =   1095
         _Version        =   720898
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "&Sair"
         ForeColor       =   0
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton CmdNovo 
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   180
         Width           =   1095
         _Version        =   720898
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "&Novo"
         ForeColor       =   4210752
         UseVisualStyle  =   -1  'True
         Picture         =   "FrmCADOTPMAQ.frx":0ACA
      End
      Begin XtremeSuiteControls.PushButton CmdSalvar 
         Height          =   375
         Left            =   3000
         TabIndex        =   4
         Top             =   180
         Width           =   1095
         _Version        =   720898
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "&Salvar"
         ForeColor       =   32768
         UseVisualStyle  =   -1  'True
         Picture         =   "FrmCADOTPMAQ.frx":0C24
      End
   End
   Begin XtremeSuiteControls.FlatEdit TxtID 
      Height          =   330
      Left            =   960
      TabIndex        =   5
      Top             =   120
      Width           =   615
      _Version        =   720898
      _ExtentX        =   1085
      _ExtentY        =   582
      _StockProps     =   77
      BackColor       =   -2147483643
      Enabled         =   0   'False
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdLov 
      Height          =   330
      Left            =   1560
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   375
      _Version        =   720898
      _ExtentX        =   661
      _ExtentY        =   582
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmCADOTPMAQ.frx":24EE
   End
   Begin XtremeSuiteControls.FlatEdit TxtDSCMAQ 
      Height          =   330
      Left            =   3000
      TabIndex        =   7
      Top             =   120
      Width           =   1560
      _Version        =   720898
      _ExtentX        =   2752
      _ExtentY        =   582
      _StockProps     =   77
      BackColor       =   -2147483643
      MaxLength       =   50
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.Label LblId 
      Height          =   330
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   510
      _Version        =   720898
      _ExtentX        =   900
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Id.:"
   End
   Begin XtremeSuiteControls.Label Label 
      Height          =   330
      Left            =   2160
      TabIndex        =   8
      Top             =   120
      Width           =   735
      _Version        =   720898
      _ExtentX        =   1296
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Código:"
   End
End
Attribute VB_Name = "FrmCADOTPMAQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Activate()
Event Load()
Event CmdSalvarClick()
Event CmdSairClick()
Event CmdNovoClick()
Event CmdExcluirClick()
Event CmdLovClick()
Event CmbIDTPMAQClick()
Private Sub CmbIDTPMAQ_Click()
   RaiseEvent CmbIDTPMAQClick
End Sub
Private Sub CmdExcluir_Click()
   RaiseEvent CmdExcluirClick
End Sub
Private Sub CmdLov_Click()
   RaiseEvent CmdLovClick
End Sub
Private Sub CmdNovo_Click()
   RaiseEvent CmdNovoClick
End Sub
Private Sub CmdSair_Click()
   RaiseEvent CmdSairClick
End Sub
Private Sub CmdSalvar_Click()
   RaiseEvent CmdSalvarClick
End Sub
Private Sub Form_Activate()
   RaiseEvent Activate
End Sub
Private Sub Form_Load()
   RaiseEvent Load
End Sub

Private Sub GrdTrat_BeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
   Dim i As Integer
   Dim bExiste As Boolean
   
   For i = 1 To Me.GrdTrat.RowCount
      If i <> lRow Then
         If Me.GrdTrat.CellValue(i, lCol) = vNewValue - 1 Then
            bExiste = True
            Exit For
         End If
      End If
   Next
   If Not bExiste Then
      eResult = igEditResCancel
   End If
End Sub
