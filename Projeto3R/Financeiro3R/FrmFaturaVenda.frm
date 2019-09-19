VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "COD063~1.OCX"
Object = "{3F666E72-7F79-447A-BCFF-C42C44E3DBE7}#1.0#0"; "MCIGrid.ocx"
Begin VB.Form FrmFaturaVenda 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  Faturas da Venda: 000000"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5550
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin iGrid251_75B4A91C.iGrid GrdFaturas 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   4048
      BorderStyle     =   1
      HighlightBackColorNoFocus=   14737632
   End
   Begin XtremeSuiteControls.PushButton CmdCancel 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
      _Version        =   720898
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Cancelar"
      ForeColor       =   192
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdOk 
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
      _Version        =   720898
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Ok"
      ForeColor       =   12582912
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdExcluir 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Excluir"
      Top             =   2520
      Width           =   1215
      _Version        =   720898
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Excluir"
      ForeColor       =   64
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmFaturaVenda.frx":0000
   End
End
Attribute VB_Name = "FrmFaturaVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'http://www.guaru.com.br/sistemas/document/pdvtef_06.asp
Option Explicit
Event Activate()
Event Load()
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event CmdCancelClick()
Event CmdOkClick()
Event CmdExcluirClick()
Private Sub CmdCancel_Click()
   RaiseEvent CmdCancelClick
End Sub
Private Sub CmdExcluir_Click()
   RaiseEvent CmdExcluirClick
End Sub
Private Sub CmdOk_Click()
   RaiseEvent CmdOkClick
End Sub
Private Sub Form_Activate()
   RaiseEvent Activate
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub Form_Load()
   RaiseEvent Load
End Sub


