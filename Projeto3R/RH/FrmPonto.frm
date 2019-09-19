VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "COD063~1.OCX"
Object = "{3F666E72-7F79-447A-BCFF-C42C44E3DBE7}#1.0#0"; "MCIGrid.ocx"
Begin VB.Form FrmPonto 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Ponto "
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   9315
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ComboBox CmbLoja 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   2175
      _Version        =   720898
      _ExtentX        =   3836
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.PushButton CmdSair 
      CausesValidation=   0   'False
      Height          =   375
      Left            =   6480
      TabIndex        =   1
      Top             =   8160
      Width           =   1215
      _Version        =   720898
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Sair"
      UseVisualStyle  =   -1  'True
   End
   Begin iGrid251_75B4A91C.iGrid GrdPonto 
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   12938
   End
   Begin XtremeSuiteControls.ComboBox CmbFunc 
      Height          =   315
      Left            =   2400
      TabIndex        =   3
      Top             =   240
      Width           =   2175
      _Version        =   720898
      _ExtentX        =   3836
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox CmbAno 
      Height          =   315
      Left            =   4680
      TabIndex        =   4
      Top             =   240
      Width           =   2175
      _Version        =   720898
      _ExtentX        =   3836
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox CmbMes 
      Height          =   315
      Left            =   6960
      TabIndex        =   5
      Top             =   240
      Width           =   2175
      _Version        =   720898
      _ExtentX        =   3836
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.PushButton CmdRefresh 
      CausesValidation=   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   8160
      Width           =   1215
      _Version        =   720898
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Atualizar"
      UseVisualStyle  =   -1  'True
   End
End
Attribute VB_Name = "FrmPonto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Activate()
Event Load()
Event CmdSairClick()
Event CmbAnoClick()
Event CmbFuncClick()
Event CmbLojaClick()
Event CmbMesClick()
Event CmdRefreshClick()
Private Sub CmbAno_Click()
    RaiseEvent CmbAnoClick
End Sub
Private Sub CmbFunc_Click()
    RaiseEvent CmbFuncClick
End Sub
Private Sub CmbLoja_Click()
    RaiseEvent CmbLojaClick
End Sub
Private Sub CmbMes_Click()
    RaiseEvent CmbMesClick
End Sub
Private Sub CmdRefresh_Click()
   RaiseEvent CmdRefreshClick
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
