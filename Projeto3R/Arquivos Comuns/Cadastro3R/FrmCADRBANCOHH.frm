VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.Controls.v11.2.2.ocx"
Object = "{3F666E72-7F79-447A-BCFF-C42C44E3DBE7}#1.0#0"; "MCIGrid.ocx"
Begin VB.Form FrmCADRBANCOHH 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Banco de Horas"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   13740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin iGrid251_75B4A91C.iGrid GrdMes 
      Height          =   7455
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   13150
   End
   Begin XtremeSuiteControls.PushButton CmdSair 
      Height          =   375
      Left            =   12240
      TabIndex        =   10
      Top             =   8640
      Width           =   1215
      _Version        =   720898
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Sair"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13455
      _Version        =   720898
      _ExtentX        =   23733
      _ExtentY        =   1508
      _StockProps     =   79
      Appearance      =   6
      Begin XtremeSuiteControls.ComboBox CmbChapa 
         Height          =   315
         Left            =   3480
         TabIndex        =   6
         Top             =   360
         Width           =   8055
         _Version        =   720898
         _ExtentX        =   14208
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Style           =   2
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox CmbAno 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1215
         _Version        =   720898
         _ExtentX        =   2143
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Style           =   2
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox CmbMes 
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   360
         Width           =   1695
         _Version        =   720898
         _ExtentX        =   2990
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Style           =   2
         UseVisualStyle  =   -1  'True
         DropDownItemCount=   12
      End
      Begin XtremeSuiteControls.PushButton CmdCarregar 
         Height          =   375
         Left            =   12000
         TabIndex        =   7
         Top             =   360
         Width           =   1215
         _Version        =   720898
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Carregar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton CmdFunc 
         Height          =   255
         Left            =   3480
         TabIndex        =   5
         Top             =   120
         Width           =   975
         _Version        =   720898
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "&Funcionário:"
         FlatStyle       =   -1  'True
         Appearance      =   1
         MultiLine       =   0   'False
         ImageAlignment  =   4
         BorderGap       =   0
         ImageGap        =   0
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Left            =   1560
         TabIndex        =   3
         Top             =   120
         Width           =   615
         _Version        =   720898
         _ExtentX        =   1085
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Mês:"
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   735
         _Version        =   720898
         _ExtentX        =   1296
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Ano:"
      End
   End
   Begin XtremeSuiteControls.PushButton CmdSalvar 
      Height          =   375
      Left            =   10440
      TabIndex        =   9
      Top             =   8640
      Width           =   1215
      _Version        =   720898
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Salvar"
      ForeColor       =   16711680
      UseVisualStyle  =   -1  'True
   End
End
Attribute VB_Name = "FrmCADRBANCOHH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Activate()
Event Load()
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event CmdFuncClick()
Event CmdCarregarClick()
Event CmdSairClick()
Event CmdSalvarClick()
Event CmbAnoClick()
Event CmbMesClick()
Event CmbChapaClick()
Event GrdMesAfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
Event GrdMesDblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)
Event GrdMesColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
Private Sub CmbAno_Click()
   RaiseEvent CmbAnoClick
End Sub
Private Sub CmbChapa_Click()
   RaiseEvent CmbChapaClick
End Sub
Private Sub CmbMes_Click()
   RaiseEvent CmbMesClick
End Sub
Private Sub CmdCarregar_Click()
   RaiseEvent CmdCarregarClick
End Sub
Private Sub CmdFunc_Click()
 RaiseEvent CmdFuncClick
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
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub Form_Load()
   RaiseEvent Load
End Sub
Private Sub GrdMes_AfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
   RaiseEvent GrdMesAfterCommitEdit(lRow, lCol)
End Sub
Private Sub GrdMes_ColHeaderBeginDrag(ByVal lCol As Long, bCancel As Boolean)
   bCancel = True
End Sub
Private Sub GrdMes_ColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
   RaiseEvent GrdMesColHeaderClick(lCol, bDoDefault, Shift, x, y)
End Sub
Private Sub GrdMes_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)
   RaiseEvent GrdMesDblClick(lRow, lCol, bRequestEdit)
End Sub
Private Sub GrdMes_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
   If lCol = 1 Then
      If Me.GrdMes.CellValue(lRow, "PONTO") = 2 Or Me.GrdMes.CellType(lRow, "PONTO") <> igCellCheck Then
         bCancel = True
      Else
         bCancel = False
      End If
   Else
      bCancel = True
   End If
End Sub
Private Sub GrdMes_Validate(Cancel As Boolean)
   Me.GrdMes.CommitEdit
End Sub
