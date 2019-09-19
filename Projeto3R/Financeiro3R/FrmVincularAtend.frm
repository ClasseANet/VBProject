VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.Controls.v11.2.2.ocx"
Object = "{3F666E72-7F79-447A-BCFF-C42C44E3DBE7}#1.0#0"; "MCIGrid.ocx"
Begin VB.Form FrmVincularAtend 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "   Atendimentos Vinculados"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12375
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   12375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton CmdCancel 
      Height          =   375
      Left            =   8760
      TabIndex        =   0
      Top             =   4080
      Width           =   1455
      _Version        =   720898
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Cancelar"
      ForeColor       =   192
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdOk 
      Default         =   -1  'True
      Height          =   375
      Left            =   10560
      TabIndex        =   1
      Top             =   4080
      Width           =   1455
      _Version        =   720898
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Ok"
      ForeColor       =   12582912
      UseVisualStyle  =   -1  'True
   End
   Begin iGrid251_75B4A91C.iGrid GrdVenda 
      Height          =   3495
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   6165
      BorderStyle     =   1
      HighlightBackColorNoFocus=   14737632
   End
   Begin iGrid251_75B4A91C.iGrid GrdAtend 
      Height          =   3495
      Left            =   5400
      TabIndex        =   3
      Top             =   360
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   6165
      BorderStyle     =   1
      HighlightBackColorNoFocus=   14737632
   End
   Begin VB.Label LblAtend 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Itens de &Atendimento Vinculados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5520
      TabIndex        =   5
      Top             =   120
      Width           =   2805
   End
   Begin VB.Label LblVenda 
      Caption         =   "Itens de &Venda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "FrmVincularAtend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Event Active()
Event Load()
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event CmdCancelClick()
Event CmdOkClick()
Event GrdVendaCurCellChange(ByVal lRow As Long, ByVal lCol As Long)
Event GrdAtendRequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
Event GrdAtendMouseEnter(ByVal lRow As Long, ByVal lCol As Long)
Event GrdAtendMouseLeave(ByVal lRow As Long, ByVal lCol As Long)
Event GrdAtendMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
Private Sub CmdCancel_Click()
   RaiseEvent CmdCancelClick
End Sub
Private Sub CmdOk_Click()
   RaiseEvent CmdOkClick
End Sub
Private Sub Form_Activate()
   RaiseEvent Active
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub Form_Load()
   RaiseEvent Load
End Sub
Private Sub GrdAtend_MouseEnter(ByVal lRow As Long, ByVal lCol As Long)
   RaiseEvent GrdAtendMouseEnter(lRow, lCol)
End Sub
Private Sub GrdAtend_MouseLeave(ByVal lRow As Long, ByVal lCol As Long)
   RaiseEvent GrdAtendMouseLeave(lRow, lCol)
End Sub

Private Sub GrdAtend_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
   RaiseEvent GrdAtendMouseUp(Button, Shift, x, y, lRow, lCol, bDoDefault)
End Sub

Private Sub GrdAtend_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
   RaiseEvent GrdAtendRequestEdit(lRow, lCol, iKeyAscii, bCancel, sText, lMaxLength, eTextEditOpt)
End Sub

Private Sub GrdVenda_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)
   RaiseEvent GrdVendaCurCellChange(lRow, lCol)
End Sub
