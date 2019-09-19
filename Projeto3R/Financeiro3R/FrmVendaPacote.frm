VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "COD063~1.OCX"
Object = "{3F666E72-7F79-447A-BCFF-C42C44E3DBE7}#1.0#0"; "MCIGrid.ocx"
Begin VB.Form FrmVendaPacote 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Pacotes"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6615
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrmeProd 
      Caption         =   " Promoções / Descontos"
      Height          =   2835
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6375
      Begin XtremeSuiteControls.FlatEdit TxtVLDESC 
         Height          =   285
         Left            =   4560
         TabIndex        =   3
         Top             =   2115
         Width           =   1185
         _Version        =   720898
         _ExtentX        =   2090
         _ExtentY        =   503
         _StockProps     =   77
         ForeColor       =   192
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0,00"
         Alignment       =   1
         Locked          =   -1  'True
         MaxLength       =   6
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit TxtVLTOTAL 
         Height          =   285
         Left            =   4560
         TabIndex        =   4
         Top             =   1845
         Width           =   1185
         _Version        =   720898
         _ExtentX        =   2090
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483626
         Text            =   "0,00"
         BackColor       =   -2147483626
         Alignment       =   1
         Locked          =   -1  'True
         MaxLength       =   6
         UseVisualStyle  =   -1  'True
      End
      Begin iGrid251_75B4A91C.iGrid GrdPacote 
         Height          =   1575
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   2778
      End
      Begin XtremeSuiteControls.FlatEdit TxtVALOR 
         Height          =   285
         Left            =   4560
         TabIndex        =   6
         Top             =   2385
         Width           =   1185
         _Version        =   720898
         _ExtentX        =   2090
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0,00"
         Alignment       =   1
         Locked          =   -1  'True
         MaxLength       =   6
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label7 
         Height          =   285
         Left            =   3240
         TabIndex        =   10
         Top             =   2385
         Width           =   1050
         _Version        =   720898
         _ExtentX        =   1852
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Valor a Pagar:"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   285
         Left            =   3240
         TabIndex        =   9
         Top             =   1845
         Width           =   1215
         _Version        =   720898
         _ExtentX        =   2143
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Valor Produtos :"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   285
         Left            =   3240
         TabIndex        =   8
         Top             =   2115
         Width           =   1140
         _Version        =   720898
         _ExtentX        =   2011
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Valor Desconto:"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label LblDesc 
         Height          =   285
         Left            =   5760
         TabIndex        =   7
         Top             =   2160
         Width           =   480
         _Version        =   720898
         _ExtentX        =   847
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "18,88%"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Transparent     =   -1  'True
      End
   End
   Begin XtremeSuiteControls.PushButton CmdSair 
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   2400
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Sai&r"
      ForeColor       =   0
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdCancel 
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   3120
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Cancelar"
      ForeColor       =   192
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdOk 
      Height          =   375
      Left            =   5040
      TabIndex        =   11
      Top             =   3120
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Ok"
      ForeColor       =   12582912
      UseVisualStyle  =   -1  'True
   End
End
Attribute VB_Name = "FrmVendaPacote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Activate()
Event Load()
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event CmdOkClick()
Event CmdCancelClick()

Event GrdPacoteAfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
Event GrdPacoteBeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
Event GrdPacoteColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
Event GrdPacoteColHeaderDblClick(ByVal lCol As Long)
Event GrdPacoteMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
Event GrdPacoteLostFocus()
Event GrdPacoteRequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
Event GrdPacoteValidate(Cancel As Boolean)
Private Sub CmdCancel_Click()
   RaiseEvent CmdCancelClick
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
Private Sub GrdPacote_AfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
   RaiseEvent GrdPacoteAfterCommitEdit(lRow, lCol)
End Sub
Private Sub GrdPacote_BeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
   RaiseEvent GrdPacoteBeforeCommitEdit(lRow, lCol, eResult, sNewText, vNewValue, lConvErr)
End Sub
Private Sub GrdPacote_ColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
   RaiseEvent GrdPacoteColHeaderClick(lCol, bDoDefault, Shift, x, y)
End Sub
Private Sub GrdPacote_ColHeaderDblClick(ByVal lCol As Long)
   RaiseEvent GrdPacoteColHeaderDblClick(lCol)
End Sub
Private Sub GrdPacote_LostFocus()
   RaiseEvent GrdPacoteLostFocus
End Sub
Private Sub GrdPacote_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
  RaiseEvent GrdPacoteMouseUp(Button, Shift, x, y, lRow, lCol, bDoDefault)
End Sub
Private Sub GrdPacote_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
   RaiseEvent GrdPacoteRequestEdit(lRow, lCol, iKeyAscii, bCancel, sText, lMaxLength, eTextEditOpt)
End Sub
Private Sub GrdPacote_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)
   Dim i As Integer
   With Me.GrdPacote
      .RowMode = (lRow = .RowCount)
      If .RowCount > 0 Then
         For i = 1 To .ColCount
            If .ColVisible(i) Then
               .CellForeColor(.RowCount, i) = IIf(lRow = .RowCount, vbHighlightText, vbGrayText)
               Exit For
            End If
         Next
      End If
   End With
End Sub
Private Sub GrdPacote_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)
   If (lRow = Me.GrdPacote.RowCount) Then bRequestEdit = False
End Sub
Private Sub GrdPacote_Validate(Cancel As Boolean)
   RaiseEvent GrdPacoteValidate(Cancel)
End Sub
