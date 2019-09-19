VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.Controls.v11.2.2.ocx"
Object = "{3F666E72-7F79-447A-BCFF-C42C44E3DBE7}#1.0#0"; "MCIGrid.ocx"
Begin VB.Form FrmCadTarefa 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Venda"
   ClientHeight    =   6165
   ClientLeft      =   2595
   ClientTop       =   2760
   ClientWidth     =   9480
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.FlatEdit TxtIDVENDA 
      Height          =   345
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   609
      _StockProps     =   77
      BackColor       =   14737632
      BackColor       =   14737632
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdChave 
      Height          =   495
      Left            =   8520
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   4680
      Visible         =   0   'False
      Width           =   495
      _Version        =   720898
      _ExtentX        =   873
      _ExtentY        =   873
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmCadTarefa.frx":0000
   End
   Begin XtremeSuiteControls.DateTimePicker CmbDTVENDA 
      Height          =   345
      Left            =   3840
      TabIndex        =   3
      Top             =   360
      Width           =   1575
      _Version        =   720898
      _ExtentX        =   2778
      _ExtentY        =   609
      _StockProps     =   68
      Format          =   1
      CurrentDate     =   40356.1843055556
   End
   Begin XtremeSuiteControls.PushButton CmdCancel 
      Height          =   375
      Left            =   5640
      TabIndex        =   26
      Top             =   5640
      Width           =   1455
      _Version        =   720898
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Cancelar"
      ForeColor       =   192
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdOk 
      Height          =   375
      Left            =   7560
      TabIndex        =   25
      Top             =   5640
      Width           =   1455
      _Version        =   720898
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Ok"
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GrpSessao 
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   840
      Width           =   5415
      _Version        =   720898
      _ExtentX        =   9551
      _ExtentY        =   1296
      _StockProps     =   79
      Appearance      =   4
      BorderStyle     =   2
      Begin XtremeSuiteControls.FlatEdit TxtTEL1 
         Height          =   345
         Left            =   4320
         TabIndex        =   8
         Top             =   240
         Width           =   1095
         _Version        =   720898
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   16777215
         BackColor       =   16777215
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit TxtNOME 
         Height          =   345
         Left            =   840
         TabIndex        =   6
         Top             =   240
         Width           =   3015
         _Version        =   720898
         _ExtentX        =   5318
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   16777215
         Text            =   "Patricia Moreira"
         BackColor       =   16777215
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton CmdIDCLIENTE 
         Height          =   345
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   705
         _Version        =   720898
         _ExtentX        =   1244
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Cliente"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         TextAlignment   =   10
         ImageAlignment  =   6
         TextImageRelation=   0
      End
      Begin XtremeSuiteControls.PushButton CmdLovCli 
         Height          =   345
         Left            =   3840
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Width           =   375
         _Version        =   720898
         _ExtentX        =   661
         _ExtentY        =   609
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Picture         =   "FrmCadTarefa.frx":059A
      End
      Begin VB.Label LblTel2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Celular/Tel."
         Height          =   240
         Left            =   4320
         TabIndex        =   7
         Top             =   0
         Width           =   1035
      End
   End
   Begin iGrid251_75B4A91C.iGrid GrdVenda 
      Height          =   1455
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2566
      BorderStyle     =   1
      HighlightBackColorNoFocus=   14737632
   End
   Begin XtremeSuiteControls.GroupBox GrpPedido 
      Height          =   3255
      Left            =   5760
      TabIndex        =   14
      Top             =   2160
      Width           =   3615
      _Version        =   720898
      _ExtentX        =   6376
      _ExtentY        =   5741
      _StockProps     =   79
      Caption         =   "VALORES"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin XtremeSuiteControls.FlatEdit TxtVLVENDA 
         Height          =   345
         Left            =   1560
         TabIndex        =   16
         Top             =   360
         Width           =   1815
         _Version        =   720898
         _ExtentX        =   3201
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
         Locked          =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit TxtVLPGTO 
         Height          =   345
         Left            =   1560
         TabIndex        =   20
         Top             =   1320
         Width           =   1815
         _Version        =   720898
         _ExtentX        =   3201
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   14737632
         BackColor       =   14737632
         Alignment       =   1
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit TxtVLDESC 
         Height          =   345
         Left            =   1560
         TabIndex        =   18
         Top             =   840
         Width           =   1815
         _Version        =   720898
         _ExtentX        =   3201
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
         Locked          =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit TxtVLTROCO 
         Height          =   345
         Left            =   1560
         TabIndex        =   22
         Top             =   1800
         Width           =   1815
         _Version        =   720898
         _ExtentX        =   3201
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   14737632
         BackColor       =   14737632
         Alignment       =   1
         Locked          =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton CmdFechar 
         Height          =   495
         Left            =   960
         TabIndex        =   23
         Top             =   2520
         Width           =   1815
         _Version        =   720898
         _ExtentX        =   3201
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "&Fechar Pagamento"
         ForeColor       =   32768
         BackColor       =   65280
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label LblTroco 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TROCO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   240
         Left            =   60
         TabIndex        =   21
         Top             =   1920
         Width           =   1500
      End
      Begin VB.Label LblDesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DESCONTO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   240
         Left            =   60
         TabIndex        =   17
         Top             =   960
         Width           =   1500
      End
      Begin VB.Label LblPgto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PAGAMENTO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   240
         Left            =   60
         TabIndex        =   19
         Top             =   1440
         Width           =   1500
      End
      Begin VB.Label LblTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   240
         Left            =   60
         TabIndex        =   15
         Top             =   480
         Width           =   1500
      End
   End
   Begin iGrid251_75B4A91C.iGrid GrdPagamento 
      Height          =   1335
      Left            =   120
      TabIndex        =   13
      Top             =   4080
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   2355
      BorderStyle     =   1
      HighlightBackColorNoFocus=   14737632
   End
   Begin iGrid251_75B4A91C.iGrid GrdAtendimento 
      Height          =   1215
      Left            =   5640
      TabIndex        =   29
      Top             =   360
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   2143
      BorderStyle     =   1
      HighlightBackColorNoFocus=   14737632
   End
   Begin XtremeSuiteControls.PushButton CmdLov 
      Height          =   345
      Left            =   1200
      TabIndex        =   31
      TabStop         =   0   'False
      ToolTipText     =   "Pesquisar"
      Top             =   360
      Width           =   375
      _Version        =   720898
      _ExtentX        =   661
      _ExtentY        =   609
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmCadTarefa.frx":071D
   End
   Begin XtremeSuiteControls.PushButton CmdExcluir 
      Height          =   375
      Left            =   3000
      TabIndex        =   27
      ToolTipText     =   "Excluir"
      Top             =   5640
      Width           =   1455
      _Version        =   720898
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Excluir"
      ForeColor       =   64
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmCadTarefa.frx":08A0
   End
   Begin XtremeSuiteControls.PushButton CmdRecibo 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   480
      TabIndex        =   28
      ToolTipText     =   "Excluir"
      Top             =   5640
      Width           =   1455
      _Version        =   720898
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Recibo 000000"
      ForeColor       =   4210688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Picture         =   "FrmCadTarefa.frx":136A
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   120
      X2              =   9360
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label LblAtendimento 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Atendimentos"
      Height          =   240
      Left            =   5640
      TabIndex        =   30
      Top             =   120
      Width           =   1230
   End
   Begin VB.Label LblVenda 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Venda"
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   600
   End
   Begin VB.Label LblItens 
      AutoSize        =   -1  'True
      Caption         =   "Itens de Venda"
      Height          =   240
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   1350
   End
   Begin VB.Label LblPagamento 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pagamento"
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   3840
      Width           =   1050
   End
   Begin VB.Label LblDTATEND 
      AutoSize        =   -1  'True
      Caption         =   "Data"
      Height          =   240
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   435
   End
End
Attribute VB_Name = "FrmCadTarefa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'http://www.guaru.com.br/sistemas/document/pdvtef_06.asp
Option Explicit
Event Activate()
Event Load()
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event Unload(Cancel As Integer)
Event Resize()
Event CmdOkClick()
Event CmdCancelClick()
Event CmdFecharClick()
Event CmdExcluirClick()
Event CmdChaveClick()
Event CmdLovClick()
Event CmdLovCliClick()
Event CmdIDCLIENTEClick()
Event CmdReciboClick()

Event GrdVendaAfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
Event GrdVendaBeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
Event GrdVendaColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
Event GrdVendaColHeaderDblClick(ByVal lCol As Long)
Event GrdVendaMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
Event GrdVendaLostFocus()
Event GrdVendaRequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
Event GrdVendaValidate(Cancel As Boolean)

Event GrdPagamentoAfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
Event GrdPagamentoBeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
Event GrdPagamentoColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
Event GrdPagamentoColHeaderDblClick(ByVal lCol As Long)
Event GrdPagamentoMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
Event GrdPagamentoLostFocus()
Event GrdPagamentoRequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
Event GrdPagamentoValidate(Cancel As Boolean)

Event GrdAtendimentoAfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
Event GrdAtendimentoBeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
Event GrdAtendimentoColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
Event GrdAtendimentoColHeaderDblClick(ByVal lCol As Long)
Event GrdAtendimentoMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
Event GrdAtendimentoLostFocus()
Event GrdAtendimentoRequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
Event GrdAtendimentoValidate(Cancel As Boolean)

Event TxtIDVENDAGotFocus()
Event TxtIDVENDALostFocus()
Event TxtNOMEChange()
Event TxtNOMEKeyPress(KeyAscii As Integer)
Event TxtTEL1Change()
Event TxtTEL1KeyPress(KeyAscii As Integer)
Event TxtTEL1LostFocus()
Event TxtVLDESCChange()
Event TxtVLPGTOChange()
Private Sub cmdCancel_Click()
   RaiseEvent CmdCancelClick
End Sub
Private Sub CmdChave_Click()
   RaiseEvent CmdChaveClick
End Sub
Private Sub CmdExcluir_Click()
   RaiseEvent CmdExcluirClick
End Sub
Private Sub CmdFechar_Click()
   RaiseEvent CmdFecharClick
End Sub
Private Sub CmdIDCLIENTE_Click()
   RaiseEvent CmdIDCLIENTEClick
End Sub
Private Sub CmdLov_Click()
   RaiseEvent CmdLovClick
End Sub
Private Sub CmdLovCli_Click()
   RaiseEvent CmdLovCliClick
End Sub
Private Sub cmdOk_Click()
   RaiseEvent CmdOkClick
End Sub
Private Sub CmdRecibo_Click()
   RaiseEvent CmdReciboClick
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
Private Sub GrdAtendimento_AfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
   RaiseEvent GrdAtendimentoAfterCommitEdit(lRow, lCol)
End Sub
Private Sub GrdAtendimento_BeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
   RaiseEvent GrdAtendimentoBeforeCommitEdit(lRow, lCol, eResult, sNewText, vNewValue, lConvErr)
End Sub
Private Sub GrdAtendimento_ColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
   RaiseEvent GrdAtendimentoColHeaderClick(lCol, bDoDefault, Shift, x, y)
End Sub
Private Sub GrdAtendimento_ColHeaderDblClick(ByVal lCol As Long)
   RaiseEvent GrdAtendimentoColHeaderDblClick(lCol)
End Sub
Private Sub GrdAtendimento_LostFocus()
   RaiseEvent GrdAtendimentoLostFocus
End Sub
Private Sub GrdAtendimento_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
  RaiseEvent GrdAtendimentoMouseUp(Button, Shift, x, y, lRow, lCol, bDoDefault)
End Sub
Private Sub GrdAtendimento_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
   RaiseEvent GrdAtendimentoRequestEdit(lRow, lCol, iKeyAscii, bCancel, sText, lMaxLength, eTextEditOpt)
End Sub
Private Sub GrdAtendimento_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)
   Dim i As Integer
   With Me.GrdAtendimento
      .RowMode = True '(lRow = .RowCount)
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
Private Sub GrdAtendimento_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)
   If (lRow = Me.GrdAtendimento.RowCount) Then bRequestEdit = False
End Sub
Private Sub GrdAtendimento_Validate(Cancel As Boolean)
   RaiseEvent GrdAtendimentoValidate(Cancel)
End Sub
Private Sub GrdPagamento_AfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
   RaiseEvent GrdPagamentoAfterCommitEdit(lRow, lCol)
End Sub
Private Sub GrdPagamento_BeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
   RaiseEvent GrdPagamentoBeforeCommitEdit(lRow, lCol, eResult, sNewText, vNewValue, lConvErr)
End Sub
Private Sub GrdPagamento_ColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
   RaiseEvent GrdPagamentoColHeaderClick(lCol, bDoDefault, Shift, x, y)
End Sub
Private Sub GrdPagamento_ColHeaderDblClick(ByVal lCol As Long)
   RaiseEvent GrdPagamentoColHeaderDblClick(lCol)
End Sub
Private Sub GrdPagamento_LostFocus()
   RaiseEvent GrdPagamentoLostFocus
End Sub
Private Sub GrdPagamento_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
  RaiseEvent GrdPagamentoMouseUp(Button, Shift, x, y, lRow, lCol, bDoDefault)
End Sub
Private Sub GrdPagamento_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
   RaiseEvent GrdPagamentoRequestEdit(lRow, lCol, iKeyAscii, bCancel, sText, lMaxLength, eTextEditOpt)
End Sub
Private Sub GrdPagamento_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)
   Dim i As Integer
   
   With Me.GrdPagamento
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
Private Sub GrdPagamento_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)
   If (lRow = Me.GrdPagamento.RowCount) Then bRequestEdit = False
End Sub
Private Sub GrdPagamento_Validate(Cancel As Boolean)
   RaiseEvent GrdPagamentoValidate(Cancel)
End Sub
Private Sub GrdVenda_AfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
   RaiseEvent GrdVendaAfterCommitEdit(lRow, lCol)
End Sub
Private Sub GrdVenda_BeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
   RaiseEvent GrdVendaBeforeCommitEdit(lRow, lCol, eResult, sNewText, vNewValue, lConvErr)
End Sub
Private Sub GrdVenda_ColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
   RaiseEvent GrdVendaColHeaderClick(lCol, bDoDefault, Shift, x, y)
End Sub
Private Sub GrdVenda_ColHeaderDblClick(ByVal lCol As Long)
   RaiseEvent GrdVendaColHeaderDblClick(lCol)
End Sub
Private Sub GrdVenda_LostFocus()
   RaiseEvent GrdVendaLostFocus
End Sub
Private Sub GrdVenda_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
  RaiseEvent GrdVendaMouseUp(Button, Shift, x, y, lRow, lCol, bDoDefault)
End Sub
Private Sub GrdVenda_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
   RaiseEvent GrdVendaRequestEdit(lRow, lCol, iKeyAscii, bCancel, sText, lMaxLength, eTextEditOpt)
End Sub
Private Sub GrdVenda_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)
   Dim i As Integer
   With Me.GrdVenda
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
Private Sub GrdVenda_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)
   If (lRow = Me.GrdVenda.RowCount) Then bRequestEdit = False
End Sub
Private Sub GrdVenda_Validate(Cancel As Boolean)
   RaiseEvent GrdVendaValidate(Cancel)
End Sub

Private Sub PushButton1_Click()

End Sub

Private Sub TxtIDVENDA_GotFocus()
   RaiseEvent TxtIDVENDAGotFocus
End Sub
Private Sub TxtIDVENDA_LostFocus()
   RaiseEvent TxtIDVENDALostFocus
End Sub
Private Sub TxtNOME_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtTEL1_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtVLDESC_Change()
   RaiseEvent TxtVLDESCChange
End Sub
Private Sub TxtVLDESC_GotFocus()
   If xVal(Me.TxtVLDESC.Text) = 0 Then Me.TxtVLDESC.Text = ""
   'Call SelecionarTexto(Me.ActiveControl)
   Me.TxtVLDESC.SelStart = 0
   Me.TxtVLDESC.SelLength = Len(Me.TxtVLDESC.Text)
   
End Sub
Private Sub TxtVLDESC_LostFocus()
   Me.TxtVLDESC.Text = ValBr(Me.TxtVLDESC.Text)
End Sub
Private Sub TxtVLPGTO_Change()
   'RaiseEvent TxtVLPGTOChange
End Sub
Private Sub TxtVLPGTO_GotFocus()
  Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtVLPGTO_LostFocus()
   Me.TxtVLPGTO.Text = ValBr(Me.TxtVLPGTO.Text)
End Sub
Private Sub TxtVLTROCO_GotFocus()
  Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtVLVENDA_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub

Private Sub TxtVLVENDA_LostFocus()
   Me.TxtVLVENDA.Text = ValBr(Me.TxtVLVENDA.Text)
End Sub
