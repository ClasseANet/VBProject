VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.ShortcutBar.v11.2.2.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#11.2#0"; "Codejock.ReportControl.v11.2.2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.Controls.v11.2.2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmTarefas 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Conta"
   ClientHeight    =   6195
   ClientLeft      =   3645
   ClientTop       =   840
   ClientWidth     =   10200
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin XtremeReportControl.ReportControl GrdMovCC 
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   9975
      _Version        =   720898
      _ExtentX        =   17595
      _ExtentY        =   4895
      _StockProps     =   64
   End
   Begin XtremeSuiteControls.GroupBox GrpPg 
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Width           =   9975
      _Version        =   720898
      _ExtentX        =   17595
      _ExtentY        =   3413
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   1
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton CmdNovo 
         Height          =   300
         Left            =   2280
         TabIndex        =   30
         Top             =   90
         Width           =   885
         _Version        =   720898
         _ExtentX        =   1561
         _ExtentY        =   529
         _StockProps     =   79
         Caption         =   "Novo"
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   4
      End
      Begin XtremeSuiteControls.PushButton CmdSalvar 
         Default         =   -1  'True
         Height          =   300
         Left            =   4320
         TabIndex        =   31
         Top             =   90
         Width           =   850
         _Version        =   720898
         _ExtentX        =   1499
         _ExtentY        =   529
         _StockProps     =   79
         Caption         =   "Salvar"
         ForeColor       =   12582912
         Enabled         =   0   'False
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   4
      End
      Begin XtremeSuiteControls.PushButton CmdCancelar 
         Cancel          =   -1  'True
         Height          =   300
         Left            =   5880
         TabIndex        =   32
         Top             =   90
         Width           =   885
         _Version        =   720898
         _ExtentX        =   1561
         _ExtentY        =   529
         _StockProps     =   79
         Caption         =   "Cancelar"
         ForeColor       =   128
         Enabled         =   0   'False
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   4
      End
      Begin XtremeSuiteControls.GroupBox GrpValor 
         Height          =   1335
         Left            =   6840
         TabIndex        =   7
         Top             =   0
         Width           =   3135
         _Version        =   720898
         _ExtentX        =   5530
         _ExtentY        =   2355
         _StockProps     =   79
         BackColor       =   16777215
         UseVisualStyle  =   -1  'True
         Appearance      =   4
         BorderStyle     =   2
         Begin XtremeSuiteControls.FlatEdit TxtDTBAIXA 
            Height          =   315
            Left            =   1200
            TabIndex        =   14
            Top             =   480
            Width           =   1545
            _Version        =   720898
            _ExtentX        =   2725
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            MaxLength       =   10
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit TxtValor 
            Height          =   315
            Left            =   1200
            TabIndex        =   19
            Top             =   840
            Width           =   1785
            _Version        =   720898
            _ExtentX        =   3149
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            MaxLength       =   12
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.DateTimePicker CmbDTBAIXA 
            Height          =   345
            Left            =   1290
            TabIndex        =   15
            Top             =   480
            Width           =   1710
            _Version        =   720898
            _ExtentX        =   3016
            _ExtentY        =   609
            _StockProps     =   68
            Format          =   1
            CurrentDate     =   40479.6280671296
         End
         Begin XtremeSuiteControls.FlatEdit TxtNDOC 
            Height          =   315
            Left            =   1200
            TabIndex        =   10
            Top             =   120
            Width           =   1785
            _Version        =   720898
            _ExtentX        =   3149
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            MaxLength       =   20
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton CmdNum 
            Height          =   255
            Left            =   80
            TabIndex        =   9
            Top             =   120
            Visible         =   0   'False
            Width           =   735
            _Version        =   720898
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "&Número:"
            BackColor       =   16777215
            FlatStyle       =   -1  'True
            Appearance      =   2
            MultiLine       =   0   'False
            ImageAlignment  =   4
            BorderGap       =   0
            ImageGap        =   0
         End
         Begin XtremeSuiteControls.Label LblVALOR 
            Height          =   285
            Left            =   120
            TabIndex        =   18
            Top             =   840
            Width           =   600
            _Version        =   720898
            _ExtentX        =   1058
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Valor:"
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label LblDTBAIXA 
            Height          =   285
            Left            =   120
            TabIndex        =   13
            Top             =   480
            Width           =   600
            _Version        =   720898
            _ExtentX        =   1058
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Data:"
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label LblNum 
            Height          =   285
            Left            =   120
            TabIndex        =   8
            Top             =   120
            Width           =   600
            _Version        =   720898
            _ExtentX        =   1058
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Número:"
            Transparent     =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox GrpDe 
         Height          =   1815
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   5655
         _Version        =   720898
         _ExtentX        =   9975
         _ExtentY        =   3201
         _StockProps     =   79
         BackColor       =   16777215
         UseVisualStyle  =   -1  'True
         Appearance      =   4
         BorderStyle     =   2
         Begin XtremeSuiteControls.FlatEdit TxtObs 
            Height          =   315
            Left            =   1320
            TabIndex        =   24
            Top             =   1440
            Width           =   4305
            _Version        =   720898
            _ExtentX        =   7594
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            MaxLength       =   80
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.ComboBox CmbDe 
            Height          =   315
            Left            =   1320
            TabIndex        =   6
            Top             =   0
            Visible         =   0   'False
            Width           =   4275
            _Version        =   720898
            _ExtentX        =   7541
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.ComboBox CmbPara 
            Height          =   315
            Left            =   1320
            TabIndex        =   12
            Top             =   360
            Visible         =   0   'False
            Width           =   4275
            _Version        =   720898
            _ExtentX        =   7541
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.ComboBox CmbFavorecido 
            Height          =   315
            Left            =   1320
            TabIndex        =   17
            Top             =   720
            Width           =   4275
            _Version        =   720898
            _ExtentX        =   7541
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.ComboBox CmbCategoria 
            Height          =   315
            Left            =   1320
            TabIndex        =   21
            Top             =   1080
            Width           =   2115
            _Version        =   720898
            _ExtentX        =   3731
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.ComboBox CmbSubCategoria 
            Height          =   315
            Left            =   3480
            TabIndex        =   22
            Top             =   1080
            Width           =   2115
            _Version        =   720898
            _ExtentX        =   3731
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.Label LblObs 
            Height          =   285
            Left            =   0
            TabIndex        =   23
            Top             =   1440
            Width           =   960
            _Version        =   720898
            _ExtentX        =   1693
            _ExtentY        =   494
            _StockProps     =   79
            Caption         =   "Observação:"
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label LblCategoria 
            Height          =   285
            Left            =   0
            TabIndex        =   20
            Top             =   1080
            Width           =   960
            _Version        =   720898
            _ExtentX        =   1693
            _ExtentY        =   494
            _StockProps     =   79
            Caption         =   "Categoria:"
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label LblDe 
            Height          =   285
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Visible         =   0   'False
            Width           =   960
            _Version        =   720898
            _ExtentX        =   1693
            _ExtentY        =   494
            _StockProps     =   79
            Caption         =   "De:"
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label LblFavorecido 
            Height          =   285
            Left            =   0
            TabIndex        =   16
            Top             =   720
            Width           =   960
            _Version        =   720898
            _ExtentX        =   1693
            _ExtentY        =   494
            _StockProps     =   79
            Caption         =   "De:/Pagar a:"
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label LblPara 
            Height          =   285
            Left            =   0
            TabIndex        =   11
            Top             =   360
            Visible         =   0   'False
            Width           =   960
            _Version        =   720898
            _ExtentX        =   1693
            _ExtentY        =   494
            _StockProps     =   79
            Caption         =   "Para:"
            Transparent     =   -1  'True
         End
      End
      Begin XtremeSuiteControls.Label LblTitulo 
         Height          =   405
         Left            =   8280
         TabIndex        =   33
         Top             =   1440
         Width           =   1560
         _Version        =   720898
         _ExtentX        =   2752
         _ExtentY        =   714
         _StockProps     =   79
         Caption         =   "Valor:"
         ForeColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Transparent     =   -1  'True
      End
   End
   Begin XtremeSuiteControls.TabControl TabConta 
      Height          =   2415
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   9975
      _Version        =   720898
      _ExtentX        =   17595
      _ExtentY        =   4260
      _StockProps     =   68
      Appearance      =   3
      Color           =   4
      PaintManager.Layout=   2
      PaintManager.BoldSelected=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.FixedTabWidth=   90
      ItemCount       =   4
      Item(0).Caption =   "Depósito"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "TabPgDeposito"
      Item(1).Caption =   "Transferência"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "TabPgTransferencia"
      Item(2).Caption =   "Retirada"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "TabPgRetirada"
      Item(3).Caption =   "Cheque"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "TabPgCheque"
      Begin XtremeSuiteControls.TabControlPage TabPgCheque 
         Height          =   2040
         Left            =   -69970
         TabIndex        =   28
         Top             =   345
         Visible         =   0   'False
         Width           =   9315
         _Version        =   720898
         _ExtentX        =   16431
         _ExtentY        =   3598
         _StockProps     =   1
         Page            =   3
      End
      Begin XtremeSuiteControls.TabControlPage TabPgRetirada 
         Height          =   2040
         Left            =   -69970
         TabIndex        =   27
         Top             =   345
         Visible         =   0   'False
         Width           =   9315
         _Version        =   720898
         _ExtentX        =   16431
         _ExtentY        =   3598
         _StockProps     =   1
         Page            =   2
      End
      Begin XtremeSuiteControls.TabControlPage TabPgTransferencia 
         Height          =   2040
         Left            =   -69970
         TabIndex        =   26
         Top             =   345
         Visible         =   0   'False
         Width           =   9315
         _Version        =   720898
         _ExtentX        =   16431
         _ExtentY        =   3598
         _StockProps     =   1
         Page            =   1
      End
      Begin XtremeSuiteControls.TabControlPage TabPgDeposito 
         Height          =   2025
         Left            =   30
         TabIndex        =   29
         Top             =   360
         Width           =   9915
         _Version        =   720898
         _ExtentX        =   17489
         _ExtentY        =   3572
         _StockProps     =   1
         Page            =   0
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3000
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTarefas.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTarefas.frx":0452
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit txtFiltrar 
      Height          =   330
      Left            =   6720
      TabIndex        =   25
      Top             =   60
      Width           =   2130
      _Version        =   720898
      _ExtentX        =   3757
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   12632256
      BackColor       =   -2147483643
      Text            =   "Pesquisar Extrato"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Image imgLupa 
      Height          =   270
      Left            =   8880
      Picture         =   "FrmTarefas.frx":05AC
      Stretch         =   -1  'True
      Top             =   60
      Width           =   255
   End
   Begin XtremeShortcutBar.ShortcutCaption SccConta 
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10200
      _Version        =   720898
      _ExtentX        =   17992
      _ExtentY        =   741
      _StockProps     =   14
      Caption         =   "Conta: Caixa"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.76
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FrmTarefas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Activate()
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event Load()
Event Rezise()
Event Unload(Cancel As Integer)
Event CmdCancelarClick()
Event CmdSalvarClick()
Event CmdNovoClick()
Event CmdNumClick()
Event CmbDeLostFocus()
Event CmbParaLostFocus()
Event CmbFavorecidoLostFocus()
Event CmbCategoriaLostFocus()
Event CmbCategoriaChange()
Event CmbCategoriaClick()
Event CmbSubCategoriaLostFocus()
Event TxtDTBAIXALostFocus()
Event CmbDTBAIXALostFocus()
Event CmbDTBAIXAChange()
Event txtFiltrarGotFocus()
Event txtFiltrarLostFocus()
Event txtFiltrarKeyPress(KeyAscii As Integer)
Event TxtNDOCLostFocus()
Event GrdMovCCRowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
Event GrdMovCCBeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
Event GrdMovCCSelectionChanged()
Event TabContaBeforeItemClick(ByVal Item As XtremeSuiteControls.ITabControlItem, Cancel As Variant)
Event TabContaSelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Private Sub CmbCategoria_Change()
   RaiseEvent CmbCategoriaChange
End Sub
Private Sub CmbCategoria_Click()
   RaiseEvent CmbCategoriaClick
End Sub
Private Sub CmbCategoria_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub CmbCategoria_LostFocus()
   RaiseEvent CmbCategoriaLostFocus
End Sub
Private Sub CmbDe_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub CmbDe_LostFocus()
   RaiseEvent CmbDeLostFocus
End Sub
Private Sub CmbDTBAIXA_Change()
   RaiseEvent CmbDTBAIXAChange
End Sub
Private Sub CmbDTBAIXA_LostFocus()
   RaiseEvent CmbDTBAIXALostFocus
End Sub
Private Sub CmbFavorecido_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub CmbFavorecido_LostFocus()
   RaiseEvent CmbFavorecidoLostFocus
End Sub
Private Sub CmbPara_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub CmbPara_LostFocus()
   RaiseEvent CmbParaLostFocus
End Sub
Private Sub CmbSubCategoria_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub CmbSubCategoria_LostFocus()
   RaiseEvent CmbSubCategoriaLostFocus
End Sub
Private Sub CmdCancelar_Click()
   RaiseEvent CmdCancelarClick
End Sub
Private Sub CmdNovo_Click()
   RaiseEvent CmdNovoClick
End Sub
Private Sub CmdNum_Click()
   RaiseEvent CmdNumClick
End Sub
Private Sub CmdSalvar_Click()
   RaiseEvent CmdSalvarClick
End Sub
Private Sub Form_Activate()
   RaiseEvent Activate
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
   RaiseEvent KeyPress(KeyAscii)
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub Form_Load()
   Screen.MousePointer = vbHourglass
   RaiseEvent Load
End Sub
Private Sub Form_Resize()
   RaiseEvent Rezise
End Sub
Private Sub Form_Unload(Cancel As Integer)
   RaiseEvent Unload(Cancel)
End Sub
Private Sub GrdMovCC_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
   RaiseEvent GrdMovCCBeforeDrawRow(Row, Item, Metrics)
End Sub
Private Sub GrdMovCC_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub GrdMovCC_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
   RaiseEvent GrdMovCCRowDblClick(Row, Item)
End Sub
Private Sub GrdMovCC_SelectionChanged()
   RaiseEvent GrdMovCCSelectionChanged
End Sub
Private Sub TabConta_BeforeItemClick(ByVal Item As XtremeSuiteControls.ITabControlItem, Cancel As Variant)
   RaiseEvent TabContaBeforeItemClick(Item, Cancel)
End Sub
Private Sub TabConta_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
   RaiseEvent TabContaSelectedChanged(Item)
End Sub

Private Sub TxtDTBAIXA_GotFocus()
   If Me.TxtDTBAIXA.Enabled Then
      Me.TxtDTBAIXA.SelStart = 0
      Me.TxtDTBAIXA.SelLength = Len(Me.TxtDTBAIXA.Text)
      Call SelecionarTexto(Me.TxtDTBAIXA)
   End If
End Sub
Private Sub TxtDTBAIXA_LostFocus()
   RaiseEvent TxtDTBAIXALostFocus
End Sub
Private Sub txtFiltrar_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
   RaiseEvent txtFiltrarGotFocus
End Sub
Private Sub txtFiltrar_KeyPress(KeyAscii As Integer)
   RaiseEvent txtFiltrarKeyPress(KeyAscii)
End Sub
Private Sub txtFiltrar_LostFocus()
   RaiseEvent txtFiltrarLostFocus
End Sub
Private Sub TxtNDOC_Change()
   Me.CmdNum.Visible = (Trim(Me.TxtNDOC.Text) <> "")
End Sub
Private Sub TxtNDOC_GotFocus()
   If Me.TxtNDOC.Enabled Then
      Me.TxtNDOC.SelStart = 0
      Me.TxtNDOC.SelLength = Len(Me.TxtNDOC.Text)
      Call SelecionarTexto(Me.TxtNDOC)
   End If
End Sub
Private Sub TxtNDOC_LostFocus()
   RaiseEvent TxtNDOCLostFocus
End Sub
Private Sub TxtObs_GotFocus()
   If Me.TxtObs.Enabled Then
      Me.ActiveControl.SelStart = 0
      Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
      Call SelecionarTexto(Me.ActiveControl)
   End If
End Sub
Private Sub TxtValor_GotFocus()
   If Me.TxtValor.Enabled Then
      Me.TxtValor.SelStart = 0
      Me.TxtValor.SelLength = Len(Me.TxtValor.Text)
      Call SelecionarTexto(Me.TxtValor)
   End If
End Sub
Private Sub TxtValor_LostFocus()
   Me.TxtValor.Text = ValBr(Me.TxtValor.Text)
End Sub
