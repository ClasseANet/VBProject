VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.ShortcutBar.v11.2.2.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#11.2#0"; "Codejock.ReportControl.v11.2.2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "COD063~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#11.2#0"; "Codejock.CommandBars.v11.2.2.ocx"
Begin VB.Form FrmMovEst 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Movimentação de Estoque"
   ClientHeight    =   6195
   ClientLeft      =   3585
   ClientTop       =   2475
   ClientWidth     =   10980
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   10980
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin XtremeReportControl.ReportControl GrdMov 
      Height          =   3015
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   10695
      _Version        =   720898
      _ExtentX        =   18865
      _ExtentY        =   5318
      _StockProps     =   64
   End
   Begin XtremeSuiteControls.ComboBox CmbProduto 
      Height          =   315
      Left            =   6000
      TabIndex        =   27
      Top             =   60
      Visible         =   0   'False
      Width           =   4875
      _Version        =   720898
      _ExtentX        =   8599
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Appearance      =   6
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GrpPg 
      Height          =   1815
      Left            =   240
      TabIndex        =   18
      Top             =   4080
      Width           =   10455
      _Version        =   720898
      _ExtentX        =   18441
      _ExtentY        =   3201
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   1
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton CmdNovo 
         Height          =   300
         Left            =   2280
         TabIndex        =   3
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
         TabIndex        =   4
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
         TabIndex        =   5
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
         Height          =   1455
         Left            =   7320
         TabIndex        =   20
         Top             =   0
         Width           =   3135
         _Version        =   720898
         _ExtentX        =   5530
         _ExtentY        =   2566
         _StockProps     =   79
         BackColor       =   16777215
         UseVisualStyle  =   -1  'True
         Appearance      =   4
         BorderStyle     =   2
         Begin XtremeSuiteControls.FlatEdit TxtDTMOV 
            Height          =   315
            Left            =   1200
            TabIndex        =   12
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
            TabIndex        =   16
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
         Begin XtremeSuiteControls.DateTimePicker CmbDTMOV 
            Height          =   345
            Left            =   1290
            TabIndex        =   17
            Top             =   480
            Width           =   1710
            _Version        =   720898
            _ExtentX        =   3016
            _ExtentY        =   609
            _StockProps     =   68
            Format          =   1
            CurrentDate     =   40479.6280671296
         End
         Begin XtremeSuiteControls.FlatEdit TxtNUMDOC 
            Height          =   315
            Left            =   1200
            TabIndex        =   8
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
            TabIndex        =   7
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
            TabIndex        =   15
            Top             =   840
            Width           =   1080
            _Version        =   720898
            _ExtentX        =   1905
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Quantidade:"
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label LblDTMOV 
            Height          =   285
            Left            =   120
            TabIndex        =   11
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
            TabIndex        =   6
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
         TabIndex        =   19
         Top             =   120
         Width           =   6135
         _Version        =   720898
         _ExtentX        =   10821
         _ExtentY        =   3201
         _StockProps     =   79
         BackColor       =   16777215
         UseVisualStyle  =   -1  'True
         Appearance      =   4
         BorderStyle     =   2
         Begin XtremeSuiteControls.FlatEdit TxtObs 
            Height          =   315
            Left            =   1440
            TabIndex        =   14
            Top             =   840
            Width           =   4305
            _Version        =   720898
            _ExtentX        =   7594
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            MaxLength       =   80
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.ComboBox CmbFavorecido 
            Height          =   315
            Left            =   1440
            TabIndex        =   10
            Top             =   480
            Width           =   4275
            _Version        =   720898
            _ExtentX        =   7541
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.Label LblObs 
            Height          =   285
            Left            =   0
            TabIndex        =   13
            Top             =   840
            Width           =   960
            _Version        =   720898
            _ExtentX        =   1693
            _ExtentY        =   494
            _StockProps     =   79
            Caption         =   "Observação:"
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label LblFavorecido 
            Height          =   285
            Left            =   0
            TabIndex        =   9
            Top             =   480
            Width           =   960
            _Version        =   720898
            _ExtentX        =   1693
            _ExtentY        =   494
            _StockProps     =   79
            Caption         =   "De:/Pagar a:"
            Transparent     =   -1  'True
         End
      End
      Begin XtremeSuiteControls.Label LblTitulo 
         Height          =   405
         Left            =   8760
         TabIndex        =   26
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
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   10695
      _Version        =   720898
      _ExtentX        =   18865
      _ExtentY        =   4048
      _StockProps     =   68
      Appearance      =   3
      Color           =   4
      PaintManager.BoldSelected=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.FixedTabWidth=   90
      ItemCount       =   3
      Item(0).Caption =   "Entrada/Compra"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "TabPgDeposito"
      Item(0).Control(1)=   "TabPgCheque"
      Item(1).Caption =   "Transferência"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "TabPgTransferencia"
      Item(2).Caption =   "Saída/Venda"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "TabPgRetirada"
      Begin XtremeSuiteControls.TabControlPage TabPgCheque 
         Height          =   1905
         Left            =   30
         TabIndex        =   24
         Top             =   360
         Visible         =   0   'False
         Width           =   10635
         _Version        =   720898
         _ExtentX        =   18759
         _ExtentY        =   3360
         _StockProps     =   1
         Page            =   1
      End
      Begin XtremeSuiteControls.TabControlPage TabPgRetirada 
         Height          =   2025
         Left            =   -69970
         TabIndex        =   23
         Top             =   360
         Visible         =   0   'False
         Width           =   10635
         _Version        =   720898
         _ExtentX        =   18759
         _ExtentY        =   3572
         _StockProps     =   1
         Page            =   3
      End
      Begin XtremeSuiteControls.TabControlPage TabPgTransferencia 
         Height          =   2025
         Left            =   -69970
         TabIndex        =   22
         Top             =   360
         Visible         =   0   'False
         Width           =   10635
         _Version        =   720898
         _ExtentX        =   18759
         _ExtentY        =   3572
         _StockProps     =   1
         Page            =   2
      End
      Begin XtremeSuiteControls.TabControlPage TabPgDeposito 
         Height          =   1905
         Left            =   30
         TabIndex        =   25
         Top             =   360
         Width           =   10635
         _Version        =   720898
         _ExtentX        =   18759
         _ExtentY        =   3360
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
            Picture         =   "FrmMovEst.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMovEst.frx":0452
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit txtFiltrar 
      Height          =   330
      Left            =   8400
      TabIndex        =   21
      Top             =   60
      Visible         =   0   'False
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
   Begin VB.Image ImgMenu 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   0
      MouseIcon       =   "FrmMovEst.frx":05AC
      MousePointer    =   99  'Custom
      Picture         =   "FrmMovEst.frx":08B6
      Top             =   120
      Width           =   240
   End
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   0
      Top             =   360
      _Version        =   720898
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      ScaleMode       =   2
      VisualTheme     =   2
   End
   Begin VB.Image imgLupa 
      Height          =   270
      Left            =   10560
      Picture         =   "FrmMovEst.frx":0E40
      Stretch         =   -1  'True
      Top             =   60
      Visible         =   0   'False
      Width           =   255
   End
   Begin XtremeShortcutBar.ShortcutCaption SccTit 
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11040
      _Version        =   720898
      _ExtentX        =   19473
      _ExtentY        =   741
      _StockProps     =   14
      Caption         =   "  Loja1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FrmMovEst"
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
Event CmdNumClick()
Event CmdNovoClick()
Event CmbDeLostFocus()
Event CmbParaLostFocus()
Event CmbFavorecidoLostFocus()
Event CmbProdutoLostFocus()
Event CmbProdutoChange()
Event CmbProdutoClick()
Event CmbDTMOVLostFocus()
Event CmbDTMOVChange()
Event CommandBarsExecute(ByVal Control As XtremeCommandBars.ICommandBarControl)
Event GrdMovRowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
Event GrdMovBeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
Event GrdMovSelectionChanged()
Event ImgMenuMouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event SccTitMouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event TxtDTMOVLostFocus()
Event TxtFiltrarGotFocus()
Event TxtFiltrarLostFocus()
Event TxtFiltrarKeyPress(KeyAscii As Integer)
Event TxtNUMDOCLostFocus()
Event TabContaBeforeItemClick(ByVal Item As XtremeSuiteControls.ITabControlItem, Cancel As Variant)
Event TabContaSelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Private Sub CmbDe_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub CmbDe_LostFocus()
   RaiseEvent CmbDeLostFocus
End Sub
Private Sub CmbProduto_Change()
   RaiseEvent CmbProdutoChange
End Sub
Private Sub CmbProduto_Click()
   RaiseEvent CmbProdutoClick
End Sub
Private Sub CmbProduto_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub CmbProduto_LostFocus()
   RaiseEvent CmbProdutoLostFocus
End Sub
Private Sub CmbDTMOV_Change()
   RaiseEvent CmbDTMOVChange
End Sub
Private Sub CmbDTMOV_LostFocus()
   RaiseEvent CmbDTMOVLostFocus
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
Private Sub CommandBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
   RaiseEvent CommandBarsExecute(Control)
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
Private Sub GrdMov_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
   RaiseEvent GrdMovBeforeDrawRow(Row, Item, Metrics)
End Sub
Private Sub GrdMov_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub GrdMov_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
   RaiseEvent GrdMovRowDblClick(Row, Item)
End Sub
Private Sub GrdMov_SelectionChanged()
   RaiseEvent GrdMovSelectionChanged
End Sub
Private Sub ImgMenu_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   RaiseEvent ImgMenuMouseDown(Button, Shift, x, y)
End Sub
Private Sub SccTit_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   RaiseEvent SccTitMouseDown(Button, Shift, x, y)
End Sub
Private Sub TabConta_BeforeItemClick(ByVal Item As XtremeSuiteControls.ITabControlItem, Cancel As Variant)
   RaiseEvent TabContaBeforeItemClick(Item, Cancel)
End Sub
Private Sub TabConta_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
   RaiseEvent TabContaSelectedChanged(Item)
End Sub

Private Sub TxtDTMOV_GotFocus()
   If Me.TxtDTMOV.Enabled Then
      Me.TxtDTMOV.SelStart = 0
      Me.TxtDTMOV.SelLength = Len(Me.TxtDTMOV.Text)
      Call SelecionarTexto(Me.TxtDTMOV)
   End If
End Sub
Private Sub TxtDTMOV_LostFocus()
   RaiseEvent TxtDTMOVLostFocus
End Sub
Private Sub txtFiltrar_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
   RaiseEvent TxtFiltrarGotFocus
End Sub
Private Sub txtFiltrar_KeyPress(KeyAscii As Integer)
   RaiseEvent TxtFiltrarKeyPress(KeyAscii)
End Sub
Private Sub txtFiltrar_LostFocus()
   RaiseEvent TxtFiltrarLostFocus
End Sub
Private Sub TxtNUMDOC_Change()
   Me.CmdNum.Visible = (Trim(Me.TxtNUMDOC.Text) <> "")
End Sub
Private Sub TxtNUMDOC_GotFocus()
   If Me.TxtNUMDOC.Enabled Then
      Me.TxtNUMDOC.SelStart = 0
      Me.TxtNUMDOC.SelLength = Len(Me.TxtNUMDOC.Text)
      Call SelecionarTexto(Me.TxtNUMDOC)
   End If
End Sub
Private Sub TxtNUMDOC_LostFocus()
   RaiseEvent TxtNUMDOCLostFocus
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
