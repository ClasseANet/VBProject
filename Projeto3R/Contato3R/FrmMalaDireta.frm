VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#11.2#0"; "Codejock.ReportControl.v11.2.2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.Controls.v11.2.2.ocx"
Begin VB.Form FrmMalaDireta 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " MalaDireta"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8655
      _Version        =   720898
      _ExtentX        =   15266
      _ExtentY        =   8493
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   9
      Color           =   4
      PaintManager.BoldSelected=   -1  'True
      PaintManager.FixedTabWidth=   90
      ItemCount       =   2
      SelectedItem    =   1
      Item(0).Caption =   "Mensagem Html"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "TabControlPage1"
      Item(1).Caption =   "Lista de Clientes"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "TabControlPage2"
      Begin XtremeSuiteControls.TabControlPage TabControlPage2 
         Height          =   4440
         Left            =   30
         TabIndex        =   2
         Top             =   345
         Width           =   8595
         _Version        =   720898
         _ExtentX        =   15161
         _ExtentY        =   7832
         _StockProps     =   1
         Page            =   1
         Begin XtremeReportControl.ReportControl GrdContato 
            Height          =   3015
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Width           =   8295
            _Version        =   720898
            _ExtentX        =   14631
            _ExtentY        =   5318
            _StockProps     =   64
         End
         Begin XtremeSuiteControls.FlatEdit txtGrupo 
            Height          =   330
            Left            =   4920
            TabIndex        =   23
            Top             =   3960
            Width           =   615
            _Version        =   720898
            _ExtentX        =   1085
            _ExtentY        =   582
            _StockProps     =   77
            BackColor       =   -2147483643
            Enabled         =   0   'False
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton CmdEnviar 
            Height          =   375
            Left            =   5640
            TabIndex        =   13
            Top             =   3960
            Width           =   1095
            _Version        =   720898
            _ExtentX        =   1931
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Enviar"
            ForeColor       =   32768
            UseVisualStyle  =   -1  'True
            Picture         =   "FrmMalaDireta.frx":0000
         End
         Begin XtremeSuiteControls.PushButton CmdCarregar 
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Top             =   60
            Width           =   1095
            _Version        =   720898
            _ExtentX        =   1931
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Carregar"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton CmdTodos 
            Height          =   255
            Left            =   1200
            TabIndex        =   9
            Top             =   3960
            Width           =   975
            _Version        =   720898
            _ExtentX        =   1720
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Todos"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton CmdNenhum 
            Height          =   255
            Left            =   2160
            TabIndex        =   10
            Top             =   3960
            Width           =   975
            _Version        =   720898
            _ExtentX        =   1720
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Nenhum"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton CmdInverter 
            Height          =   255
            Left            =   3120
            TabIndex        =   11
            Top             =   3960
            Width           =   975
            _Version        =   720898
            _ExtentX        =   1720
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Inverter"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton CmdSair 
            Height          =   375
            Left            =   7440
            TabIndex        =   14
            Top             =   3960
            Width           =   975
            _Version        =   720898
            _ExtentX        =   1720
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Sair"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton CmdSelecionar 
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   3960
            Width           =   975
            _Version        =   720898
            _ExtentX        =   1720
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Selecionar"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton CmdImportar 
            Height          =   375
            Left            =   1320
            TabIndex        =   18
            Top             =   60
            Width           =   1095
            _Version        =   720898
            _ExtentX        =   1931
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Importar"
            Enabled         =   0   'False
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkEmEspera 
            Height          =   255
            Left            =   4800
            TabIndex        =   19
            Top             =   120
            Width           =   1095
            _Version        =   720898
            _ExtentX        =   1931
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Em Espera"
            BackColor       =   -2147483643
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   5
            Value           =   1
         End
         Begin XtremeSuiteControls.CheckBox ChkInativo 
            Height          =   255
            Left            =   3840
            TabIndex        =   20
            Top             =   120
            Width           =   855
            _Version        =   720898
            _ExtentX        =   1508
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Inativos"
            BackColor       =   -2147483643
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   5
         End
         Begin XtremeSuiteControls.CheckBox ChkAtivo 
            Height          =   255
            Left            =   3000
            TabIndex        =   21
            Top             =   120
            Width           =   735
            _Version        =   720898
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Ativos"
            BackColor       =   -2147483643
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   5
            Value           =   1
         End
         Begin XtremeSuiteControls.FlatEdit txtFiltrar 
            Height          =   330
            Left            =   6000
            TabIndex        =   22
            Top             =   120
            Width           =   2130
            _Version        =   720898
            _ExtentX        =   3757
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   12632256
            BackColor       =   -2147483643
            Text            =   "Pesquisar Contatos"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   330
            Left            =   4440
            TabIndex        =   24
            Top             =   3960
            Width           =   510
            _Version        =   720898
            _ExtentX        =   900
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Grupo:"
            Transparent     =   -1  'True
         End
         Begin VB.Image imgLupa 
            Height          =   270
            Left            =   8160
            Picture         =   "FrmMalaDireta.frx":0452
            Stretch         =   -1  'True
            Top             =   140
            Width           =   255
         End
         Begin VB.Label LblStGrd 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Total: (0 Itens)          Selecionados: (0 Itens)"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   3600
            Width           =   8295
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage1 
         Height          =   4440
         Left            =   -69970
         TabIndex        =   1
         Top             =   345
         Visible         =   0   'False
         Width           =   8595
         _Version        =   720898
         _ExtentX        =   15161
         _ExtentY        =   7832
         _StockProps     =   1
         Page            =   0
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Interval        =   60000
            Left            =   4920
            Top             =   3960
         End
         Begin XtremeSuiteControls.PushButton CmdOpen 
            Height          =   315
            Left            =   7680
            TabIndex        =   5
            Top             =   3600
            Width           =   375
            _Version        =   720898
            _ExtentX        =   661
            _ExtentY        =   556
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "FrmMalaDireta.frx":08A4
            BorderGap       =   0
         End
         Begin XtremeSuiteControls.FlatEdit TxtHtml 
            Height          =   315
            Left            =   120
            TabIndex        =   4
            Top             =   3600
            Width           =   7455
            _Version        =   720898
            _ExtentX        =   13150
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.WebBrowser WebBrowser1 
            Height          =   2775
            Left            =   120
            TabIndex        =   3
            Top             =   720
            Width           =   8295
            _Version        =   720898
            _ExtentX        =   14631
            _ExtentY        =   4895
            _StockProps     =   173
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit TxtTitulo 
            Height          =   285
            Left            =   720
            TabIndex        =   15
            Top             =   345
            Width           =   7695
            _Version        =   720898
            _ExtentX        =   13573
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton CmdEdit 
            Height          =   315
            Left            =   8040
            TabIndex        =   6
            Top             =   3600
            Width           =   375
            _Version        =   720898
            _ExtentX        =   661
            _ExtentY        =   556
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "FrmMalaDireta.frx":095E
            BorderGap       =   0
         End
         Begin XtremeSuiteControls.CommonDialog CommonDialog1 
            Left            =   5520
            Top             =   4080
            _Version        =   720898
            _ExtentX        =   423
            _ExtentY        =   423
            _StockProps     =   4
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   555
            _Version        =   720898
            _ExtentX        =   979
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Titulo:"
         End
      End
   End
End
Attribute VB_Name = "FrmMalaDireta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Activate()
Event Load()
Event Resize()
Event Timer1Timer()
Event CmdOpen(sFile As String)
Event CmdEdit()
Event CmdEnviar()
Event CmdCarregarClick()
Event GrdContatoRowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
Event GrdContatoItemCheck(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
Event GrdContatoKeyDown(KeyCode As Integer, Shift As Integer)
Event txtFiltrarGotFocus()
Event TxtFiltrarKeyPress(KeyAscii)
Event txtFiltrarLostFocus()
Event TxtHtmlLostFocus()
Event CmdTodosClick()
Event CmdNenhumClick()
Event CmdImportarClick()
Event CmdInverterClick()
Event CmdSelecionarClick()
Event CmdSairClick()
Event ChkAtivo(Index As Integer)
Private Sub ChkAtivo_Click()
   RaiseEvent ChkAtivo(1)
End Sub
Private Sub ChkEmEspera_Click()
   RaiseEvent ChkAtivo(2)
End Sub
Private Sub ChkInativo_Click()
   RaiseEvent ChkAtivo(0)
End Sub
Private Sub CmdCarregar_Click()
   RaiseEvent CmdCarregarClick
End Sub

Private Sub CmdEdit_Click()
   RaiseEvent CmdEdit
End Sub

Private Sub CmdEnviar_Click()
   RaiseEvent CmdEnviar
End Sub

Private Sub CmdImportar_Click()
   RaiseEvent CmdImportarClick
End Sub

Private Sub CmdInverter_Click()
   RaiseEvent CmdInverterClick
End Sub
Private Sub CmdNenhum_Click()
   RaiseEvent CmdNenhumClick
End Sub
Private Sub CmdOpen_Click()
   RaiseEvent CmdOpen("")
End Sub
Private Sub CmdSair_Click()
   RaiseEvent CmdSairClick
End Sub
Private Sub CmdSelecionar_Click()
   RaiseEvent CmdSelecionarClick
End Sub
Private Sub CmdTodos_Click()
   RaiseEvent CmdTodosClick
End Sub
Private Sub Form_Activate()
   RaiseEvent Activate
End Sub
Private Sub Form_Load()
   RaiseEvent Load
End Sub
Private Sub Form_Resize()
   RaiseEvent Resize
End Sub
Private Sub GrdContato_ItemCheck(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
   RaiseEvent GrdContatoItemCheck(Row, Item)
End Sub

Private Sub GrdContato_KeyDown(KeyCode As Integer, Shift As Integer)
   RaiseEvent GrdContatoKeyDown(KeyCode, Shift)
End Sub

Private Sub GrdContato_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
   RaiseEvent GrdContatoRowDblClick(Row, Item)
End Sub
Private Sub Timer1_Timer()
  RaiseEvent Timer1Timer
End Sub
Private Sub txtFiltrar_GotFocus()
   RaiseEvent txtFiltrarGotFocus
End Sub
Private Sub txtFiltrar_KeyPress(KeyAscii As Integer)
   RaiseEvent TxtFiltrarKeyPress(KeyAscii)
End Sub
Private Sub txtFiltrar_LostFocus()
   RaiseEvent txtFiltrarLostFocus
End Sub

Private Sub TxtHtml_LostFocus()
   RaiseEvent TxtHtmlLostFocus
End Sub
