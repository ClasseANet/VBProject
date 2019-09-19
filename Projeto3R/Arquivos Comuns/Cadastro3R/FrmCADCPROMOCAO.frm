VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.Controls.v11.2.2.ocx"
Object = "{3F666E72-7F79-447A-BCFF-C42C44E3DBE7}#1.0#0"; "MCIGrid.ocx"
Begin VB.Form FrmCADCPROMOCAO 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Promoções / Descontos"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6660
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl TabControl 
      Height          =   4215
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   6495
      _Version        =   720898
      _ExtentX        =   11456
      _ExtentY        =   7435
      _StockProps     =   68
      Color           =   8
      PaintManager.Layout=   2
      PaintManager.BoldSelected=   -1  'True
      ItemCount       =   2
      Item(0).Caption =   "Produtos"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "TabControlPage1"
      Item(1).Caption =   "Serviços"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "TabControlPage2"
      Begin XtremeSuiteControls.TabControlPage TabControlPage2 
         Height          =   3825
         Left            =   -69970
         TabIndex        =   44
         Top             =   330
         Visible         =   0   'False
         Width           =   6405
         _Version        =   720898
         _ExtentX        =   11298
         _ExtentY        =   6747
         _StockProps     =   1
         Page            =   1
         Begin XtremeSuiteControls.GroupBox FrmeTrat 
            Height          =   2895
            Left            =   120
            TabIndex        =   29
            Top             =   840
            Width           =   6135
            _Version        =   720898
            _ExtentX        =   10821
            _ExtentY        =   5106
            _StockProps     =   79
            Caption         =   "Serviços, tratamentos e áreas permitidas"
            UseVisualStyle  =   -1  'True
            Appearance      =   4
            Begin XtremeSuiteControls.ListView LstAREA 
               Height          =   2055
               Left            =   3600
               TabIndex        =   37
               Top             =   480
               Width           =   2415
               _Version        =   720898
               _ExtentX        =   4260
               _ExtentY        =   3625
               _StockProps     =   77
               BackColor       =   -2147483643
               Checkboxes      =   -1  'True
               HideSelection   =   0   'False
               View            =   2
               FullRowSelect   =   -1  'True
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.ListView LstTPTRATAMENTO 
               Height          =   2055
               Left            =   1560
               TabIndex        =   34
               Top             =   480
               Width           =   1935
               _Version        =   720898
               _ExtentX        =   3413
               _ExtentY        =   3625
               _StockProps     =   77
               BackColor       =   -2147483643
               Checkboxes      =   -1  'True
               HideSelection   =   0   'False
               View            =   2
               FullRowSelect   =   -1  'True
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.ListView LstTPSERVICO 
               Height          =   2055
               Left            =   120
               TabIndex        =   31
               Top             =   480
               Width           =   1335
               _Version        =   720898
               _ExtentX        =   2355
               _ExtentY        =   3625
               _StockProps     =   77
               BackColor       =   -2147483643
               Checkboxes      =   -1  'True
               HideSelection   =   0   'False
               View            =   2
               FullRowSelect   =   -1  'True
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.PushButton CmdTodos 
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   32
               Top             =   2520
               Width           =   615
               _Version        =   720898
               _ExtentX        =   1085
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Todos"
               ForeColor       =   4210752
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.PushButton CmdTodos 
               Height          =   255
               Index           =   1
               Left            =   1560
               TabIndex        =   35
               Top             =   2520
               Width           =   615
               _Version        =   720898
               _ExtentX        =   1085
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Todos"
               ForeColor       =   4210752
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.PushButton CmdTodos 
               Height          =   255
               Index           =   2
               Left            =   3600
               TabIndex        =   38
               Top             =   2520
               Width           =   615
               _Version        =   720898
               _ExtentX        =   1085
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Todos"
               ForeColor       =   4210752
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.Label Label10 
               Height          =   285
               Left            =   3600
               TabIndex        =   36
               Top             =   240
               Width           =   1215
               _Version        =   720898
               _ExtentX        =   2143
               _ExtentY        =   503
               _StockProps     =   79
               Caption         =   "Áreas :"
               Transparent     =   -1  'True
            End
            Begin XtremeSuiteControls.Label Label9 
               Height          =   285
               Left            =   1560
               TabIndex        =   33
               Top             =   240
               Width           =   1215
               _Version        =   720898
               _ExtentX        =   2143
               _ExtentY        =   503
               _StockProps     =   79
               Caption         =   "Tratamentos :"
               Transparent     =   -1  'True
            End
            Begin XtremeSuiteControls.Label Label8 
               Height          =   285
               Left            =   120
               TabIndex        =   30
               Top             =   240
               Width           =   1215
               _Version        =   720898
               _ExtentX        =   2143
               _ExtentY        =   503
               _StockProps     =   79
               Caption         =   "Serviços :"
               Transparent     =   -1  'True
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox1 
            Height          =   735
            Left            =   120
            TabIndex        =   25
            Top             =   60
            Width           =   6135
            _Version        =   720898
            _ExtentX        =   10821
            _ExtentY        =   1296
            _StockProps     =   79
            Caption         =   " Após compra, cliente deve manter o cumpom no(a) mesmo(a)..."
            UseVisualStyle  =   -1  'True
            Appearance      =   4
            Begin XtremeSuiteControls.CheckBox ChkFlgServ 
               Height          =   255
               Left            =   360
               TabIndex        =   26
               Top             =   240
               Width           =   975
               _Version        =   720898
               _ExtentX        =   1720
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Serviço"
               Appearance      =   2
               MultiLine       =   0   'False
            End
            Begin XtremeSuiteControls.CheckBox ChkFlgTrat 
               Height          =   255
               Left            =   1920
               TabIndex        =   27
               Top             =   240
               Width           =   1335
               _Version        =   720898
               _ExtentX        =   2355
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Tratamento"
               Appearance      =   2
               MultiLine       =   0   'False
            End
            Begin XtremeSuiteControls.CheckBox chkFlgArea 
               Height          =   255
               Left            =   3960
               TabIndex        =   28
               Top             =   240
               Width           =   1335
               _Version        =   720898
               _ExtentX        =   2355
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Área"
               Appearance      =   2
               MultiLine       =   0   'False
            End
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage1 
         Height          =   3825
         Left            =   30
         TabIndex        =   43
         Top             =   360
         Width           =   6405
         _Version        =   720898
         _ExtentX        =   11298
         _ExtentY        =   6747
         _StockProps     =   1
         Page            =   0
         Begin XtremeSuiteControls.GroupBox FrmeProd 
            Height          =   2895
            Left            =   0
            TabIndex        =   14
            Top             =   480
            Width           =   6375
            _Version        =   720898
            _ExtentX        =   11245
            _ExtentY        =   5106
            _StockProps     =   79
            Caption         =   "Produtos de Venda"
            UseVisualStyle  =   -1  'True
            Appearance      =   4
            Begin XtremeSuiteControls.FlatEdit TxtVLDESC 
               Height          =   285
               Left            =   4680
               TabIndex        =   21
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
               MaxLength       =   6
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.FlatEdit TxtVLTOTAL 
               Height          =   285
               Left            =   4680
               TabIndex        =   19
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
               MaxLength       =   6
               UseVisualStyle  =   -1  'True
            End
            Begin iGrid251_75B4A91C.iGrid GrdProd 
               Height          =   1575
               Left            =   60
               TabIndex        =   15
               Top             =   240
               Width           =   6255
               _ExtentX        =   11033
               _ExtentY        =   2778
            End
            Begin XtremeSuiteControls.FlatEdit TxtVALOR 
               Height          =   285
               Left            =   4680
               TabIndex        =   24
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
               MaxLength       =   6
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.ComboBox CmbNPARCELA 
               Height          =   315
               Left            =   1200
               TabIndex        =   17
               Top             =   2115
               Width           =   1455
               _Version        =   720898
               _ExtentX        =   2566
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
               Style           =   2
               UseVisualStyle  =   -1  'True
               Text            =   "ComboBox1"
            End
            Begin XtremeSuiteControls.Label Label13 
               Height          =   285
               Left            =   120
               TabIndex        =   16
               Top             =   2115
               Width           =   1095
               _Version        =   720898
               _ExtentX        =   1931
               _ExtentY        =   503
               _StockProps     =   79
               Caption         =   "Nº Parcelas :"
               Transparent     =   -1  'True
            End
            Begin XtremeSuiteControls.Label Label7 
               Height          =   285
               Left            =   3480
               TabIndex        =   23
               Top             =   2385
               Width           =   1170
               _Version        =   720898
               _ExtentX        =   2064
               _ExtentY        =   503
               _StockProps     =   79
               Caption         =   "Valor a Pagar:"
               Transparent     =   -1  'True
            End
            Begin XtremeSuiteControls.Label Label3 
               Height          =   285
               Left            =   3480
               TabIndex        =   18
               Top             =   1845
               Width           =   1335
               _Version        =   720898
               _ExtentX        =   2355
               _ExtentY        =   503
               _StockProps     =   79
               Caption         =   "Valor Produtos :"
               Transparent     =   -1  'True
            End
            Begin XtremeSuiteControls.Label Label4 
               Height          =   285
               Left            =   3120
               TabIndex        =   20
               Top             =   2115
               Width           =   1500
               _Version        =   720898
               _ExtentX        =   2646
               _ExtentY        =   503
               _StockProps     =   79
               Caption         =   "Valor Desconto:"
               Alignment       =   1
               Transparent     =   -1  'True
            End
            Begin XtremeSuiteControls.Label LblDesc 
               Height          =   285
               Left            =   5880
               TabIndex        =   22
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
      End
   End
   Begin XtremeSuiteControls.FlatEdit TxtID 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   705
      _Version        =   720898
      _ExtentX        =   1244
      _ExtentY        =   503
      _StockProps     =   77
      BackColor       =   -2147483643
      MaxLength       =   20
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit TxtDSCPRO 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   480
      Width           =   4665
      _Version        =   720898
      _ExtentX        =   8229
      _ExtentY        =   503
      _StockProps     =   77
      BackColor       =   -2147483643
      MaxLength       =   30
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdSair 
      Height          =   375
      Left            =   4920
      TabIndex        =   42
      Top             =   6000
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Sai&r"
      ForeColor       =   0
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdExcluir 
      Height          =   375
      Left            =   3240
      TabIndex        =   41
      Top             =   6000
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Excluir"
      ForeColor       =   192
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdSalvar 
      Height          =   375
      Left            =   240
      TabIndex        =   39
      Top             =   6000
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Salvar"
      ForeColor       =   12582912
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdNovo 
      Height          =   375
      Left            =   1560
      TabIndex        =   40
      Top             =   6000
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Novo"
      ForeColor       =   4210752
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.CheckBox ChkATIVO 
      Height          =   255
      Left            =   4800
      TabIndex        =   2
      Top             =   120
      Width           =   975
      _Version        =   720898
      _ExtentX        =   1720
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Ativo "
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      Appearance      =   2
      RightToLeft     =   -1  'True
      MultiLine       =   0   'False
      Value           =   1
   End
   Begin XtremeSuiteControls.FlatEdit TxtDTINIV 
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   840
      Width           =   1185
      _Version        =   720898
      _ExtentX        =   2090
      _ExtentY        =   503
      _StockProps     =   77
      BackColor       =   -2147483643
      MaxLength       =   10
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit TxtDTINI 
      Height          =   285
      Left            =   1440
      TabIndex        =   10
      Top             =   1200
      Width           =   1185
      _Version        =   720898
      _ExtentX        =   2090
      _ExtentY        =   503
      _StockProps     =   77
      BackColor       =   -2147483643
      MaxLength       =   10
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit TxtDTFIMV 
      Height          =   285
      Left            =   4080
      TabIndex        =   8
      Top             =   840
      Width           =   1185
      _Version        =   720898
      _ExtentX        =   2090
      _ExtentY        =   503
      _StockProps     =   77
      BackColor       =   -2147483643
      MaxLength       =   10
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit TxtDTFIM 
      Height          =   285
      Left            =   4080
      TabIndex        =   12
      Top             =   1200
      Width           =   1185
      _Version        =   720898
      _ExtentX        =   2090
      _ExtentY        =   503
      _StockProps     =   77
      BackColor       =   -2147483643
      MaxLength       =   10
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   855
      _Version        =   720898
      _ExtentX        =   1508
      _ExtentY        =   503
      _StockProps     =   79
      Caption         =   "Id.:"
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label12 
      Height          =   195
      Left            =   3000
      TabIndex        =   11
      Top             =   1200
      Width           =   990
      _Version        =   720898
      _ExtentX        =   1746
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Fim Consumo:"
      Transparent     =   -1  'True
      AutoSize        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label11 
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   1155
      _Version        =   720898
      _ExtentX        =   2037
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Início Consumo:"
      Transparent     =   -1  'True
      AutoSize        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label6 
      Height          =   195
      Left            =   3000
      TabIndex        =   7
      Top             =   840
      Width           =   870
      _Version        =   720898
      _ExtentX        =   1535
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Fim Vendas:"
      Transparent     =   -1  'True
      AutoSize        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label5 
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   1035
      _Version        =   720898
      _ExtentX        =   1826
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Início Vendas:"
      Transparent     =   -1  'True
      AutoSize        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   855
      _Version        =   720898
      _ExtentX        =   1508
      _ExtentY        =   503
      _StockProps     =   79
      Caption         =   "Descrição :"
      Transparent     =   -1  'True
   End
End
Attribute VB_Name = "FrmCADCPROMOCAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Activate()
Event Load()
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event CmdSalvarClick()
Event CmdSairClick()
Event CmdNovoClick()
Event CmdExcluirClick()
Event CmdTodosClick(Index As Integer)
Event TxtIDLostFocus()
Event TxtVALORChange()
Event TxtVLDESCChange()
Event TxtVLTOTALChange()
Event GrdProdAfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
Event GrdProdBeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
Event GrdProdColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
Event GrdProdColHeaderDblClick(ByVal lCol As Long)
Event GrdProdLostFocus()
Event GrdProdMouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
Event GrdProdRequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
Event GrdProdValidate(Cancel As Boolean)
Private Sub CmdExcluir_Click()
   RaiseEvent CmdExcluirClick
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
Private Sub CmdTodos_Click(Index As Integer)
   RaiseEvent CmdTodosClick(Index)
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
Private Sub GrdProd_AfterCommitEdit(ByVal lRow As Long, ByVal lCol As Long)
   RaiseEvent GrdProdAfterCommitEdit(lRow, lCol)
End Sub
Private Sub GrdProd_BeforeCommitEdit(ByVal lRow As Long, ByVal lCol As Long, eResult As iGrid251_75B4A91C.EEditResults, ByVal sNewText As String, vNewValue As Variant, ByVal lConvErr As Long)
   RaiseEvent GrdProdBeforeCommitEdit(lRow, lCol, eResult, sNewText, vNewValue, lConvErr)
End Sub
Private Sub GrdProd_ColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)
   RaiseEvent GrdProdColHeaderClick(lCol, bDoDefault, Shift, x, y)
End Sub
Private Sub GrdProd_ColHeaderDblClick(ByVal lCol As Long)
   RaiseEvent GrdProdColHeaderDblClick(lCol)
End Sub
Private Sub GrdProd_LostFocus()
   RaiseEvent GrdProdLostFocus
End Sub
Private Sub GrdProd_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, ByVal lRow As Long, ByVal lCol As Long, bDoDefault As Boolean)
  RaiseEvent GrdProdMouseUp(Button, Shift, x, y, lRow, lCol, bDoDefault)
End Sub
Private Sub GrdProd_CurCellChange(ByVal lRow As Long, ByVal lCol As Long)
   Dim i As Integer
   With Me.GrdProd
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
Private Sub GrdProd_DblClick(ByVal lRow As Long, ByVal lCol As Long, bRequestEdit As Boolean)
   If (lRow = Me.GrdProd.RowCount) Then bRequestEdit = False
End Sub
Private Sub GrdProd_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean, sText As String, lMaxLength As Long, eTextEditOpt As iGrid251_75B4A91C.ETextEditFlags)
   RaiseEvent GrdProdRequestEdit(lRow, lCol, iKeyAscii, bCancel, sText, lMaxLength, eTextEditOpt)
End Sub
Private Sub GrdProd_Validate(Cancel As Boolean)
   RaiseEvent GrdProdValidate(Cancel)
End Sub
Private Sub TxtDSCPRO_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub

Private Sub TxtDTFIM_LostFocus()
   Me.TxtDTFIM.Text = FormatarData(Me.TxtDTFIM.Text)
End Sub
Private Sub TxtDTFIMV_LostFocus()
   Me.TxtDTFIMV.Text = FormatarData(Me.TxtDTFIMV.Text)
End Sub
Private Sub TxtDTINI_LostFocus()
   Me.TxtDTINI.Text = FormatarData(Me.TxtDTINI.Text)
End Sub
Private Sub TxtDTINIV_LostFocus()
   Me.TxtDTINIV.Text = FormatarData(Me.TxtDTINIV.Text)
End Sub
Private Sub TxtID_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtID_LostFocus()
   RaiseEvent TxtIDLostFocus
End Sub
Private Sub TxtVALOR_Change()
   RaiseEvent TxtVALORChange
End Sub
Private Sub TxtVALOR_GotFocus()
   With Me.TxtVALOR
      If xVal(.Text) = 0 Then .Text = ""
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub
Private Sub TxtVALOR_KeyPress(KeyAscii As Integer)
   If Not InArray(KeyAscii, Array(8, 44)) Then
      If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
   End If
End Sub
Private Sub TxtVALOR_LostFocus()
   Me.TxtVALOR.Text = ValBr(xVal(Me.TxtVALOR.Text, 2))
End Sub
Private Sub TxtVLDESC_Change()
   RaiseEvent TxtVLDESCChange
End Sub
Private Sub TxtVLDESC_GotFocus()
   With Me.TxtVLDESC
      If xVal(.Text) = 0 Then .Text = ""
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub
Private Sub TxtVLDESC_KeyPress(KeyAscii As Integer)
   If Not InArray(KeyAscii, Array(8, 44)) Then
      If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
   End If
End Sub
Private Sub TxtVLDESC_LostFocus()
   Me.TxtVLDESC.Text = ValBr(xVal(Me.TxtVLDESC.Text, 2))
End Sub
Private Sub TxtVLTOTAL_Change()
   RaiseEvent TxtVLTOTALChange
End Sub
