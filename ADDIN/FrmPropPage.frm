VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmPropPage 
   AutoRedraw      =   -1  'True
   Caption         =   "Propriedades"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   2115
   ClientWidth     =   5685
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5850
   ScaleWidth      =   5685
   Visible         =   0   'False
   Begin TabDlg.SSTab TabProp 
      Height          =   4815
      Left            =   240
      TabIndex        =   25
      Top             =   240
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   8493
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Propriedades"
      TabPicture(0)   =   "FrmPropPage.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Lbl(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Lbl(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "CmbBasedOn_DataType"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ChkTopLevel"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "TxtNome"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "FrmeColl"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "FrmeProp"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "FrmeMetoEve"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Atributos"
      TabPicture(1)   =   "FrmPropPage.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Lbl(2)"
      Tab(1).Control(1)=   "Lbl(3)"
      Tab(1).Control(2)=   "LblHelpFile"
      Tab(1).Control(3)=   "Lbl(4)"
      Tab(1).Control(4)=   "TxtDesc"
      Tab(1).Control(5)=   "TxtHelpID"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Criação do Objeto"
      TabPicture(2)   =   "FrmPropPage.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.Frame FrmeMetoEve 
         Height          =   3615
         Left            =   2280
         TabIndex        =   17
         Top             =   4440
         Width           =   4695
         Begin Threed.SSCommand CmdArg 
            Height          =   375
            Index           =   0
            Left            =   4200
            TabIndex        =   32
            Top             =   360
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            _StockProps     =   78
            Picture         =   "FrmPropPage.frx":0054
         End
         Begin VB.ListBox LstArg 
            Height          =   1815
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   3975
         End
         Begin Threed.SSCommand CmdArg 
            Height          =   375
            Index           =   1
            Left            =   4200
            TabIndex        =   33
            Top             =   840
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            _StockProps     =   78
            Picture         =   "FrmPropPage.frx":0636
         End
         Begin Threed.SSCommand CmdArg 
            Height          =   375
            Index           =   2
            Left            =   4200
            TabIndex        =   34
            Top             =   1320
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            _StockProps     =   78
            Picture         =   "FrmPropPage.frx":0C18
         End
         Begin Threed.SSCommand CmdArg 
            Height          =   375
            Index           =   3
            Left            =   4200
            TabIndex        =   35
            Top             =   1800
            Width           =   375
            _Version        =   65536
            _ExtentX        =   661
            _ExtentY        =   661
            _StockProps     =   78
            Picture         =   "FrmPropPage.frx":11FA
         End
         Begin VB.Frame FrmeMetodo 
            Height          =   1335
            Left            =   120
            TabIndex        =   21
            Top             =   2160
            Width           =   4455
            Begin VB.ComboBox CmbEscopo 
               Height          =   315
               ItemData        =   "FrmPropPage.frx":17DC
               Left            =   840
               List            =   "FrmPropPage.frx":17EC
               Style           =   2  'Dropdown List
               TabIndex        =   20
               Top             =   240
               Width           =   3435
            End
            Begin VB.ComboBox CmbReturn 
               Height          =   315
               ItemData        =   "FrmPropPage.frx":1813
               Left            =   1440
               List            =   "FrmPropPage.frx":1841
               Style           =   2  'Dropdown List
               TabIndex        =   10
               Top             =   720
               Width           =   2835
            End
            Begin VB.Label Lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo de Retorno :"
               Height          =   195
               Index           =   6
               Left            =   120
               TabIndex        =   22
               Top             =   840
               Width           =   1245
            End
            Begin VB.Label Lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Escopo :"
               Height          =   195
               Index           =   7
               Left            =   120
               TabIndex        =   36
               Top             =   360
               Width           =   630
            End
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Argumentos :"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   18
            Top             =   120
            Width           =   930
         End
      End
      Begin VB.Frame FrmeProp 
         Caption         =   "Declaração"
         Height          =   2775
         Left            =   240
         TabIndex        =   12
         Top             =   1680
         Width           =   4695
         Begin VB.TextBox TxtItem 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   2520
            MaxLength       =   50
            TabIndex        =   44
            Top             =   2040
            Width           =   3820
         End
         Begin MSFlexGridLib.MSFlexGrid GrdEnumType 
            Height          =   1815
            Left            =   2520
            TabIndex        =   43
            Top             =   1680
            Width           =   4185
            _ExtentX        =   7382
            _ExtentY        =   3201
            _Version        =   393216
            Rows            =   5
            Cols            =   1
            FixedCols       =   0
            RowHeightMin    =   330
            BackColor       =   12648447
            BackColorSel    =   12582912
            AllowBigSelection=   0   'False
            SelectionMode   =   1
         End
         Begin VB.TextBox TxtValorPadrao 
            Height          =   315
            Left            =   1560
            TabIndex        =   42
            Top             =   2400
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.OptionButton OptDeclaração 
            Caption         =   "Dim Variable"
            Height          =   255
            Index           =   6
            Left            =   3000
            TabIndex        =   41
            Top             =   960
            Width           =   1335
         End
         Begin VB.CheckBox ChkisConst 
            Caption         =   "Constante ?"
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   2400
            Width           =   1215
         End
         Begin VB.OptionButton OptDeclaração 
            Caption         =   "Static Variable"
            Height          =   255
            Index           =   5
            Left            =   3000
            TabIndex        =   39
            Top             =   600
            Width           =   1455
         End
         Begin VB.OptionButton OptDeclaração 
            Caption         =   "Private Variable"
            Height          =   255
            Index           =   4
            Left            =   3000
            TabIndex        =   38
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton OptDeclaração 
            Caption         =   "Global Variable"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   37
            Top             =   1320
            Width           =   1455
         End
         Begin VB.CheckBox ChkDefaultProp 
            Caption         =   "Propriedade Padrão ?"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   2160
            Width           =   1935
         End
         Begin VB.OptionButton OptDeclaração 
            Caption         =   "Public Variable"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   15
            Top             =   960
            Width           =   1575
         End
         Begin VB.OptionButton OptDeclaração 
            Caption         =   "Friend Public ( Let, Get, Set )"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   14
            Top             =   600
            Width           =   2415
         End
         Begin VB.OptionButton OptDeclaração 
            Caption         =   "Property Public ( Let, Get, Set )"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   13
            Top             =   240
            Value           =   -1  'True
            Width           =   2535
         End
      End
      Begin VB.Frame FrmeColl 
         Caption         =   "Coleção de "
         Height          =   2775
         Left            =   240
         TabIndex        =   5
         Top             =   1680
         Width           =   4695
         Begin VB.CommandButton CmdNewClasse 
            Caption         =   "Propriedades da Nova Classe..."
            Height          =   255
            Left            =   1560
            TabIndex        =   11
            Top             =   2400
            Width           =   2775
         End
         Begin VB.ListBox LstClasses 
            Height          =   1425
            Left            =   240
            TabIndex        =   7
            Top             =   480
            Width           =   4215
         End
         Begin VB.TextBox TxtClassName 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1560
            TabIndex        =   9
            Text            =   "Class1"
            Top             =   2040
            Width           =   2775
         End
         Begin VB.OptionButton OptCollOf 
            Caption         =   "Classe Nova."
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   8
            Top             =   2040
            Width           =   1335
         End
         Begin VB.OptionButton OptCollOf 
            Caption         =   "Classe Existente."
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Value           =   -1  'True
            Width           =   1815
         End
      End
      Begin VB.TextBox TxtNome 
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   4695
      End
      Begin VB.TextBox TxtHelpID 
         Height          =   375
         Left            =   -74760
         TabIndex        =   30
         Top             =   3240
         Width           =   4695
      End
      Begin VB.TextBox TxtDesc 
         Height          =   1215
         Left            =   -74760
         MultiLine       =   -1  'True
         TabIndex        =   27
         Top             =   840
         Width           =   4695
      End
      Begin VB.CheckBox ChkTopLevel 
         Caption         =   "Classe é de 1º do nível ?"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   4440
         Width           =   4695
      End
      Begin VB.ComboBox CmbBasedOn_DataType 
         Height          =   315
         ItemData        =   "FrmPropPage.frx":18B8
         Left            =   240
         List            =   "FrmPropPage.frx":18BA
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1320
         Width           =   4695
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Help ID :"
         Height          =   195
         Index           =   4
         Left            =   -74760
         TabIndex        =   31
         Top             =   3000
         Width           =   630
      End
      Begin VB.Label LblHelpFile 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   -74760
         TabIndex        =   29
         Top             =   2400
         Width           =   4695
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Arquivo de Help do Projeto :"
         Height          =   195
         Index           =   3
         Left            =   -74760
         TabIndex        =   28
         Top             =   2160
         Width           =   1995
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição : "
         Height          =   195
         Index           =   2
         Left            =   -74760
         TabIndex        =   26
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Baseado em : "
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   1020
      End
      Begin VB.Label Lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome : "
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   555
      End
   End
   Begin Threed.SSCommand CmdOper 
      Height          =   375
      Index           =   0
      Left            =   4320
      TabIndex        =   23
      Top             =   5280
      WhatsThisHelpID =   10287
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand CmdOper 
      Height          =   375
      Index           =   1
      Left            =   2880
      TabIndex        =   24
      Top             =   5280
      WhatsThisHelpID =   10289
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "&Cancelar"
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FrmPropPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Load()
Event Active()
Event Resize()
Event CmbEscopoChange()
Event CmdArgClick(Index As Integer)
Event CmdOperClick(Index As Integer)
Event ChkisConstClick()
Event LstArgDblClick()
Event OptDeclaraçãoClick(Index As Integer)

Public MyGrd As New MSGrid
Public Function ValidaCampos()
   ValidaCampos = True
End Function
Public Sub Popula_PropPage()
End Sub

Private Sub ChkisConst_Click()
   RaiseEvent ChkisConstClick
End Sub
Private Sub CmbEscopo_Change()
   RaiseEvent CmbEscopoChange
End Sub
Private Sub CmdArg_Click(Index As Integer)
   RaiseEvent CmdArgClick(Index)
End Sub
Private Sub CmdOper_Click(Index As Integer)
   RaiseEvent CmdOperClick(Index)
End Sub
Private Sub Form_Activate()
   RaiseEvent Active
End Sub
Private Sub Form_Load()
   With MyGrd
      Set .Grd = GrdEnumType
      .MaxLin = 5
      .CollLin.Add TxtItem, "0"
   End With

   RaiseEvent Load
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   If Me.ActiveControl.Name = Me.LstArg.Name Then
      Select Case KeyCode
         Case vbKeyInsert: Call CmdArg_Click(0)
         Case vbKeyDelete: Call CmdArg_Click(1)
         Case vbKeyUp: Call CmdArg_Click(2)
         Case vbKeyDown: Call CmdArg_Click(3)
      End Select
   End If
End Sub
Private Sub Form_Resize()
   RaiseEvent Resize
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set MyGrd = Nothing
   Call SetDefault(hwnd)
End Sub

Private Sub GrdEnumType_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyDelete: Call DeleteLinhaItem
   End Select
End Sub

Private Sub GrdEnumType_LeaveCell()
   Call MyGrd.LeaveCell
End Sub

Private Sub GrdEnumType_RowColChange()
   With Me.GrdEnumType
      If .Row = 0 Then Exit Sub
      '* Preencher Linha de Edição
      TxtItem = .TextMatrix(.Row, 0)
      Call SelRowMSGrid(Me.GrdEnumType, .Row)
      Call MyGrd.MoverLinha
     
   End With
End Sub

Private Sub GrdEnumType_Scroll()
   If Me.GrdEnumType.RowPos(Me.GrdEnumType.RowSel) = 0 Then
      Me.GrdEnumType.Row = Me.GrdEnumType.RowSel + 1
   End If
   If Me.GrdEnumType.RowPos(Me.GrdEnumType.RowSel) = (Me.GrdEnumType.Height - 15) Then
      Me.GrdEnumType.Row = Me.GrdEnumType.RowSel - 1
   End If
   Call MyGrd.MoverLinha
End Sub

Private Sub LstArg_DblClick()
   RaiseEvent LstArgDblClick
End Sub

Private Sub OptDeclaração_Click(Index As Integer)
   RaiseEvent OptDeclaraçãoClick(Index)
End Sub
Private Sub TxtItem_KeyDown(KeyCode As Integer, Shift As Integer)
   Call MyGrd.UpDownGrid(KeyCode, Shift)
End Sub

Private Sub TxtItem_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      Call InserirLinhaItem
   End If
End Sub

Private Sub TxtNome_GotFocus()
   Call SelecionarTexto(Me.TxtNome)
End Sub
Public Sub DeleteLinhaItem()
   Dim i%
   Exit Sub
   If Me.GrdEnumType.Row <> 0 And Me.GrdEnumType.Row <> Me.GrdEnumType.Rows - 1 Then
      Me.GrdEnumType.RemoveItem Me.GrdEnumType.Row
      For i = Me.GrdEnumType.Row To Me.GrdEnumType.Rows - 2
         Me.GrdEnumType.TextMatrix(i, 0) = i
      Next
      Call GrdEnumType_RowColChange
   End If
End Sub
Public Sub InserirLinhaItem()
   '* Verificar Linhas Item
   If TxtItem = "" Then
'      Call ExibirAviso(LoadMsg(27) & vbCrLf & Me.GrdEnumType.TextMatrix(0, 1), LoadMsg(1))
      DoEvents
      TxtItem.SetFocus
      Exit Sub
   End If
   
   '* Incluir Linha n Grid
   With Me.GrdEnumType
      If .Row <> (.Rows - 1) Then
         .Row = .Row + 1
      Else
         .Rows = .Rows + 1
         .Row = .Rows - 1
      End If
'      .TextMatrix(.Rows - 2, 0) = .Rows - 2
   End With
   TxtItem.Width = IIf(GrdEnumType.Rows <= MyGrd.MaxLin, 4120, 3880)
   DoEvents
  
   If TxtItem.Visible Then TxtItem.SetFocus
End Sub

