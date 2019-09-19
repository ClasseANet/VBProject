VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmPropAddin 
   AutoRedraw      =   -1  'True
   Caption         =   "Propriedades"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7710
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   7710
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtMicro 
      Height          =   330
      Left            =   2280
      LinkTimeout     =   30
      TabIndex        =   1
      Top             =   5040
      Visible         =   0   'False
      WhatsThisHelpID =   10541
      Width           =   1335
   End
   Begin TabDlg.SSTab TabProp 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   8916
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Geral"
      TabPicture(0)   =   "FrmPropAddin.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frme(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Indentação"
      TabPicture(1)   =   "FrmPropAddin.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ChkIndent(3)"
      Tab(1).Control(1)=   "TxtIndent"
      Tab(1).Control(2)=   "ChkIndent(2)"
      Tab(1).Control(3)=   "ChkIndent(1)"
      Tab(1).Control(4)=   "ChkIndent(0)"
      Tab(1).Control(5)=   "TxtSpcIndent"
      Tab(1).Control(6)=   "Lbl(0)"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Comentários"
      TabPicture(2)   =   "FrmPropAddin.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Lbl(1)"
      Tab(2).Control(1)=   "Lbl(8)"
      Tab(2).Control(2)=   "TxtCharComment"
      Tab(2).Control(3)=   "TxtUserComment"
      Tab(2).Control(4)=   "Frme(2)"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Tratamento de Erro"
      TabPicture(3)   =   "FrmPropAddin.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frme(0)"
      Tab(3).ControlCount=   1
      Begin VB.CheckBox ChkIndent 
         Caption         =   "Insert a line blank before procedure"
         Height          =   495
         Index           =   3
         Left            =   -70560
         TabIndex        =   49
         Top             =   3120
         Width           =   2895
      End
      Begin VB.TextBox TxtIndent 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4575
         Left            =   -74880
         MultiLine       =   -1  'True
         TabIndex        =   46
         Text            =   "FrmPropAddin.frx":0070
         Top             =   360
         Width           =   4215
      End
      Begin VB.CheckBox ChkIndent 
         Caption         =   "Indent with ""Select Cade"" lines"
         Height          =   255
         Index           =   2
         Left            =   -70560
         TabIndex        =   45
         Top             =   2880
         Width           =   2895
      End
      Begin VB.CheckBox ChkIndent 
         Caption         =   "Indent comments to alig with code"
         Height          =   255
         Index           =   1
         Left            =   -70560
         TabIndex        =   44
         Top             =   2520
         Width           =   2895
      End
      Begin VB.CheckBox ChkIndent 
         Caption         =   "Indent everything within procedure"
         Height          =   255
         Index           =   0
         Left            =   -70560
         TabIndex        =   43
         Top             =   2160
         Width           =   2895
      End
      Begin VB.Frame Frme 
         Caption         =   "Modelo"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   3495
         Index           =   2
         Left            =   -74640
         TabIndex        =   14
         Top             =   360
         Width           =   6735
         Begin VB.CheckBox ChkItemSel 
            Caption         =   "Exibir Somente Itens Selecionados"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   240
            TabIndex        =   37
            Top             =   3210
            Width           =   3615
         End
         Begin VB.ListBox LstItens 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   2610
            Left            =   240
            Style           =   1  'Checkbox
            TabIndex        =   34
            Top             =   600
            Width           =   6375
         End
         Begin VB.CommandButton CmdModelo 
            Caption         =   "&Excluir"
            Height          =   320
            Index           =   1
            Left            =   5890
            TabIndex        =   33
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton CmdModelo 
            Caption         =   "&Salvar"
            Height          =   320
            Index           =   0
            Left            =   5160
            TabIndex        =   32
            Top             =   240
            Width           =   735
         End
         Begin VB.ComboBox CmbNmModelo 
            Height          =   315
            Left            =   1800
            TabIndex        =   31
            Text            =   "Combo1"
            Top             =   240
            Width           =   3375
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Nome do Modelo :"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   225
            Index           =   11
            Left            =   240
            TabIndex        =   15
            Top             =   240
            WhatsThisHelpID =   10540
            Width           =   1500
         End
      End
      Begin VB.TextBox TxtUserComment 
         Height          =   615
         Left            =   -74640
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   36
         Top             =   4080
         Width           =   6735
      End
      Begin VB.TextBox TxtSpcIndent 
         Height          =   285
         Left            =   -69480
         TabIndex        =   29
         Text            =   "3"
         Top             =   1320
         Width           =   735
      End
      Begin VB.Frame Frme 
         Caption         =   "Tratamento de Erro"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   4215
         Index           =   0
         Left            =   -74760
         TabIndex        =   16
         Top             =   480
         Width           =   6855
         Begin VB.CheckBox ChkLabelSaida 
            Height          =   255
            Left            =   30
            TabIndex        =   21
            Top             =   1560
            Value           =   1  'Checked
            Width           =   200
         End
         Begin VB.TextBox TxtErrorLabel 
            Height          =   285
            Left            =   1320
            TabIndex        =   20
            Text            =   "Trata_Erro"
            Top             =   360
            Width           =   1935
         End
         Begin VB.TextBox TxtErrorFunction 
            Height          =   525
            Left            =   480
            MultiLine       =   -1  'True
            TabIndex        =   19
            Text            =   "FrmPropAddin.frx":0222
            Top             =   3120
            Width           =   2775
         End
         Begin VB.TextBox TxtSaidaLabel 
            Height          =   285
            Left            =   240
            TabIndex        =   18
            Text            =   "Saida"
            Top             =   1560
            Width           =   975
         End
         Begin VB.TextBox TxtSaidaFunction 
            Height          =   525
            Left            =   480
            MultiLine       =   -1  'True
            TabIndex        =   17
            Text            =   "FrmPropAddin.frx":0245
            Top             =   1920
            Width           =   2775
         End
         Begin VB.Label LblOnError 
            AutoSize        =   -1  'True
            Caption         =   "On Error Goto "
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   28
            Top             =   360
            Width           =   1020
         End
         Begin VB.Label LblOnError 
            Caption         =   $"FrmPropAddin.frx":0267
            Height          =   675
            Index           =   1
            Left            =   240
            TabIndex        =   27
            Top             =   720
            Width           =   3015
         End
         Begin VB.Label LblErrorLabel 
            AutoSize        =   -1  'True
            Caption         =   "Trata_Error: "
            Height          =   195
            Left            =   240
            TabIndex        =   26
            Top             =   2880
            Width           =   885
         End
         Begin VB.Label LblOnError 
            AutoSize        =   -1  'True
            Caption         =   ": "
            Height          =   195
            Index           =   2
            Left            =   1245
            TabIndex        =   25
            Top             =   1560
            Width           =   90
         End
         Begin VB.Label LblSaidaLabel 
            AutoSize        =   -1  'True
            Caption         =   "Saida"
            Height          =   195
            Left            =   885
            TabIndex        =   24
            Top             =   3720
            Width           =   405
         End
         Begin VB.Label LblOnError 
            AutoSize        =   -1  'True
            Caption         =   "Goto "
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   4
            Left            =   480
            TabIndex        =   23
            Top             =   3720
            Width           =   390
         End
         Begin VB.Label LblOnError 
            AutoSize        =   -1  'True
            Caption         =   "Exit Function"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   3
            Left            =   480
            TabIndex        =   22
            Top             =   2520
            Width           =   915
         End
      End
      Begin VB.Frame Frme 
         Caption         =   "Configuração Geral"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   3855
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   6735
         Begin VB.TextBox TxtExeAuxiliar 
            Height          =   330
            Left            =   1560
            LinkTimeout     =   30
            TabIndex        =   51
            Top             =   360
            WhatsThisHelpID =   10528
            Width           =   4755
         End
         Begin VB.CommandButton CmdDrv 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   " ..."
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   6360
            TabIndex        =   50
            Top             =   360
            WhatsThisHelpID =   10527
            Width           =   315
         End
         Begin VB.ComboBox CmbIdioma 
            Height          =   315
            ItemData        =   "FrmPropAddin.frx":0325
            Left            =   1560
            List            =   "FrmPropAddin.frx":0341
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   3120
            WhatsThisHelpID =   10522
            Width           =   1665
         End
         Begin VB.CommandButton CmdDrv 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   " ..."
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   6360
            TabIndex        =   39
            Top             =   720
            WhatsThisHelpID =   10527
            Width           =   315
         End
         Begin VB.TextBox TxtDbLib 
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   330
            Left            =   1560
            LinkTimeout     =   30
            TabIndex        =   38
            Top             =   720
            WhatsThisHelpID =   10528
            Width           =   4755
         End
         Begin VB.TextBox TxtTelefone 
            Height          =   330
            Left            =   1560
            LinkTimeout     =   30
            TabIndex        =   12
            Top             =   2640
            WhatsThisHelpID =   10528
            Width           =   4755
         End
         Begin VB.TextBox TxteMail 
            Height          =   330
            Left            =   1560
            LinkTimeout     =   30
            TabIndex        =   10
            Top             =   2160
            WhatsThisHelpID =   10528
            Width           =   4755
         End
         Begin VB.TextBox TxtWebSite 
            Height          =   330
            Left            =   1560
            LinkTimeout     =   30
            TabIndex        =   8
            Top             =   1680
            WhatsThisHelpID =   10528
            Width           =   4755
         End
         Begin VB.TextBox TxtDesenv 
            Height          =   330
            Left            =   1560
            LinkTimeout     =   30
            TabIndex        =   6
            Top             =   1200
            WhatsThisHelpID =   10528
            Width           =   4755
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Prog. Auxiliar  :"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   225
            Index           =   10
            Left            =   240
            TabIndex        =   52
            Top             =   360
            WhatsThisHelpID =   10540
            Width           =   1320
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Idioma             :"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   225
            Index           =   9
            Left            =   240
            TabIndex        =   48
            Top             =   3240
            WhatsThisHelpID =   10540
            Width           =   1215
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "DataBase          :"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   225
            Index           =   7
            Left            =   240
            TabIndex        =   40
            Top             =   720
            WhatsThisHelpID =   10540
            Width           =   1305
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Telefone             :"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   225
            Index           =   5
            Left            =   240
            TabIndex        =   13
            Top             =   2760
            WhatsThisHelpID =   10540
            Width           =   1335
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   """e-Mail""            :"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   225
            Index           =   4
            Left            =   240
            TabIndex        =   11
            Top             =   2280
            WhatsThisHelpID =   10540
            Width           =   1320
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   """Web Site""        :"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   225
            Index           =   3
            Left            =   240
            TabIndex        =   9
            Top             =   1680
            WhatsThisHelpID =   10540
            Width           =   1305
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Desenvolvedor :"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   225
            Index           =   2
            Left            =   240
            TabIndex        =   7
            Top             =   1200
            WhatsThisHelpID =   10540
            Width           =   1320
         End
      End
      Begin VB.TextBox TxtCharComment 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -68760
         TabIndex        =   41
         Text            =   "'* "
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Caracter :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   8
         Left            =   -69600
         TabIndex        =   42
         Top             =   3840
         WhatsThisHelpID =   10540
         Width           =   840
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Espaços :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   0
         Left            =   -70440
         TabIndex        =   30
         Top             =   1320
         WhatsThisHelpID =   10540
         Width           =   780
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Comentário do Usuário :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   1
         Left            =   -74640
         TabIndex        =   35
         Top             =   3840
         WhatsThisHelpID =   10540
         Width           =   2040
      End
   End
   Begin Threed.SSCommand CmdOper 
      Height          =   405
      Index           =   1
      Left            =   5160
      TabIndex        =   2
      Top             =   5400
      WhatsThisHelpID =   10543
      Width           =   945
      _Version        =   65536
      _ExtentX        =   1667
      _ExtentY        =   714
      _StockProps     =   78
      ForeColor       =   255
      Picture         =   "FrmPropAddin.frx":036B
   End
   Begin Threed.SSCommand CmdOper 
      Height          =   405
      Index           =   0
      Left            =   6600
      TabIndex        =   3
      Top             =   5400
      WhatsThisHelpID =   10542
      Width           =   945
      _Version        =   65536
      _ExtentX        =   1667
      _ExtentY        =   714
      _StockProps     =   78
      Picture         =   "FrmPropAddin.frx":0DA5
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Micro"
      Height          =   195
      Index           =   6
      Left            =   2640
      TabIndex        =   4
      Top             =   4320
      Visible         =   0   'False
      WhatsThisHelpID =   10544
      Width           =   390
   End
End
Attribute VB_Name = "FrmPropAddin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public FrmTemplates As Collection
Public EcrevendoTela  As Boolean
Public CarregandoConfig As Boolean
Public Function ValidaCampos()
   ValidaCampos = False
'   If Right(Me.TxtDbDrive, 1) <> "\" Then Me.TxtDbDrive = Me.TxtDbDrive & "\"
'   If Right(Me.TxtDrvRpt, 1) <> "\" Then Me.TxtDrvRpt = Me.TxtDrvRpt & "\"
'   If Trim$(Me.TxtDbDrive) = "" Or Not FileExists(Me.TxtDbDrive + Me.TxtDbName) Then
'      Call ExibirAviso(LoadRes("S27"), LoadRes("S1"))
'      Call Set_Focus(Me.TxtDbDrive)
'      Exit Function
'   End If
'   If Trim$(Me.TxtDrvRpt) = "" Then
'      Call ExibirAviso(LoadRes("S27"), LoadRes("S1"))
'      Call Set_Focus(Me.TxtDrvRpt)
'      Exit Function
'   End If
'   If Trim$(Me.TxtDbName) = "" Or Not FileExists(Me.TxtDbDrive + Me.TxtDbName) Then
'      Call ExibirAviso(LoadRes("S27"), LoadRes("S1"))
'      Call Set_Focus(Me.TxtDbName)
'      Exit Function
'   End If
   ValidaCampos = True
End Function

Private Sub ChkIndent_Click(Index As Integer)
   Dim IndComm As Boolean
   Dim IndFunc As Boolean
   Dim IndSele As Boolean
   Dim LineBlank As Boolean
   Dim SpcInd As Integer
   
   If CarregandoConfig Then Exit Sub
   IndComm = (Me.ChkIndent(1).Value = vbChecked)
   IndFunc = (Me.ChkIndent(0).Value = vbChecked)
   IndSele = (Me.ChkIndent(2).Value = vbChecked)
   LineBlank = (Me.ChkIndent(3).Value = vbChecked)
   SpcInd = Me.TxtSpcIndent.Text
   Me.TxtIndent.Text = IndentarFuncao(, , Me.TxtIndent.Text, IndComm, IndFunc, IndSele, SpcInd, LineBlank)
End Sub

Private Sub ChkItemSel_Click()
   Dim i%, StrAux$
   If ChkItemSel.Value Then
      For i = Me.LstItens.ListCount - 1 To 0 Step -1
         If Not Me.LstItens.Selected(i) Then
            Me.LstItens.RemoveItem i
         End If
      Next
   Else
      If Me.LstItens.SelCount > 0 Then
         For i = 0 To Me.LstItens.ListCount - 1
            If Not Me.LstItens.Selected(i) Then
               StrAux$ = StrAux$ & ItemModeloIndex(Me.LstItens.List(i)) & "|"
            End If
         Next
      End If
      Me.LstItens.Clear
      Call MontarLstItens
      Call SetModelo(StrAux$)
   End If
End Sub

Private Sub ChkLabelSaida_Click()
   Me.TxtSaidaLabel.Enabled = (ChkLabelSaida.Value = vbChecked)
   Call HabilitaLabelSaida(ChkLabelSaida.Value = vbChecked)
   If ChkLabelSaida.Value = vbChecked Then
      Me.TxtSaidaLabel.SetFocus
   End If
End Sub
Private Sub CmbIdioma_Click()
   If EcrevendoTela Then Exit Sub
   If Me.CmbIdioma.ListIndex < 0 Then Exit Sub
   If Me.CmbIdioma.ListIndex <= 1 Then
      Sys.Proj.Idioma = Me.CmbIdioma.ItemData(Me.CmbIdioma.ListIndex)
      Sys.Edit.Idioma = Me.CmbIdioma.ItemData(Me.CmbIdioma.ListIndex)
      Call EscreveTela
   Else
      Me.CmbIdioma.ListIndex = -1
   End If
End Sub

Private Sub CmbNmModelo_Change()
   Dim i%
   If Me.CmbNmModelo.ListIndex = -1 Then
      Me.CmdModelo(0).Enabled = (CmbNmModelo.ListIndex <> 0)
      Me.CmdModelo(1).Enabled = (CmbNmModelo.ListIndex <> 0)
      If Len(Me.CmbNmModelo) <= 1 Then
         For i = 0 To Me.LstItens.ListCount - 1
             Me.LstItens.Selected(i) = False
         Next
      End If
      If Me.ChkItemSel.Value = vbChecked Then
         Me.ChkItemSel.Value = vbUnchecked
      End If
   End If
End Sub

Private Sub CmbNmModelo_Click()
   Dim StrAux$, Pos%
   Dim Bool As Boolean
   If CmbNmModelo.ListIndex < 0 Then Exit Sub
   Me.CmdModelo(0).Enabled = (CmbNmModelo.ListIndex <> 0)
   Me.CmdModelo(1).Enabled = (CmbNmModelo.ListIndex <> 0)
   Bool = (Me.ChkItemSel.Value = vbChecked)
   If Bool Then
      Me.ChkItemSel.Value = vbUnchecked
   End If
   StrAux$ = FrmTemplates(CmbNmModelo.ListIndex + 1)
   Pos = InStr(StrAux$, " ")
   If Pos > 0 Then
      StrAux = Mid(StrAux$, Pos + 1)
      Call SetModelo(StrAux)
   End If
   If Bool Then
      Me.ChkItemSel.Value = vbChecked
   End If
End Sub

Private Sub CmbNmModelo_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeySpace Then
      KeyAscii = 0
   End If
End Sub
Private Sub CmdDrv_Click(Index As Integer)
   Dim Arq As String
   Select Case Index
      Case 0
         Arq = ProcurarArquivo(SysMdi.CmDialog, LoadRes(gwAbrir) & " " & LoadRes(gwArquivo), , LoadRes(gwArquivo) & " (*.exe)|*.exe")
         If Trim(Arq) <> "" Then
            Sys.Edit.ExeAuxiliar = SysMdi.CmDialog.Tag & Arq
            Me.TxtExeAuxiliar.Text = SysMdi.CmDialog.Tag & Arq
         End If
      Case 1: Call GetPath(hwnd, Me.TxtDbLib)
'      Case 2: Call GetPath(hwnd, Me.TxtDrvRpt)
'       Case 3: Call Getarquivo(hwnd, Me.TxtDbDrive)
   End Select
End Sub

Private Sub CmdModelo_Click(Index As Integer)
   Dim i%, StrAux$
   Dim Achou As Boolean, IndAchou%
   If Me.LstItens.SelCount = 0 Or Me.CmbNmModelo.Text = Me.CmbNmModelo.List(0) Then
      Me.CmbNmModelo.ListIndex = 0
      Exit Sub
   End If
   If Me.CmbNmModelo.Text = Me.CmbNmModelo.List(0) Then
      Me.CmbNmModelo.ListIndex = 0
      Exit Sub
   End If
      
   Select Case Index
      Case 0
         IndAchou = LocalizarCombo(Me.CmbNmModelo, Me.CmbNmModelo.Text, False)
         Achou = (IndAchou >= 0)
         For i = 0 To Me.LstItens.ListCount - 1
            If Me.LstItens.Selected(i) Then
               StrAux$ = StrAux$ & ItemModeloIndex(Me.LstItens.List(i)) & "|"
            End If
         Next
         
         If Achou Then FrmTemplates.Remove IndAchou + 1
         FrmTemplates.Add Me.CmbNmModelo.Text & " " & StrAux$, Me.CmbNmModelo.Text
         If Not Achou Then Me.CmbNmModelo.AddItem Me.CmbNmModelo.Text

      Case 1
         IndAchou = LocalizarCombo(Me.CmbNmModelo, Me.CmbNmModelo.Text, False)
         If IndAchou >= 0 Then
            FrmTemplates.Remove IndAchou + 1
            Me.CmbNmModelo.RemoveItem IndAchou
            Me.CmbNmModelo.ListIndex = 0
         Else
            Me.CmbNmModelo.ListIndex = Me.CmbNmModelo.ListCount - 1
         End If
   End Select
End Sub

Private Sub CmdOper_Click(Index As Integer)
   Select Case Index
      Case 0
         If ValidaCampos() Then
            Call CmdModelo_Click(0)
            Call UpdateConfig
            Call SaveConfig
         Else
            Exit Sub
         End If
'      Case 1: BANCO.TB_CONFIG.Cancelado = True
   End Select
   Unload Me
End Sub

Private Sub Form_Activate()
   Call CarregaConfig
   Me.TabProp.Tab = 0
   With Sys.Edit
      Me.TxtIndent.Text = IndentarFuncao(, , Me.TxtIndent.Text, .IndentComment, .IndentFunction, .IndentSelect, .SpcIndent)
   End With
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'   KeyAscii = SendTab(Me, KeyAscii)
End Sub

Private Sub Form_Load()
   Dim i%, Pos%
   
   
   Call EscreveTela
   
   Call ConfigForm(Me, SysMdi.Icon, Sys.Proj.FundoTela)
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
 '  Set sys.MDIFilho = Nothing
   Set FrmTemplates = Nothing
End Sub
Private Sub TxtErrorLabel_Change()
   Me.LblErrorLabel = Trim(Me.TxtErrorLabel) & ":"
End Sub

Private Sub TxtMicro_GotFocus()
   Call SelecionarTexto(Me.ActiveControl)
End Sub

Private Sub TxtSaidaLabel_Change()
   Me.LblSaidaLabel = Trim(Me.TxtSaidaLabel)
   Me.ChkLabelSaida.Value = IIf(Trim(Me.TxtSaidaLabel.Text) = "", vbUnchecked, vbChecked)
End Sub
Private Sub HabilitaLabelSaida(Bool As Boolean)
   Me.TxtSaidaFunction.Enabled = Bool
   Me.LblOnError(2).Visible = Bool
   Me.LblOnError(3).Enabled = Bool
   Me.LblOnError(4).Visible = Bool
   Me.LblSaidaLabel.Visible = Bool
End Sub
Private Sub TxtSpcIndent_Change()
   Dim IndComm As Boolean
   Dim IndFunc As Boolean
   Dim IndSele As Boolean
   Dim SpcInd As Integer
      
   If CarregandoConfig Then Exit Sub
   IndComm = (Me.ChkIndent(1).Value = vbChecked)
   IndFunc = (Me.ChkIndent(0).Value = vbChecked)
   IndSele = (Me.ChkIndent(2).Value = vbChecked)
   SpcInd = Val(Me.TxtSpcIndent.Text)
   Me.TxtIndent.Text = IndentarFuncao("", Me.TxtIndent.Text, IndComm, IndFunc, IndSele, SpcInd)
End Sub
Private Sub TxtSpcIndent_KeyPress(KeyAscii As Integer)
   If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack Then
      KeyAscii = 0
   End If
End Sub
Private Sub UpdateConfig()
   Dim i%, n As Variant
   With Sys
      With .Edit
         .ExeAuxiliar = Me.TxtExeAuxiliar.Text
         
         .Desenvolvedor = Me.TxtDesenv.Text
         .WebSite = Me.TxtWebSite.Text
         .eMail = Me.TxteMail.Text
         .Telefone = Me.TxtTelefone.Text
         .Idioma = Me.CmbIdioma.ItemData(Me.CmbIdioma.ListIndex)
         .ErrorLabel = Me.TxtErrorLabel.Text
         .ErrorFunction = Me.TxtErrorFunction.Text
         .SaidaLabel = Me.TxtSaidaLabel.Text
         .SaidaFunction = Me.TxtSaidaFunction.Text
         .SpcIndent = CInt(Me.TxtSpcIndent.Text)
         
         
         .CharComment = Me.TxtCharComment.Text
         .UserComment = Me.TxtUserComment.Text
         Set .Templates = Nothing
         Set .Templates = New Collection
         For Each n In FrmTemplates
            .Templates.Add n, Mid(n, 1, InStr(n, " ") - 1)
         Next
         If Me.CmbNmModelo.ListIndex < 0 Then
            If LocalizarCombo(Me.CmbNmModelo, .Template, False) < 0 Then
               .Template = Me.CmbNmModelo.List(0)
            End If
         Else
            .Template = Me.CmbNmModelo.Text
         End If
         .IndentFunction = (Me.ChkIndent(0).Value = vbChecked)
         .IndentComment = (Me.ChkIndent(1).Value = vbChecked)
         .IndentSelect = (Me.ChkIndent(2).Value = vbChecked)
         .LineBlankBefore = (Me.ChkIndent(3).Value = vbChecked)
         
      End With
   End With
End Sub
Public Sub CarregaConfig()
   Dim n As Variant
   Dim i%
   CarregandoConfig = True
   With Sys
      With .Edit
         Me.TxtExeAuxiliar.Text = .ExeAuxiliar
         
         Me.TxtDesenv.Text = .Desenvolvedor
         Me.TxtWebSite.Text = .WebSite
         Me.TxteMail.Text = .eMail
         Me.TxtTelefone.Text = .Telefone
         Select Case .Idioma
            Case 5000: Me.CmbIdioma.ListIndex = 0 '* "Português"
            Case 6000: Me.CmbIdioma.ListIndex = 1 '* "Inglês"
            Case 7000: Me.CmbIdioma.ListIndex = 2 '* "Francês"
            Case 8000: Me.CmbIdioma.ListIndex = 3 '* "Espanhol"
         End Select
         Me.TxtErrorLabel.Text = .ErrorLabel
         Me.TxtErrorFunction.Text = .ErrorFunction
         Me.TxtSaidaLabel.Text = .SaidaLabel
         Me.TxtSaidaFunction.Text = .SaidaFunction
         Me.TxtSpcIndent.Text = CStr(.SpcIndent)
         Me.ChkLabelSaida.Value = IIf(Trim(.SaidaLabel) = "", vbUnchecked, vbChecked)
         
         Me.TxtCharComment.Text = .CharComment
         Me.TxtUserComment.Text = .UserComment
         Set FrmTemplates = New Collection
         For Each n In .Templates
            FrmTemplates.Add n, Mid(n, 1, InStr(n, " ") - 1)
         Next
         Call MontarLstItens
         For Each n In FrmTemplates
            If InStr(n, " ") > 1 Then
               Me.CmbNmModelo.AddItem Mid(n, 1, InStr(n, " ") - 1)
            End If
         Next
         Me.CmbNmModelo.ListIndex = LocalizarCombo(Me.CmbNmModelo, .Template, False)
         Me.ChkIndent(0).Value = IIf(.IndentFunction, vbChecked, vbUnchecked)
         Me.ChkIndent(1).Value = IIf(.IndentComment, vbChecked, vbUnchecked)
         Me.ChkIndent(2).Value = IIf(.IndentSelect, vbChecked, vbUnchecked)
         Me.ChkIndent(3).Value = IIf(.LineBlankBefore, vbChecked, vbUnchecked)
      End With
   End With
   CarregandoConfig = False
End Sub
Public Sub MontarLstItens()
   Dim i%
   With Me.LstItens
      For i = 1 To 12
         .AddItem ItemModelo(i)
         .Selected(.NewIndex) = True
      Next
   End With
End Sub
Public Sub SetModelo(Itens$)
   Dim i%
   For i = 0 To Me.LstItens.ListCount - 1
      Me.LstItens.Selected(i) = (InStr(Itens, CStr(i + 1) & "|") <> 0)
   Next
   Me.LstItens.ListIndex = 0
End Sub
Public Function ItemModelo(Index As Integer) As String
  Select Case Index
      Case 1:    ItemModelo = LoadRes(gwDesenvolvedor)
      Case 2:    ItemModelo = LoadRes(gwWebSite)
      Case 3:    ItemModelo = LoadRes(gweMail)
      Case 4:    ItemModelo = LoadRes(gwTelefone)
      Case 5:    ItemModelo = LoadRes(gwData_Hora)
      Case 6:    ItemModelo = LoadRes(gwNome_do_Projeto)
      Case 7:    ItemModelo = LoadRes(gwNome_do_Modulo)
      Case 8:    ItemModelo = LoadRes(gwNome_do_Arquivo)
      Case 9:    ItemModelo = LoadRes(gwNome_da_Funcao)
      Case 10:   ItemModelo = LoadRes(gwParametros)
      Case 11:   ItemModelo = LoadRes(gwComentario)
      Case 12:   ItemModelo = LoadRes(gwComentario_do_Usuario)
      Case Else: ItemModelo = ""
   End Select
End Function
Public Function ItemModeloIndex(Item As String) As Integer
  Select Case Item
      Case LoadRes(gwDesenvolvedor):         ItemModeloIndex = 1
      Case LoadRes(gwWebSite):               ItemModeloIndex = 2
      Case LoadRes(gweMail):                 ItemModeloIndex = 3
      Case LoadRes(gwTelefone):              ItemModeloIndex = 4
      Case LoadRes(gwData_Hora):             ItemModeloIndex = 5
      Case LoadRes(gwNome_do_Projeto):       ItemModeloIndex = 6
      Case LoadRes(gwNome_do_Modulo):        ItemModeloIndex = 7
      Case LoadRes(gwNome_do_Arquivo):       ItemModeloIndex = 8
      Case LoadRes(gwNome_da_Funcao):        ItemModeloIndex = 9
      Case LoadRes(gwParametros):            ItemModeloIndex = 10
      Case LoadRes(gwComentario):            ItemModeloIndex = 11
      Case LoadRes(gwComentario_do_Usuario): ItemModeloIndex = 12
      Case Else: ItemModeloIndex = 0
   End Select
End Function

Private Sub EscreveTela()
   Dim Ind%
   EcrevendoTela = True
   With Me
      .Caption = LoadRes(gwConfigurar) & App.ProductName
      .TabProp.TabCaption(0) = LoadRes(gwGeral)
         .Frme(1).Caption = LoadRes(gwConfiguracao)
         .Lbl(10).Caption = LoadRes(gwProg_Auxiliar) & Space(18 - Len(LoadRes(gwProg_Auxiliar))) & ":"
         .Lbl(7).Caption = LoadRes(gwBanco_de_Dados) & Space(18 - Len(LoadRes(gwBanco_de_Dados))) & ":"
         .Lbl(2).Caption = LoadRes(gwDesenvolvedor) & Space(18 - Len(LoadRes(gwDesenvolvedor))) & ":"
         .Lbl(3).Caption = LoadRes(gwWebSite) & Space(18 - Len(LoadRes(gwWebSite))) & ":"
         .Lbl(4).Caption = LoadRes(gweMail) & Space(18 - Len(LoadRes(gweMail))) & ":"
         .Lbl(5).Caption = LoadRes(gwTelefone) & Space(18 - Len(LoadRes(gwTelefone))) & ":"
         .Lbl(9).Caption = LoadRes(gwIdioma) & Space(18 - Len(LoadRes(gwIdioma))) & ":"
         Ind% = CmbIdioma.ListIndex
         .CmbIdioma.Clear
         .CmbIdioma.AddItem LoadRes(gwPortugues)
         .CmbIdioma.ItemData(.CmbIdioma.NewIndex) = 5000
         .CmbIdioma.AddItem LoadRes(gwIngles)
         .CmbIdioma.ItemData(.CmbIdioma.NewIndex) = 6000
         .CmbIdioma.AddItem LoadRes(gwFrances)
         .CmbIdioma.ItemData(.CmbIdioma.NewIndex) = 7000
         .CmbIdioma.AddItem LoadRes(gwEspanhol)
         .CmbIdioma.ItemData(.CmbIdioma.NewIndex) = 8000
         .CmbIdioma.ListIndex = Ind%
      .TabProp.TabCaption(1) = LoadRes(gwIndentacao)
         .Lbl(0).Caption = LoadRes(gwEspacos)
         .ChkIndent(0).Caption = LoadRes(gwIndentar_Com_Função)
         .ChkIndent(1).Caption = LoadRes(gwIndentar_Comentário)
         .ChkIndent(2).Caption = LoadRes(gwIndentar_Select_Case)
         .ChkIndent(3).Caption = LoadRes(gwInserir_Linha_em_Branco_Antes_da_Funcao)
      .TabProp.TabCaption(2) = LoadRes(gwComentario)
         .Frme(2).Caption = LoadRes(gwModelo)
         .Lbl(11).Caption = LoadRes(gwNome_do_Modelo)
         .Lbl(1).Caption = LoadRes(gwComentario_do_Usuario)
         .Lbl(8).Caption = LoadRes(gwCaracter)
         .CmdModelo(0).Caption = "&" & LoadRes(gwSalvar)
         .CmdModelo(1).Caption = "&" & LoadRes(gwExcluir)
         .ChkItemSel.Caption = LoadRes(gwExibir_Somente_Itens_Selecionados)
      .TabProp.TabCaption(3) = LoadRes(gwTratamento_de_Erro)
         .Frme(0).Caption = LoadRes(gwTratamento_de_Erro)
   End With
   EcrevendoTela = False
End Sub

