VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FrmSeg 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Segurança"
   ClientHeight    =   5460
   ClientLeft      =   2835
   ClientTop       =   2010
   ClientWidth     =   7440
   Icon            =   "Seg.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5460
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab TabSeg 
      Height          =   4020
      Left            =   360
      TabIndex        =   27
      Top             =   240
      Width           =   6720
      _ExtentX        =   11853
      _ExtentY        =   7091
      _Version        =   393216
      Tab             =   1
      TabHeight       =   529
      TabCaption(0)   =   "&Grupo Usuário"
      TabPicture(0)   =   "Seg.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frme(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frme(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "&Usuário"
      TabPicture(1)   =   "Seg.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frme(2)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frme(3)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "&Permissão Acesso"
      TabPicture(2)   =   "Seg.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frme(5)"
      Tab(2).Control(1)=   "Frme(4)"
      Tab(2).ControlCount=   2
      Begin VB.Frame Frme 
         Caption         =   "&Módulos"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   2625
         Index           =   5
         Left            =   -74775
         TabIndex        =   21
         Top             =   1260
         Width           =   6360
         Begin ComctlLib.TreeView TreModu 
            Height          =   2025
            Left            =   150
            TabIndex        =   22
            Top             =   180
            Width           =   6045
            _ExtentX        =   10663
            _ExtentY        =   3572
            _Version        =   327682
            HideSelection   =   0   'False
            LabelEdit       =   1
            Style           =   7
            Appearance      =   1
         End
         Begin VB.CheckBox ChkAcesso 
            Caption         =   "   Alteração"
            Height          =   195
            Index           =   3
            Left            =   4950
            TabIndex        =   26
            Top             =   2300
            Value           =   1  'Checked
            Width           =   1185
         End
         Begin VB.CheckBox ChkAcesso 
            Caption         =   "   Inclusão"
            Height          =   195
            Index           =   2
            Left            =   3270
            TabIndex        =   25
            Top             =   2300
            Value           =   1  'Checked
            Width           =   1035
         End
         Begin VB.CheckBox ChkAcesso 
            Caption         =   "   Exclusão"
            Height          =   195
            Index           =   1
            Left            =   1725
            TabIndex        =   24
            Top             =   2300
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox ChkAcesso 
            Caption         =   "   Leitura"
            Height          =   195
            Index           =   0
            Left            =   300
            TabIndex        =   23
            Top             =   2300
            Value           =   1  'Checked
            Width           =   1050
         End
      End
      Begin VB.Frame Frme 
         Caption         =   "&Identificação"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   885
         Index           =   4
         Left            =   -74760
         TabIndex        =   18
         Top             =   360
         Width           =   6360
         Begin VB.ComboBox CmbGrpUser 
            DataSource      =   "Data1"
            Height          =   315
            Index           =   1
            ItemData        =   "Seg.frx":0496
            Left            =   840
            List            =   "Seg.frx":0498
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   360
            Width           =   4680
         End
         Begin VB.Label Lbl 
            Caption         =   "Grupo Usuário"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   460
            Index           =   6
            Left            =   90
            TabIndex        =   19
            Top             =   270
            Width           =   645
         End
      End
      Begin VB.Frame Frme 
         Caption         =   "&Dados"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1860
         Index           =   3
         Left            =   195
         TabIndex        =   3
         Top             =   1470
         Width           =   6360
         Begin VB.TextBox TxtPwdUser 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   4440
            LinkTimeout     =   30
            MaxLength       =   10
            PasswordChar    =   "*"
            TabIndex        =   9
            Top             =   825
            Width           =   1080
         End
         Begin VB.TextBox TxtPwdUser 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   870
            LinkTimeout     =   30
            MaxLength       =   10
            PasswordChar    =   "*"
            TabIndex        =   7
            Top             =   810
            Width           =   1080
         End
         Begin VB.TextBox TxtNmUser 
            Height          =   330
            Left            =   870
            LinkTimeout     =   30
            MaxLength       =   30
            TabIndex        =   5
            Top             =   315
            Width           =   5310
         End
         Begin VB.ComboBox CmbGrpUser 
            DataSource      =   "Data1"
            Height          =   315
            Index           =   0
            ItemData        =   "Seg.frx":049A
            Left            =   885
            List            =   "Seg.frx":049C
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1260
            Width           =   5355
         End
         Begin VB.Label LblGrp 
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   255
            Left            =   960
            TabIndex        =   31
            Top             =   1320
            Width           =   4815
         End
         Begin VB.Label Lbl 
            Caption         =   "Confirme  sua Senha"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   3
            Left            =   2640
            TabIndex        =   8
            Top             =   915
            Width           =   1800
         End
         Begin VB.Label Lbl 
            Caption         =   "Senha"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   180
            TabIndex        =   6
            Top             =   885
            Width           =   555
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Nome"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   180
            TabIndex        =   4
            Top             =   420
            Width           =   465
         End
         Begin VB.Label Lbl 
            Caption         =   "Grupo Usuário"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   460
            Index           =   4
            Left            =   180
            TabIndex        =   10
            Top             =   1200
            Width           =   640
         End
      End
      Begin VB.Frame Frme 
         Caption         =   "&Identificação"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   885
         Index           =   2
         Left            =   195
         TabIndex        =   0
         Top             =   480
         Width           =   6360
         Begin VB.TextBox TxtIdUser 
            Height          =   285
            Left            =   915
            MaxLength       =   10
            TabIndex        =   2
            Top             =   405
            Width           =   1065
         End
         Begin VB.Label LblId 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   1
            Left            =   180
            MouseIcon       =   "Seg.frx":049E
            MousePointer    =   99  'Custom
            TabIndex        =   1
            Top             =   405
            Width           =   570
         End
      End
      Begin VB.Frame Frme 
         Caption         =   "&Dados"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   900
         Index           =   1
         Left            =   -74805
         TabIndex        =   15
         Top             =   1920
         Width           =   6360
         Begin VB.TextBox TxtDscGrpUser 
            Height          =   330
            Left            =   885
            MaxLength       =   30
            TabIndex        =   17
            Top             =   330
            Width           =   5295
         End
         Begin VB.Label Lbl 
            AutoSize        =   -1  'True
            Caption         =   "Nome"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   225
            TabIndex        =   16
            Top             =   435
            Width           =   465
         End
      End
      Begin VB.Frame Frme 
         Caption         =   "&Identificação"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   885
         Index           =   0
         Left            =   -74805
         TabIndex        =   12
         Top             =   720
         Width           =   6360
         Begin MSMask.MaskEdBox MskIdGrpUser 
            Height          =   285
            Left            =   960
            TabIndex        =   14
            Top             =   465
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   3
            Mask            =   "999"
            PromptChar      =   "_"
         End
         Begin VB.Label LblId 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   0
            Left            =   225
            MouseIcon       =   "Seg.frx":07A8
            MousePointer    =   99  'Custom
            TabIndex        =   13
            Top             =   465
            Width           =   570
         End
      End
   End
   Begin Threed.SSCommand CmdSeg 
      Height          =   435
      Index           =   0
      Left            =   900
      TabIndex        =   28
      Top             =   4650
      Width           =   1035
      _Version        =   65536
      _ExtentX        =   1826
      _ExtentY        =   767
      _StockProps     =   78
      Caption         =   "&Inserir"
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
   Begin Threed.SSCommand CmdSeg 
      Height          =   435
      Index           =   1
      Left            =   3120
      TabIndex        =   29
      Top             =   4650
      Width           =   1035
      _Version        =   65536
      _ExtentX        =   1826
      _ExtentY        =   767
      _StockProps     =   78
      Caption         =   "&Excluir"
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
   Begin Threed.SSCommand CmdSeg 
      Height          =   435
      Index           =   2
      Left            =   5460
      TabIndex        =   30
      Top             =   4650
      Width           =   1035
      _Version        =   65536
      _ExtentX        =   1826
      _ExtentY        =   767
      _StockProps     =   78
      Caption         =   "&Sair"
      ForeColor       =   0
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
   Begin VB.Label LblOper 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   360
      TabIndex        =   32
      Top             =   4560
      Width           =   6675
   End
End
Attribute VB_Name = "FrmSeg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Event Load()
Event UnLoad(Cancel As Integer)
Event TabSegClick(PreviousTab As Integer)
Event ChkAcessoClick(Index As Integer)
Event CmdSegClick(Index As Integer)
Event LblIdClick(Index As Integer)
Event CmbGrpUserClick(Index As Integer)
Event TreModuNodeClick(ByVal Node As ComctlLib.Node)
Event TxtIdUserLostFocus()
Event MskIdGrpUserLostFocus()

Public Suja%
Public Function ValidaCampos()
   ValidaCampos = True
   Select Case Me.TabSeg.Tab
      Case 0
         If Val(Me.MskIdGrpUser) = 0 Or Trim$(Me.TxtDscGrpUser) = "" Then
            Call ClsMsg.ExibirAviso(ClsMsg.LoadMsg(27), ClsMsg.LoadMsg(1))
            If Val(Me.MskIdGrpUser) = 0 Then Call ClsCtrl.Set_Focus(Me.MskIdGrpUser)
            If Trim$(Me.MskIdGrpUser) = "" Then Call ClsCtrl.Set_Focus(Me.TxtDscGrpUser)
            Exit Function
         End If
      Case 1
         If Trim(Me.TxtIdUser) = "" Or Trim$(Me.TxtNmUser) = "" Or Trim$(Me.TxtPwdUser(0)) = "" Then
            Call ClsMsg.ExibirAviso(ClsMsg.LoadMsg(27), ClsMsg.LoadMsg(1))
            If Trim$(Me.TxtIdUser) = "" Then Call ClsCtrl.Set_Focus(Me.TxtIdUser)
            If Trim$(Me.TxtNmUser) = "" Then Call ClsCtrl.Set_Focus(Me.TxtNmUser)
            If Trim$(Me.TxtPwdUser(0)) = "" Then Call ClsCtrl.Set_Focus(Me.TxtPwdUser(0))
            Exit Function
         End If
   End Select
   ValidaCampos = True
End Function
Public Sub F_EXCLUIR()
   Call CmdSeg_Click(1)
End Sub
Public Sub F_PROCURAR()
   Select Case Me.TabSeg.Tab
      Case 0: Call LblId_Click(0)
      Case 1: Call LblId_Click(1)
   End Select
End Sub
Public Sub F_SALVAR()
   Call CmdSeg_Click(0)
End Sub
Private Sub ChkAcesso_Click(Index As Integer)
   RaiseEvent ChkAcessoClick(Index)
End Sub
Private Sub CmbGrpUser_Click(Index As Integer)
   RaiseEvent CmbGrpUserClick(Index)
End Sub
Private Sub CmbGrpUser_LostFocus(Index As Integer)
'   If Screen.ActiveForm.Name <> FrmSeg.Name Then Exit Sub
   Select Case Index
      Case 0: ' FrmSeg.CmdSeg(0).SetFocus
   End Select
End Sub
Private Sub CmdSeg_Click(Index As Integer)
   RaiseEvent CmdSegClick(Index)
End Sub

Private Sub Form_Activate()
   Screen.MousePointer = vbDefault
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
      Case vbKeyEscape: UnLoad Me
      Case Else: KeyAscii = ClsDsr.SendTab(Me, KeyAscii)
   End Select
   DoEvents
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF2: Call LblId_Click(Me.TabSeg.Tab)
   End Select
   DoEvents
End Sub

Private Sub Form_Load()
   RaiseEvent Load
End Sub
Private Sub Form_Unload(Cancel As Integer)
   RaiseEvent UnLoad(Cancel)
End Sub
Private Sub LblId_Click(Index As Integer)
   RaiseEvent LblIdClick(Index)
End Sub
Private Sub MskIdGrpUser_GotFocus()
   Call ClsDsr.SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub MskIdGrpUser_LostFocus()
   RaiseEvent MskIdGrpUserLostFocus
End Sub
Private Sub TabSeg_Click(PreviousTab As Integer)
   RaiseEvent TabSegClick(PreviousTab)
End Sub
Private Sub TreModu_DblClick()
   Dim i%
   For i = 0 To Me.ChkAcesso.UBound
      Me.ChkAcesso(i).Value = vbChecked
   Next
End Sub
Private Sub TreModu_NodeClick(ByVal Node As ComctlLib.Node)
   RaiseEvent TreModuNodeClick(Node)
End Sub
Private Sub TxtDscGrpUser_GotFocus()
   Call ClsDsr.SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtDscGrpUser_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then CmdSeg(0).SetFocus
End Sub

Private Sub TxtIdUser_GotFocus()
   Call ClsDsr.SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtIdUser_LostFocus()
   RaiseEvent TxtIdUserLostFocus
End Sub
Private Sub TxtNmUser_GotFocus()
   Call ClsDsr.SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtPwdUser_GotFocus(Index As Integer)
   Call ClsDsr.SelecionarTexto(Me.ActiveControl)
End Sub
Private Sub TxtPwdUser_LostFocus(Index As Integer)
   Select Case Index
    Case 1
       If TxtPwdUser(1) = "" Then Exit Sub
       If Trim$(TxtPwdUser(1)) <> Trim$(TxtPwdUser(0)) Then
          Call ClsMsg.ExibirStop(ClsMsg.LoadMsg(59)) '"Senha inválida."
          TxtPwdUser(Index) = ""
          TxtPwdUser(Index).SetFocus
       End If
   End Select
End Sub
