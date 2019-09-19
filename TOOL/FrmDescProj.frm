VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmDescProj 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Label1"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   2100
   ClientWidth     =   9675
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6645
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PicMoveDiv 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   5040
      Left            =   4440
      ScaleHeight     =   2194.633
      ScaleMode       =   0  'User
      ScaleWidth      =   156
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.TextBox TxtProj 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   330
      Left            =   840
      LinkTimeout     =   30
      TabIndex        =   4
      Top             =   120
      Width           =   8115
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
      Height          =   330
      Left            =   9120
      TabIndex        =   3
      Top             =   120
      Width           =   330
   End
   Begin ComctlLib.TreeView TreProj 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   4150
      _ExtentX        =   7329
      _ExtentY        =   9128
      _Version        =   327682
      Indentation     =   529
      LineStyle       =   1
      PathSeparator   =   " / "
      Style           =   7
      ImageList       =   "ImgList"
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand CmdOper 
      Height          =   405
      Index           =   0
      Left            =   8640
      TabIndex        =   1
      Top             =   6120
      Width           =   945
      _Version        =   65536
      _ExtentX        =   1667
      _ExtentY        =   714
      _StockProps     =   78
      ForeColor       =   -2147483637
      Picture         =   "FrmDescProj.frx":0000
   End
   Begin Threed.SSCommand CmdOper 
      Height          =   405
      Index           =   1
      Left            =   7680
      TabIndex        =   2
      Top             =   6120
      Width           =   945
      _Version        =   65536
      _ExtentX        =   1667
      _ExtentY        =   714
      _StockProps     =   78
      ForeColor       =   12632256
      Font3D          =   4
      Picture         =   "FrmDescProj.frx":120A
   End
   Begin ComctlLib.ListView lvListView 
      Height          =   5400
      Left            =   4200
      TabIndex        =   12
      Top             =   600
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   9525
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblTitle 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " TreeView:"
      Height          =   270
      Index           =   0
      Left            =   0
      TabIndex        =   13
      Tag             =   " TreeView:"
      Top             =   600
      Width           =   4150
   End
   Begin VB.Image ImgDivisao 
      Height          =   5145
      Left            =   4080
      MousePointer    =   9  'Size W E
      Top             =   840
      Width           =   200
   End
   Begin VB.Label LblProp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   4560
      TabIndex        =   10
      Top             =   1800
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label LblProp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   4560
      TabIndex        =   9
      Top             =   1440
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label LblProp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   4560
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label LblProp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   4560
      TabIndex        =   7
      Top             =   720
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label LblProp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   4560
      TabIndex        =   6
      Top             =   360
      Visible         =   0   'False
      Width           =   915
   End
   Begin ComctlLib.ImageList ImgList 
      Left            =   0
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   9
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmDescProj.frx":1C44
            Key             =   "CLASSE"
            Object.Tag             =   "CLASSE"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmDescProj.frx":1DC6
            Key             =   "DLL"
            Object.Tag             =   "DLL"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmDescProj.frx":1F48
            Key             =   "FORM"
            Object.Tag             =   "FORM"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmDescProj.frx":20CA
            Key             =   "FORMCHILD"
            Object.Tag             =   "FORMCHILD"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmDescProj.frx":224C
            Key             =   "MODULO"
            Object.Tag             =   "MODULO"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmDescProj.frx":23CE
            Key             =   "PROJ"
            Object.Tag             =   "PROJ"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmDescProj.frx":2550
            Key             =   "RES"
            Object.Tag             =   "RES"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmDescProj.frx":26D2
            Key             =   "PASTA"
            Object.Tag             =   "PASTA"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmDescProj.frx":2854
            Key             =   "PASTAA"
            Object.Tag             =   "PASTAA"
         EndProperty
      EndProperty
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Projeto"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   600
   End
End
Attribute VB_Name = "FrmDescProj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Acesso$
Public Suja As Boolean
Public FormDesc As New DESCPROJ
Public FormProj As New PROJETO
Dim fMoving As Boolean
Public Sub F_SALVAR()
End Sub

Private Sub CmdDrv_Click()
   Dim PATH$
   Dim Tit$, Filtro$, Arq$, Ind%
   Set FormDesc = Nothing
   Set FormProj = Nothing
   Tit$ = "Localizar Projeto"
   Filtro = "Project Files (*.vbp; *.vbg; *.mak)|*.vbp; *.vbg; *.mak"
   Ind% = 1
   If Trim(Me.TxtProj.Tag) = "" Then
      SysMdi.CmDialog.InitDir = "C:\SISTEMAS\"
   Else
      SysMdi.CmDialog.InitDir = Me.TxtProj.Tag
   End If
   Arq$ = ProcurarArquivo(SysMdi.CmDialog, Tit$, Arq$, Filtro$, Ind%)
   If Trim(Arq) <> "" Then
      Me.TxtProj.Text = UCase(SysMdi.CmDialog.Tag) & Arq
      Me.TxtProj.Tag = UCase(SysMdi.CmDialog.Tag)
      Set FormDesc = Nothing
      If Me.TxtProj = "" Then Call CmdDrv_Click
      With FormDesc
         .FileProj = Mid(Me.TxtProj, Len(Me.TxtProj.Tag) + 1)
         .PathProj = Me.TxtProj.Tag
         Set .fMe = Me
         Call SetHourglass(hWnd)
         Call .CarregaProjeto(.PathProj, .FileProj)
         Me.TreProj.Nodes.Clear
         Call .MontaTreeProjeto
         Call SetDefault(hWnd)
      End With
   End If
End Sub

Private Sub CmdOper_Click(Index As Integer)
   Dim Sql$
   Select Case Index
      Case 0:  If Me.Suja Then F_SALVAR Else Unload Me
      Case 1: Unload Me
   End Select
End Sub
Private Sub Form_Load()
   Dim i%
   Call SetHourglass(hWnd)
   Set MDIFilho = Me
   Call ConfigForm(Me, SysMdi.Icon, FundoTela)
   For i = Me.LblProp.LBound To Me.LblProp.UBound
      Me.LblProp(i) = ""
   Next
   Call SetDefault(hWnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set FormDesc = Nothing
   Set FormProj = Nothing
   Set MDIFilho = Nothing
End Sub

Private Sub Lbl_Click(Index As Integer)
   Select Case Index
      Case 2: Call CmdDrv_Click
   End Select
End Sub

Private Sub TreProj_Collapse(ByVal Node As ComctlLib.Node)
   If Node.Image = "PASTAA" Then Node.Image = "PASTA"
End Sub
Private Sub TreProj_Expand(ByVal Node As ComctlLib.Node)
   If Node.Image = "PASTA" Then Node.Image = "PASTAA"
End Sub

Private Sub TreProj_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 And TreProj.Nodes.Count > 0 Then PopupMenu MdiPrincipal.MnuMouse(0)
End Sub

Private Sub TreProj_NodeClick(ByVal Node As ComctlLib.Node)
   Dim Pai$, Filho$
   Dim No As New Node
   Dim MyProj As New PROJETO
   Dim MyBas As New Modulo
   Dim MyForm As New FORMULARIO
   Dim MyClass As New CLASSE
   Dim MyFunc As New FUNCAO
   Dim i%
   On Error GoTo Fim
   For i = Me.LblProp.LBound To Me.LblProp.UBound
      Me.LblProp(i) = ""
   Next
   Set No = Node
   If UCase(Node) = "SISTEMA" Then
      Exit Sub
   End If
   Pai = UCase(Node.Parent)
   
   While UCase(No.Parent.Key) <> "ROOT"
      Set No = No.Parent
   Wend
   Set MyProj = FormDesc.PROJETOS(No.Key)
   Select Case Pai
      Case "SISTEMA" 'Projeto
         Me.LblProp(0) = "Nome       : " & MyProj.NOME
         Me.LblProp(1) = "Descrição  : " & MyProj.DESCRIÇÃO
         Me.LblProp(2) = "Localização: " & MyProj.PATH
         Me.LblProp(3) = "Arquivo    : " & MyProj.FILENAME
         Me.LblProp(4) = "Linhas     : " & MyProj.LINHAS
      Case "MODULOS"
         Chave = UCase(Mid(Node.Key, 1, Len(Node.Key) - 1))
         Set MyBas = MyProj.MODULOS(Chave)
         Me.LblProp(0) = "Nome       : " & MyBas.NOME
         'Me.LblProp(1) = "Descrição  : " & MyBas.DESCRIÇÃO
         Me.LblProp(1) = "Localização: " & MyBas.PATH
         Me.LblProp(2) = "Arquivo    : " & MyBas.FILENAME
         Me.LblProp(3) = "Linhas     : " & MyBas.LINHAS
   End Select
   For i = Me.LblProp.LBound To Me.LblProp.UBound
      Me.LblProp(i).Visible = (Len(Me.LblProp(i)) <> 13)
   Next
Fim:
   Call ShowError("FrmdescProj.NodeClick")
End Sub
Private Sub ImgDivisao_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   PicMoveDiv.Move ImgDivisao.Left, ImgDivisao.Top, PicMoveDiv.Width, ImgDivisao.Height - 20
   PicMoveDiv.Visible = True
   fMoving = True
End Sub
Private Sub ImgDivisao_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If fMoving Then PicMoveDiv.Left = x + ImgDivisao.Left
End Sub
Private Sub ImgDivisao_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call SizeControls(PicMoveDiv.Left)
    PicMoveDiv.Visible = False
    fMoving = False
End Sub
Public Sub SizeControls(Pos As Single)
   Dim Tree As TreeView, Lst As ListView
   Const Limite% = 1500
   Const Separa% = 50
 
   Set Tree = TreProj
   Set Lst = lvListView
   
   On Error Resume Next
   Select Case Pos
      Case Is < Limite: Pos = Limite
      Case Is > Me.Width - Limite: Pos = Me.Width - Limite
   End Select
   Tree.Move Tree.Left, Tree.Top, Pos
   Lst.Move Pos + Separa, Lst.Top, Me.Width - (Tree.Width + 140), Tree.Height
   ImgDivisao.Move Pos, Tree.Top, ImgDivisao.Width, ImgDivisao.Height
Lst.Top = lblTitle(0).Top
Lst.Height = Lst.Height + lblTitle(0).Height
   lblTitle(0).Width = Tree.Width
'   lblTitle(1).Left = Lst.Left + 50
'   lblTitle(1).Width = Lst.Width - 40
End Sub

