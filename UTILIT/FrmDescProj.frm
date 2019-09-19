VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form FrmDescProj 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Label1"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   2100
   ClientWidth     =   11355
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6465
   ScaleWidth      =   11355
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtProj 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   330
      Left            =   240
      LinkTimeout     =   30
      TabIndex        =   4
      Top             =   240
      Width           =   3675
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
      Left            =   3960
      TabIndex        =   3
      Top             =   240
      Width           =   435
   End
   Begin ComctlLib.TreeView TreProj 
      Height          =   5775
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   10186
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
      Left            =   8880
      TabIndex        =   1
      Top             =   4680
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
      Left            =   7920
      TabIndex        =   2
      Top             =   4680
      Width           =   945
      _Version        =   65536
      _ExtentX        =   1667
      _ExtentY        =   714
      _StockProps     =   78
      ForeColor       =   12632256
      Font3D          =   4
      Picture         =   "FrmDescProj.frx":120A
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
      Left            =   240
      TabIndex        =   5
      Top             =   0
      Width           =   600
   End
   Begin VB.Image ImgFundo 
      BorderStyle     =   1  'Fixed Single
      Height          =   990
      Left            =   0
      Top             =   120
      Width           =   990
   End
End
Attribute VB_Name = "FrmDescProj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Acesso$
Public Suja As Boolean
Dim FormDesc As Object
Public Sub F_SALVAR()
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
   Call SetHourglass(hwnd)
   Set Sys.MDIFilho = Me
   Call ConfigForm(Me, SysMdi.Icon, Sys.FundoTela)
   For i = Me.LblProp.LBound To Me.LblProp.UBound
      Me.LblProp(i) = ""
   Next
   Call SetDefault(hwnd)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set FormDesc = Nothing
'   Set FormProj = Nothing
   Set Sys.MDIFilho = Nothing
End Sub

Private Sub TreProj_Collapse(ByVal Node As ComctlLib.Node)
   If Node.Image = "PASTAA" Then Node.Image = "PASTA"
End Sub

Private Sub TreProj_Expand(ByVal Node As ComctlLib.Node)
   If Node.Image = "PASTA" Then Node.Image = "PASTAA"
End Sub

Private Sub TreProj_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If x = TreProj.Width - 75 Then
'      Me.TreProj.MouseIcon = LoadResPicture("100", vbResCursor)
'      Screen.MousePointer = LoadResPicture("100", vbResIcon)
   End If
End Sub
Private Sub TreProj_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 And TreProj.Nodes.Count > 0 Then PopupMenu MdiPrincipal.MnuMouse(0)
End Sub

Private Sub TreProj_NodeClick(ByVal Node As ComctlLib.Node)
   Dim Pai$, Filho$
   Dim No As New Node
   Dim MyBas As Object
   Dim MyProj As Object
   Dim Chave$
   
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
         Me.LblProp(2) = "Localização: " & MyProj.Path
         Me.LblProp(3) = "Arquivo    : " & MyProj.filename
         Me.LblProp(4) = "Linhas     : " & MyProj.LINHAS
      Case "MODULOS"
         Chave = UCase(Mid(Node.Key, 1, Len(Node.Key) - 1))
         Set MyBas = MyProj.MODULOS(Chave)
         Me.LblProp(0) = "Nome       : " & MyBas.NOME
         'Me.LblProp(1) = "Descrição  : " & MyBas.DESCRIÇÃO
         Me.LblProp(1) = "Localização: " & MyBas.Path
         Me.LblProp(2) = "Arquivo    : " & MyBas.filename
         Me.LblProp(3) = "Linhas     : " & MyBas.LINHAS
   End Select
   For i = Me.LblProp.LBound To Me.LblProp.UBound
      Me.LblProp(i).Visible = (Len(Me.LblProp(i)) <> 13)
   Next
Fim:
   Call ShowError("FrmdescProj.NodeClick")
End Sub
