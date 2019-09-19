VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmAddIn 
   AutoRedraw      =   -1  'True
   ClientHeight    =   6225
   ClientLeft      =   1245
   ClientTop       =   1500
   ClientWidth     =   10545
   Icon            =   "FrmAddIn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "FrmAddIn.frx":058A
   ScaleHeight     =   6225
   ScaleWidth      =   10545
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   873
      ButtonWidth     =   767
      ButtonHeight    =   714
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImgLstToolbar"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   5
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "SAIR"
            Object.ToolTipText     =   "Sair do Sistema"
            Object.Tag             =   ""
            ImageKey        =   "EXIT"
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "MONTAR CLASSE"
            Object.ToolTipText     =   "Montar Classe de Banco"
            Object.Tag             =   ""
            ImageKey        =   "CLSWIZARD"
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "PROPRIEDADES"
            Object.ToolTipText     =   "Propriedades"
            Object.Tag             =   ""
            ImageKey        =   "PROP"
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   10
      Left            =   2520
      Top             =   5280
   End
   Begin ComctlLib.ListView LstItens 
      Height          =   4605
      Left            =   3720
      TabIndex        =   4
      Top             =   3120
      WhatsThisHelpID =   10464
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   8123
      View            =   3
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   15925247
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Descrição"
         Object.Width           =   8705
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   2540
      EndProperty
   End
   Begin TabDlg.SSTab TabComp 
      Height          =   5325
      Left            =   3360
      TabIndex        =   2
      Top             =   480
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   9393
      _Version        =   393216
      Tabs            =   6
      Tab             =   5
      TabsPerRow      =   6
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Propriedades"
      TabPicture(0)   =   "FrmAddIn.frx":08CC
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "&Métodos"
      TabPicture(1)   =   "FrmAddIn.frx":08E8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "&Eventos"
      TabPicture(2)   =   "FrmAddIn.frx":0904
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "&Var/Const"
      TabPicture(3)   =   "FrmAddIn.frx":0920
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "&Tudo"
      TabPicture(4)   =   "FrmAddIn.frx":093C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "&Código"
      TabPicture(5)   =   "FrmAddIn.frx":0958
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "TxtCode"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "CmbControl"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "CmbMember"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "RtfTemp"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).ControlCount=   4
      Begin RichTextLib.RichTextBox RtfTemp 
         Height          =   495
         Left            =   3720
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   840
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         _Version        =   393217
         BorderStyle     =   0
         ScrollBars      =   3
         BulletIndent    =   3
         RightMargin     =   60000
         TextRTF         =   $"FrmAddIn.frx":0974
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ComboBox CmbMember 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3720
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   480
         Width           =   3255
      End
      Begin VB.ComboBox CmbControl 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "FrmAddIn.frx":0A62
         Left            =   120
         List            =   "FrmAddIn.frx":0A64
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   480
         Width           =   3375
      End
      Begin RichTextLib.RichTextBox TxtCode 
         Height          =   4815
         Left            =   120
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   360
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   8493
         _Version        =   393217
         BackColor       =   15925247
         BorderStyle     =   0
         ScrollBars      =   3
         BulletIndent    =   3
         RightMargin     =   60000
         TextRTF         =   $"FrmAddIn.frx":0A66
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComDlg.CommonDialog CmDialog 
      Left            =   1920
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.TreeView TreProj 
      Height          =   4845
      Left            =   60
      TabIndex        =   1
      Top             =   960
      WhatsThisHelpID =   10234
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   8546
      _Version        =   327682
      HideSelection   =   0   'False
      Indentation     =   529
      LineStyle       =   1
      PathSeparator   =   " / "
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ImgLstIcons"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Spl 
      Height          =   4815
      Left            =   1080
      ScaleHeight     =   4755
      ScaleWidth      =   7875
      TabIndex        =   9
      Top             =   960
      Width           =   7935
   End
   Begin VB.Label Lbl 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Projeto : "
      Height          =   495
      Left            =   60
      TabIndex        =   3
      Top             =   480
      Width           =   3225
   End
   Begin ComctlLib.ImageList ImgLstToolbar 
      Left            =   840
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   22
      ImageHeight     =   21
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAddIn.frx":0B54
            Key             =   "EXIT"
            Object.Tag             =   "EXIT"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAddIn.frx":119E
            Key             =   "CLSWIZARD"
            Object.Tag             =   "CLSWIZARD"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmAddIn.frx":17E8
            Key             =   "PROP"
            Object.Tag             =   "PROP"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImgLstIcons 
      Left            =   120
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   327682
   End
   Begin VB.Menu Mnu00 
      Caption         =   "Arquivo"
      Index           =   0
      Begin VB.Menu Mnu0000 
         Caption         =   "&Abrir Projeto"
         Index           =   0
      End
      Begin VB.Menu Mnu0000 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu Mnu0000 
         Caption         =   "Opções"
         Index           =   2
      End
      Begin VB.Menu Mnu0000 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu Mnu0000 
         Caption         =   "&Sair"
         Index           =   4
      End
   End
   Begin VB.Menu MnuMouse_Main 
      Caption         =   "Menu Mouse"
      Begin VB.Menu MnuMouse 
         Caption         =   "TreeView (MnuMouse)"
         Index           =   0
         Begin VB.Menu MnuMouse00 
            Caption         =   "&Adicionar"
            Index           =   0
         End
         Begin VB.Menu MnuMouse00 
            Caption         =   "&Adicionar..."
            Index           =   1
            Begin VB.Menu MnuMouse0000 
               Caption         =   "Classe de Banco"
               Index           =   0
            End
            Begin VB.Menu MnuMouse0000 
               Caption         =   "Classe"
               Index           =   1
            End
            Begin VB.Menu MnuMouse0000 
               Caption         =   "Coleção"
               Index           =   2
            End
            Begin VB.Menu MnuMouse0000 
               Caption         =   "Propriedade,Variável"
               Index           =   3
            End
            Begin VB.Menu MnuMouse0000 
               Caption         =   "Método"
               Index           =   4
            End
            Begin VB.Menu MnuMouse0000 
               Caption         =   "Evento"
               Index           =   5
            End
         End
         Begin VB.Menu MnuMouse00 
            Caption         =   "Salvar"
            Index           =   2
         End
         Begin VB.Menu MnuMouse00 
            Caption         =   "Salvar Como..."
            Index           =   3
         End
         Begin VB.Menu MnuMouse00 
            Caption         =   "Recarregar &Objeto"
            Index           =   4
         End
         Begin VB.Menu MnuMouse00 
            Caption         =   "-"
            Index           =   5
         End
         Begin VB.Menu MnuMouse00 
            Caption         =   "Cortar"
            Index           =   6
         End
         Begin VB.Menu MnuMouse00 
            Caption         =   "Copiar"
            Index           =   7
         End
         Begin VB.Menu MnuMouse00 
            Caption         =   "Colar"
            Index           =   8
         End
         Begin VB.Menu MnuMouse00 
            Caption         =   "-"
            Index           =   9
         End
         Begin VB.Menu MnuMouse00 
            Caption         =   "Excluir"
            Index           =   10
         End
         Begin VB.Menu MnuMouse00 
            Caption         =   "&Renomear"
            Index           =   11
         End
         Begin VB.Menu MnuMouse00 
            Caption         =   "-"
            Index           =   12
         End
         Begin VB.Menu MnuMouse00 
            Caption         =   "Propriedades..."
            Index           =   13
         End
         Begin VB.Menu MnuMouse00 
            Caption         =   "Definir Como "
            Index           =   14
         End
         Begin VB.Menu MnuMouse00 
            Caption         =   "-"
            Index           =   15
         End
         Begin VB.Menu MnuMouse00 
            Caption         =   "&Ir Para Projeto"
            Index           =   16
         End
      End
   End
End
Attribute VB_Name = "FrmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Load()
Event Active()
Event Resize()
Event KeyUp(KeyCode As Integer, Shift As Integer)
   
Event CmbControlClick()
Event CmbMemberClick()

Event LstItensClick()
Event LstItensDblClick()
Event LstItensMouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event LstItensAfterLabelEdit(Cancel As Integer, NewString As String)

Event MnuMouseClick(Menu As String, Index As Integer)
Event MenuClick(Menu As String, Index As Integer)

Event ToolbarButtonClick(ByVal Button As ComctlLib.Button)
Event TabCompClick(PreviousTab As Integer)

Event TreProjMouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event TreProjMouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event TreProjExpand(ByVal Node As ComctlLib.Node)
Event TreProjNodeClick(Node)
Event TreProjBeforeLabelEdit(Cancel As Integer)
Event TreProjAfterLabelEdit(Cancel As Integer, NewString As String)
Event TxtCodeClick(IsMouseClick As Boolean)
Event TreProjDblClick()
Event TxtCodeKeyUp(KeyCode As Integer, Shift As Integer)
Event TxtCodeMouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event TxtCodeMouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

Public PrimeiraVez As Boolean
Public Sub PrintFlood(Str As String)
  CurrentY = 0
  CurrentX = 40
'  For i = i To 40
'     CurrentX = i
     Me.Print Str
'  Next
  Me.Refresh
End Sub
Private Sub CmbControl_Click()
   RaiseEvent CmbControlClick
End Sub
Private Sub CmbMember_Click()
   RaiseEvent CmbMemberClick
End Sub

'Public VBInstance As VBIDE.VBE
'Public Connect As Connect
'Public fComp As VBComponent
'Public fMember As Member
'Public fCode As CodeModule
Private Sub Form_Activate()
   RaiseEvent Active
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub Form_Load()
   Me.TxtCode.ZOrder 0
   RaiseEvent Load
End Sub
Private Sub Form_Resize()
   RaiseEvent Resize
End Sub

Private Sub LstItens_AfterLabelEdit(Cancel As Integer, NewString As String)
   RaiseEvent LstItensAfterLabelEdit(Cancel, NewString)
End Sub

Private Sub LstItens_Click()
   RaiseEvent LstItensClick
End Sub
Private Sub LstItens_DblClick()
   RaiseEvent LstItensDblClick
End Sub
Private Sub LstItens_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
   RaiseEvent LstItensMouseUp(Button, Shift, x, Y)
End Sub

Private Sub Mnu0000_Click(Index As Integer)
   RaiseEvent MenuClick("0000", Index)
End Sub

Private Sub MnuMouse00_Click(Index As Integer)
   RaiseEvent MnuMouseClick("00", Index)
End Sub
Private Sub MnuMouse0000_Click(Index As Integer)
   RaiseEvent MnuMouseClick("0000", Index)
End Sub
Private Sub Spl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
   With Me
'      Call .Spl.SizeControls(.Spl.Left, .TreProj, .TabComp, .Lbl)
      .LstItens.Move .TabComp.Left + 60, .LstItens.Top, .TabComp.Width - 120, .TabComp.Height - 480
      .TxtCode.Move 60, 400, .TabComp.Width - 120, .TabComp.Height - 480
      
      .TabComp.Move .TabComp.Left, .TabComp.Top, .Width - (.TabComp.Left + 180), .Height - .TabComp.Top - 960
   End With

End Sub
Private Sub TabComp_Click(PreviousTab As Integer)
   RaiseEvent TabCompClick(PreviousTab)
End Sub
Private Sub Timer_Timer(Index As Integer)
   Select Case Index
      Case 0
         Dim Pnt As PointAPI
         GetCursorPos Pnt
         If Pnt.x >= Me.Left / Screen.TwipsPerPixelX And Pnt.x <= (Me.Left + Me.ScaleWidth) / Screen.TwipsPerPixelX + 7 And _
            Pnt.Y >= Me.Top / Screen.TwipsPerPixelY And Pnt.Y <= (Me.Top + Me.ScaleHeight) / Screen.TwipsPerPixelY + 23 Then
            '* No Formulário
            Me.TxtCode.Enabled = True
            Me.Timer(0).Enabled = False
         End If

   End Select
   
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As ComctlLib.Button)
   RaiseEvent ToolbarButtonClick(Button)
End Sub
Private Sub TreProj_AfterLabelEdit(Cancel As Integer, NewString As String)
   RaiseEvent TreProjAfterLabelEdit(Cancel, NewString)
End Sub

Private Sub TreProj_BeforeLabelEdit(Cancel As Integer)
   RaiseEvent TreProjBeforeLabelEdit(Cancel)
End Sub

Private Sub TreProj_DblClick()
   RaiseEvent TreProjDblClick
End Sub

Private Sub TreProj_Expand(ByVal Node As ComctlLib.Node)
   RaiseEvent TreProjExpand(Node)
End Sub

Private Sub TreProj_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   RaiseEvent TreProjMouseDown(Button, Shift, x, Y)
End Sub

Private Sub TreProj_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
   RaiseEvent TreProjMouseUp(Button, Shift, x, Y)
End Sub

Private Sub TreProj_NodeClick(ByVal Node As ComctlLib.Node)
   RaiseEvent TreProjNodeClick(Node)
End Sub
Private Sub TxtCode_Click()
   If Me.TxtCode.Enabled Then
      RaiseEvent TxtCodeClick(False)
   End If
End Sub
Private Sub TxtCode_KeyPress(KeyAscii As Integer)
   'Call DoAutocomplete(Me.TxtCode)
End Sub

Private Sub TxtCode_KeyUp(KeyCode As Integer, Shift As Integer)
   RaiseEvent TxtCodeKeyUp(KeyCode, Shift)
End Sub

Private Sub TxtCode_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   RaiseEvent TxtCodeMouseDown(Button, Shift, x, Y)
End Sub

Private Sub TxtCode_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
   RaiseEvent TxtCodeMouseUp(Button, Shift, x, Y)
End Sub

'Public Sub MontaLstItens(TabAtual As Integer, Optional LimpaList = True)
'   Dim i%, lKey$, lPai$, lFilho$, lICON$
'   Dim LinIni&, QtdLin&
'   Dim myNo As Node
'   Dim ItemX As ListItem
'   Dim StrLinha$, PosTexto&, PosAnt&
'
''   On Error Resume Next
'   '* Definir Chave do Componente
'   If Not fComp Is Nothing Then lKey$ = fComp.Name
'
'   With Me.LstItens
'      If LimpaList Then .ListItems.Clear
'      Select Case TabAtual
'         Case 0 '* Propriedades
'            lPai$ = lKey$ & "PROPRIEDADE"
'            lICON$ = "PROPRIEDADE"
'         Case 1 '* Métodos
'            .ListItems.Clear
'            lPai$ = lKey$ & "METODO"
'            lICON$ = "METODO"
'         Case 2 '* Eventos
'            lPai$ = lKey$ & "EVENTO"
'            lICON$ = "EVENTO"
'         Case 3 '* Variáveis e Constantes
'            lPai$ = lKey$ & "VARIAVEL"
'            lICON$ = "VARIAVEL"
'         Case 4 '* Tudo
'            For i = 0 To Me.TabComp.Tabs - 1
'               If Me.TabComp.TabVisible(i) And i <> 4 And i <> 5 Then
'                  Call MontaLstItens(i, False)
'               End If
'            Next
'            lPai$ = "" 'lKEY$ & "TUDO"
'         Case 5 '* Código
'            Call ExibirCodigo
'      End Select
'      If ExisteNo(Me.TreProj, lPai$) Then
'         For i = 1 To Me.TreProj.Nodes(lPai$).Children
'            lFilho$ = lPai$ & Trim(CStr(i))
'             lKey = Me.TreProj.Nodes(lFilho$).Text
'            Set ItemX = Me.LstItens.ListItems.ADD(, lKey, Me.TreProj.Nodes(lFilho$).Text, lICON$, lICON$)
''****************************************
'            Set fMember = fComp.CodeModule.Members.Item(Me.TreProj.Nodes(lFilho$).Text)
'            'lTipo = CStr(MyComp.CodeModule.Members.Item(i).Type)
'            Select Case fMember.Scope
'               Case 1: ItemX.SubItems(1) = "'Private'"
'               Case 2: ItemX.SubItems(1) = "'Public'"
'               Case 3: ItemX.SubItems(1) = "'Friend'"
'               Case Else: ItemX.SubItems(1) = "'Outro'"
'            End Select
'            ItemX.SubItems(1) = ItemX.SubItems(1) & ", " & fMember.Category
''****************************************
'         Next
'         If VBA.Right(lPai$, 8) = "VARIAVEL" Then
'            lPai$ = Mid(lPai$, 1, Len(lPai) - 8) & "CONSTANTE"
'            lICON$ = "CONSTANTE"
'            If ExisteNo(Me.TreProj, lPai$) Then
'               For i = 1 To Me.TreProj.Nodes(lPai$).Children
'                  lFilho$ = lPai$ & Trim(CStr(i))
'                  lKey =  Me.TreProj.Nodes(lFilho$).Text
'                  Set ItemX = Me.LstItens.ListItems.ADD(, lKey, Me.TreProj.Nodes(lFilho$).Text, lICON$, lICON$)
'               Next
'            End If
'         End If
'      End If
'   End With
'Fim:
'End Sub
