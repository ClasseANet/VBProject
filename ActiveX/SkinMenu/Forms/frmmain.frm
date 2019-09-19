VERSION 5.00
Object = "*\A..\Projects\SkinMenu.vbp"
Begin VB.Form frmmain 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demonistration Plz Vote"
   ClientHeight    =   5415
   ClientLeft      =   4125
   ClientTop       =   3360
   ClientWidth     =   8535
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   8535
   Begin VB.CommandButton cmdapply 
      Caption         =   "Apply Settings"
      Height          =   495
      Left            =   5400
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "frmmain.frx":0000
      Left            =   5400
      List            =   "frmmain.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1080
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmmain.frx":0079
      Left            =   5400
      List            =   "frmmain.frx":0089
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   720
      Width           =   2655
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   360
      Picture         =   "frmmain.frx":00CB
      ScaleHeight     =   4425
      ScaleWidth      =   2985
      TabIndex        =   7
      Top             =   3600
      Width           =   3015
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "XpStyle Skinize menu"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   9
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   $"frmmain.frx":0981
         Height          =   1335
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   2775
      End
   End
   Begin SkinMenu.sMenu sMenu1 
      Height          =   975
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1720
      Style           =   1
      HighLightStyle  =   0
      ItemCount       =   8
      BeginProperty ItemsFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MenuCaption1    =   "Arquivo"
      MenuName1       =   "Arquivo"
      MenuIdent2      =   1
      MenuCaption2    =   "Novo"
      MenuName2       =   "Novo"
      MenuIdent3      =   1
      MenuCaption3    =   "Salvar"
      MenuName3       =   "Salvar"
      MenuIdent4      =   1
      MenuCaption4    =   "-"
      MenuName4       =   "Sep1"
      MenuIdent5      =   1
      MenuCaption5    =   "Sair"
      MenuName5       =   "Sair"
      MenuCaption6    =   "Edição"
      MenuName6       =   "Edição"
      MenuIdent7      =   1
      MenuCaption7    =   "Cortar"
      MenuName7       =   "Cortar"
      MenuIdent8      =   1
      MenuCaption8    =   "Colar"
      MenuName8       =   "Colar"
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Highlight Style:"
      Height          =   195
      Left            =   3600
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmmain.frx":0A72
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   3600
      TabIndex        =   5
      Top             =   2280
      Width           =   4815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   0
      Width           =   5055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Appaerance:"
      Height          =   195
      Left            =   3600
      TabIndex        =   2
      Top             =   780
      Width           =   1515
   End
   Begin VB.Menu x 
      Caption         =   "ghghjfggh"
      Begin VB.Menu v 
         Caption         =   "hgfdee"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdapply_Click()
sMenu1.Style = Combo1.ListIndex
sMenu1.HighLightStyle = Combo2.ListIndex
sMenu1.Refresh



End Sub

Private Sub ctxHookMenu1_ItemClick(Key As String)

End Sub

Private Sub Form_Load()
Combo1.ListIndex = sMenu1.Style
Combo2.ListIndex = sMenu1.HighLightStyle
End Sub

Private Sub sMenu1_ItemClick(Key As String)
Select Case Key
Case "mnuFileNew"
MsgBox "clicked on new", vbinfo




End Select

End Sub

Private Sub sMenu1_ItemDescription(Description As String)
Label2.Caption = Description

End Sub

Private Sub sMenu2_ItemClick(Key As String)

End Sub
