VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmADD 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Adicionar"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdChk 
      Caption         =   "Selecionados"
      Height          =   360
      Index           =   3
      Left            =   3000
      TabIndex        =   10
      Top             =   4100
      Width           =   1095
   End
   Begin VB.CommandButton CmdChk 
      Caption         =   "Inverter"
      Height          =   360
      Index           =   2
      Left            =   2040
      TabIndex        =   9
      Top             =   4100
      Width           =   975
   End
   Begin VB.CommandButton CmdChk 
      Caption         =   "Nenhum"
      Height          =   360
      Index           =   1
      Left            =   1080
      TabIndex        =   8
      Top             =   4100
      Width           =   975
   End
   Begin VB.CommandButton CmdChk 
      Caption         =   "Todos"
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   4100
      Width           =   975
   End
   Begin Threed.SSCommand CmdLocalizar 
      Height          =   330
      Left            =   3960
      TabIndex        =   6
      ToolTipText     =   "Filtrar Lista"
      Top             =   4440
      Width           =   330
      _Version        =   65536
      _ExtentX        =   582
      _ExtentY        =   582
      _StockProps     =   78
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      Font3D          =   3
      RoundedCorners  =   0   'False
      MouseIcon       =   "FrmAdd.frx":0000
      Picture         =   "FrmAdd.frx":001C
   End
   Begin Threed.SSCommand CmdFiltrar 
      Height          =   330
      Left            =   3600
      TabIndex        =   5
      ToolTipText     =   "Filtrar Lista"
      Top             =   4440
      Width           =   330
      _Version        =   65536
      _ExtentX        =   582
      _ExtentY        =   582
      _StockProps     =   78
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      Font3D          =   3
      RoundedCorners  =   0   'False
      MouseIcon       =   "FrmAdd.frx":0038
      Picture         =   "FrmAdd.frx":0054
   End
   Begin VB.TextBox TxtChave 
      Height          =   315
      Left            =   1920
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   4440
      Width           =   1695
   End
   Begin VB.ComboBox CmbCampo 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Text            =   "CmbCampo"
      Top             =   4440
      Width           =   1815
   End
   Begin Threed.SSCommand CmdOper 
      Height          =   360
      Index           =   1
      Left            =   4440
      TabIndex        =   1
      Tag             =   "Sair"
      Top             =   4095
      Width           =   990
      _Version        =   65536
      _ExtentX        =   1746
      _ExtentY        =   635
      _StockProps     =   78
      Caption         =   "&Cancelar"
      ForeColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
   End
   Begin Threed.SSCommand CmdOper 
      Default         =   -1  'True
      Height          =   360
      Index           =   0
      Left            =   5520
      TabIndex        =   0
      Tag             =   "Sair"
      Top             =   4095
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   635
      _StockProps     =   78
      Caption         =   "&Ok"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
   End
   Begin MSComctlLib.ListView LstItens 
      Height          =   4095
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   7223
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "FrmADD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Event CmdOperClick(Index As Integer)
Event CmdChkClick(Index As Integer)
Event Activate()
Event Load()
Private Sub CmdChk_Click(Index As Integer)
   RaiseEvent CmdChkClick(Index)
End Sub
Private Sub CmdOper_Click(Index As Integer)
   RaiseEvent CmdOperClick(Index)
End Sub
Private Sub Form_Activate()
   Screen.MousePointer = vbDefault
   RaiseEvent Activate
   Call ClsCtrl.PintarFundo(Me)
End Sub
Private Sub Form_Load()
   RaiseEvent Load
   Call ClsCtrl.PintarFundo(Me)
End Sub
Private Sub LstSelecao_DblClick()
   RaiseEvent CmdOperClick(0)
End Sub
Private Sub LstItens_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   With Me.LstItens
      If Not .Sorted Then
         .SortOrder = lvwDescending
      End If
      .Sorted = True
      If .SortKey = ColumnHeader.Index - 1 Then
         If .SortOrder = lvwDescending Then
            .SortOrder = lvwAscending
         Else
            .SortOrder = lvwDescending
         End If
      Else
         .SortKey = ColumnHeader.Index - 1
         .SortOrder = lvwAscending
      End If
      .Sorted = True
   End With
End Sub

Private Sub LstItens_ItemCheck(ByVal Item As MSComctlLib.ListItem)
   With Me.LstItens
      If .MultiSelect Then
         For i = 1 To .ListItems.Count
            .ListItems(i).Selected = .ListItems(i).Checked
         Next
      End If
   End With
End Sub
