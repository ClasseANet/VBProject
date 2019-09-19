VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmItensRelacionados 
   AutoRedraw      =   -1  'True
   Caption         =   "Itens Relacionados"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6315
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   6315
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSCommand CmdADD 
      Height          =   765
      Left            =   4800
      TabIndex        =   1
      Top             =   480
      WhatsThisHelpID =   10463
      Width           =   1440
      _Version        =   65536
      _ExtentX        =   2540
      _ExtentY        =   1349
      _StockProps     =   78
      Caption         =   "Adicionar"
      Font3D          =   3
      Picture         =   "ItensRelacionados.frx":0000
   End
   Begin Threed.SSCommand CmdREMOVE 
      Height          =   765
      Left            =   4800
      TabIndex        =   2
      Top             =   1320
      WhatsThisHelpID =   10462
      Width           =   1440
      _Version        =   65536
      _ExtentX        =   2540
      _ExtentY        =   1349
      _StockProps     =   78
      Caption         =   "Remover"
      Font3D          =   3
      Picture         =   "ItensRelacionados.frx":031A
   End
   Begin Threed.SSCommand CmdOper 
      Height          =   615
      Left            =   4800
      TabIndex        =   3
      Tag             =   "Sair"
      Top             =   2760
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "Sai&r"
      ForeColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
      RoundedCorners  =   0   'False
   End
   Begin MSComctlLib.ListView LstItens 
      Height          =   3135
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   5530
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Itens"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      WhatsThisHelpID =   10246
      Width           =   510
   End
End
Attribute VB_Name = "FrmItensRelacionados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event CmdADDClick()
Event CmdOperClick()
Event CmdREMOVEClick()
Event Load()
Event Activate()
Private Sub CmdADD_Click()
   RaiseEvent CmdADDClick
End Sub
Private Sub CmdOper_Click()
   RaiseEvent CmdOperClick
End Sub
Private Sub CmdREMOVE_Click()
   RaiseEvent CmdREMOVEClick
End Sub
Private Sub Form_Activate()
   RaiseEvent Activate
End Sub
Private Sub Form_Load()
   RaiseEvent Load
End Sub
Private Sub LstItens_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   With Me.LstItens
      If .SortKey = ColumnHeader.index - 1 Then
         If .SortOrder = lvwDescending Then
            .SortOrder = lvwAscending
         Else
            .SortOrder = lvwDescending
         End If
      Else
         .SortKey = ColumnHeader.index - 1
         .SortOrder = lvwAscending
      End If
      .Sorted = True
   End With
End Sub
