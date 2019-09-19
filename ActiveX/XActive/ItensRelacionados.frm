VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.Controls.v11.2.2.ocx"
Begin VB.Form FrmItensRelacionados 
   AutoRedraw      =   -1  'True
   Caption         =   "Itens Relacionados"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6315
   Icon            =   "ItensRelacionados.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   6315
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton CmdAdd 
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   600
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Adcionar"
      UseVisualStyle  =   -1  'True
   End
   Begin MSComctlLib.ListView LstItens 
      Height          =   3135
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   5530
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
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
   Begin XtremeSuiteControls.PushButton CmdOper 
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   2880
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Ok"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdRemove 
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Remover"
      UseVisualStyle  =   -1  'True
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
