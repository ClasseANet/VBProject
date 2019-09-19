VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "COD063~1.OCX"
Begin VB.Form FrmADD 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Adicionar"
   ClientHeight    =   4575
   ClientLeft      =   16185
   ClientTop       =   6540
   ClientWidth     =   6630
   ClipControls    =   0   'False
   Icon            =   "FrmAdd.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.PushButton CmdOk 
      Default         =   -1  'True
      Height          =   375
      Left            =   5520
      TabIndex        =   0
      Top             =   4120
      Width           =   975
      _Version        =   720898
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Ok"
      ForeColor       =   12582912
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdCancelar 
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   4120
      Width           =   975
      _Version        =   720898
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Cancelar"
      ForeColor       =   192
      UseVisualStyle  =   -1  'True
   End
   Begin VB.CommandButton CmdChk 
      Caption         =   "Selecionados"
      Height          =   360
      Index           =   3
      Left            =   3000
      TabIndex        =   10
      Top             =   4120
      Width           =   1095
   End
   Begin VB.CommandButton CmdChk 
      Caption         =   "Inverter"
      Height          =   360
      Index           =   2
      Left            =   2040
      TabIndex        =   9
      Top             =   4120
      Width           =   975
   End
   Begin VB.CommandButton CmdChk 
      Caption         =   "Nenhum"
      Height          =   360
      Index           =   1
      Left            =   1080
      TabIndex        =   8
      Top             =   4120
      Width           =   975
   End
   Begin VB.CommandButton CmdChk 
      Caption         =   "Todos"
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   4120
      Width           =   975
   End
   Begin VB.TextBox TxtChave 
      Height          =   315
      Left            =   1920
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   4680
      Width           =   1695
   End
   Begin VB.ComboBox CmbCampo 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Text            =   "CmbCampo"
      Top             =   4680
      Width           =   1815
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
      Appearance      =   0
      NumItems        =   0
   End
End
Attribute VB_Name = "FrmADD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Event CmdOkClick()
Event CmdCancelarClick()
Event CmdChkClick(Index As Integer)
Event LstItensDblClick()
Event LstItensItemClick(ByVal Item As MSComctlLib.ListItem)
Event Activate()
Event Load()
Private Sub CmdCancelar_Click()
   RaiseEvent CmdCancelarClick
End Sub
Private Sub CmdChk_Click(Index As Integer)
   RaiseEvent CmdChkClick(Index)
End Sub
Private Sub CmdOk_Click()
   RaiseEvent CmdOkClick
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
   RaiseEvent CmdOkClick
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
Private Sub LstItens_DblClick()
   RaiseEvent LstItensDblClick
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

Private Sub LstItens_ItemClick(ByVal Item As MSComctlLib.ListItem)
   RaiseEvent LstItensItemClick(Item)
End Sub
