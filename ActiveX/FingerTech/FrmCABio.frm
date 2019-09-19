VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmCABio 
   Caption         =   "Biometria"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4695
   Icon            =   "FrmCABio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkServico 
      Caption         =   "Habilitar Serviço"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   3120
         Top             =   840
      End
      Begin VB.CommandButton CmdIdent 
         Caption         =   "Identificar"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton CmdReg 
         Caption         =   "Registrar"
         Height          =   375
         Left            =   1440
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton CmdLoadDB 
         Caption         =   "Carregar Banco (.fdb)"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CommandButton CmdSaveDB 
         Caption         =   "Salvar Banco (.fdb)"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   2040
         Width           =   2055
      End
      Begin VB.CommandButton CmdApagar 
         Caption         =   "Apagar"
         Height          =   375
         Left            =   2760
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox TxtStrConect 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   2640
         Width           =   4215
      End
      Begin VB.CommandButton CmdUnloadDb 
         Caption         =   "Descarregar Banco (.fdb)"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1575
         Left            =   2760
         Top             =   720
         Width           =   1575
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmCABio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Activate()
Event Load()
Event ChkServicoClick()
Event CmdApagarClick()
Event CmdIdentClick()
Event CmdLoadDBClick()
Event CmdRegClick()
Event CmdSaveDBClick()
Event CmdUnloadDbClick()
Event Timer1Timer()
Private Sub ChkServico_Click()
   RaiseEvent ChkServicoClick
End Sub
Private Sub CmdApagar_Click()
   RaiseEvent CmdApagarClick
End Sub
Private Sub CmdIdent_Click()
   RaiseEvent CmdIdentClick
End Sub
Private Sub CmdLoadDB_Click()
   RaiseEvent CmdLoadDBClick
End Sub
Private Sub CmdReg_Click()
   RaiseEvent CmdRegClick
End Sub
Private Sub CmdSaveDB_Click()
   RaiseEvent CmdSaveDBClick
End Sub
Private Sub CmdUnloadDb_Click()
   RaiseEvent CmdUnloadDbClick
End Sub
Private Sub Form_Activate()
   RaiseEvent Activate
End Sub
Private Sub Form_Load()
   RaiseEvent Load
End Sub
Private Sub Timer1_Timer()
   RaiseEvent Timer1Timer
End Sub
