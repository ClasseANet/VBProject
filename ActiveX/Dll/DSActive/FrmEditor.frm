VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form FrmEditor 
   AutoRedraw      =   -1  'True
   Caption         =   "Erros"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7635
   ClipControls    =   0   'False
   Icon            =   "FrmEditor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   7635
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox TxtTexto 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   7011
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"FrmEditor.frx":08CA
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
      Left            =   6240
      TabIndex        =   1
      Top             =   4200
      WhatsThisHelpID =   10501
      Width           =   1185
      _Version        =   65536
      _ExtentX        =   2090
      _ExtentY        =   714
      _StockProps     =   78
      Caption         =   "&Ok"
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
      MouseIcon       =   "FrmEditor.frx":094C
      Picture         =   "FrmEditor.frx":0968
   End
   Begin Threed.SSCommand CmdOper 
      Height          =   405
      Index           =   1
      Left            =   4680
      TabIndex        =   2
      Top             =   4200
      WhatsThisHelpID =   10501
      Width           =   1185
      _Version        =   65536
      _ExtentX        =   2090
      _ExtentY        =   714
      _StockProps     =   78
      Caption         =   "&Cancel"
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
      AutoSize        =   1
      MouseIcon       =   "FrmEditor.frx":0984
      Picture         =   "FrmEditor.frx":09A0
   End
   Begin Threed.SSCommand CmdImprimir 
      Height          =   405
      Left            =   120
      TabIndex        =   3
      Top             =   4200
      WhatsThisHelpID =   10501
      Width           =   1185
      _Version        =   65536
      _ExtentX        =   2090
      _ExtentY        =   714
      _StockProps     =   78
      Caption         =   "&Imprimir"
      ForeColor       =   -2147483641
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   1
      MouseIcon       =   "FrmEditor.frx":09BC
      Picture         =   "FrmEditor.frx":09D8
   End
   Begin Crystal.CrystalReport CryRprt 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "FrmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PrimeiraVez As Boolean
Event Active()
Event Load()
Event Resize()
Event CmdOperClick(index As Integer)
Event CmdImprimirClick()
Event TxtTextoKeyPress(KeyAscii As Integer)
Event TxtTextoKeyDown(KeyCode As Integer, Shift As Integer)
Event TxtTextoChange()
Private Sub CmdOper_Click(index As Integer)
   RaiseEvent CmdOperClick(index)
End Sub
Private Sub CmdImprimir_Click()
   RaiseEvent CmdImprimirClick
End Sub
Private Sub Form_Activate()
   If PrimeiraVez Then
      PrimeiraVez = False
      RaiseEvent Resize
   End If
   RaiseEvent Active
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then
      Call CmdOper_Click(0)
   End If
End Sub
Private Sub Form_Load()
   PrimeiraVez = True
   RaiseEvent Load
End Sub
Private Sub Form_Resize()
   RaiseEvent Resize
End Sub
Private Sub TxtTexto_Change()
   RaiseEvent TxtTextoChange
End Sub
Private Sub TxtTexto_KeyDown(KeyCode As Integer, Shift As Integer)
   RaiseEvent TxtTextoKeyDown(KeyCode, Shift)
End Sub
Private Sub TxtTexto_KeyPress(KeyAscii As Integer)
   RaiseEvent TxtTextoKeyPress(KeyAscii)
End Sub
