VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.Controls.v11.2.2.ocx"
Begin VB.Form FrmCAVersao 
   BackColor       =   &H8000000C&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "C.A. Versão..."
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5310
   Icon            =   "CAVersao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton CmdFTP 
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2175
      _Version        =   720898
      _ExtentX        =   3836
      _ExtentY        =   1508
      _StockProps     =   79
      Caption         =   "F.T.P."
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdVerificar 
      Height          =   855
      Left            =   2880
      TabIndex        =   1
      Top             =   360
      Width           =   2175
      _Version        =   720898
      _ExtentX        =   3836
      _ExtentY        =   1508
      _StockProps     =   79
      Caption         =   "Verificar Versão..."
      UseVisualStyle  =   -1  'True
   End
End
Attribute VB_Name = "FrmCAVersao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdFTP_Click()
   Dim MyVerif As Object
   If "DIO" = UCase(InputBox("Digite a chave de entrada.", "Chave...")) Then
      Set MyVerif = CriarObjeto("VersaoFTP.TL_VerifVersao")
      MyVerif.ShowFTP
   End If
   Set MyVerif = Nothing
End Sub
Private Sub CmdVerificar_Click()
   Dim MyVerif As Object
   
   Set MyVerif = CriarObjeto("VersaoFTP.TL_VerifVersao")
   MyVerif.ShowCAVs
   Set MyVerif = Nothing
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Me.MousePointer = vbHourglass
   Screen.MousePointer = vbHourglass
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Screen.MousePointer = vbHourglass
End Sub
