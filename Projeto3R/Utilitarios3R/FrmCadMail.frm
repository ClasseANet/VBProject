VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "CODEJO~1.OCX"
Begin VB.Form FrmCadMail 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Mensagem"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   10380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin XtremeSuiteControls.PushButton CmdSendeMAIL 
      Height          =   675
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   1191
      _StockProps     =   79
      Caption         =   "&Enviar e-Mail"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.WebBrowser TxtMail 
      Height          =   6135
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   10095
      _Version        =   720898
      _ExtentX        =   17806
      _ExtentY        =   10821
      _StockProps     =   173
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.PushButton CmdTo 
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   975
      _Version        =   720898
      _ExtentX        =   1720
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "&Para"
      Enabled         =   0   'False
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdCC 
      Height          =   315
      Left            =   1320
      TabIndex        =   3
      Top             =   480
      Width           =   975
      _Version        =   720898
      _ExtentX        =   1720
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "&Copia"
      Enabled         =   0   'False
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit TxtTo 
      Height          =   315
      Left            =   2400
      TabIndex        =   2
      Top             =   120
      Width           =   7815
      _Version        =   720898
      _ExtentX        =   13785
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   16777215
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit TxtCC 
      Height          =   315
      Left            =   2400
      TabIndex        =   4
      Top             =   480
      Width           =   7815
      _Version        =   720898
      _ExtentX        =   13785
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   16777215
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit TxtSubject 
      Height          =   315
      Left            =   2400
      TabIndex        =   6
      Top             =   840
      Width           =   7815
      _Version        =   720898
      _ExtentX        =   13785
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   16777215
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit TxtAnexos 
      Height          =   315
      Left            =   2400
      TabIndex        =   7
      Top             =   1200
      Width           =   7815
      _Version        =   720898
      _ExtentX        =   13785
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483644
      BackColor       =   -2147483644
      Appearance      =   6
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdAnexos 
      Height          =   315
      Left            =   1320
      TabIndex        =   9
      Top             =   1200
      Width           =   975
      _Version        =   720898
      _ExtentX        =   1720
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "&Anexos"
      Enabled         =   0   'False
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdSendSMS 
      Height          =   675
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
      _Version        =   720898
      _ExtentX        =   1931
      _ExtentY        =   1191
      _StockProps     =   79
      Caption         =   "&Enviar SMS"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Label LblTitulo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ass&unto: "
      Height          =   195
      Left            =   1560
      TabIndex        =   5
      Top             =   960
      Width           =   660
   End
End
Attribute VB_Name = "FrmCadMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mvarSys As Object
Public xMail As Object
Public sFileHtm As String
Public sMensagem As String
Public sAssunto As String
Public sDestino As String
Public sNomeDest As String
Public sCopia As String
Public sAnexo As String
Public bResult As Boolean
Private Sub CmdSendeMail_Click()
   Call EnviarMensagem
End Sub
Public Sub EnviarMensagem()
   Me.MousePointer = vbHourglass
   If xMail Is Nothing Then
      Set xMail = CriarObjeto("CAMail.SendMail")
      Call MontarMail(mvarSys, xMail)
   End If
   With xMail
      .Subject = Me.TxtSubject
      .Message = Me.TxtMail.InnerHTML
      If ExisteArquivo(sAnexo) Then .Attachment = sAnexo
      .FromDisplayName = mvarSys.GetParam("FromDisplayName") 'FromDisplayName ' "Diogenes"
      .Recipient = Me.TxtTo.Text
      If Trim(.Message) <> "" And Trim(.Recipient) <> "" Then
         .Recipient = Me.TxtTo.Text
         If Trim(sNomeDest) <> "" Then .RecipientDisplayName = sNomeDest
         .Connect
         .Send
      End If
      .Disconnect
      bResult = .SendSuccesful
      
      ExibirInformacao IIf(bResult, "Mensagem enviada com Sucesso!!", "Envio Falhou!!")
   End With
   Me.MousePointer = vbDefault
   If bResult Then
      Unload Me
   End If
End Sub

Private Sub CmdSendSMS_Click()
   Call EnviarSMS
End Sub
Private Sub EnviarSMS()
'   On Error GoTo TrataErro
'   Me.MSComm1.CommPort = 3 '1
'   'Me.MSComm1.CommPort = 1
'   Me.MSComm1.Settings = "2400,N,8,1"
'   'Me.MSComm1.Settings = "115200,n,8,1"
'   ''Me.MSComm1.Settings = "9600,n,8,1"
'
'   Me.MSComm1.RThreshold = 2
'   Me.MSComm1.InputLen = 2
'   'Me.MSComm1.InputLen = 0
'
'   Me.MSComm1.DTREnable = False
'   Me.MSComm1.PortOpen = True
'
'   Me.MSComm1.Output = "AT" & Chr$(13) & Chr(10)
'   Me.MSComm1.Output = "AT+CMGF=1" & Chr$(13) & Chr(10)
'   Me.MSComm1.Output = "AT+CSCA=+550310000010" & Chr(13)
'   Me.MSComm1.Output = "AT+CMGS=" & Chr$(34) & "+5521988108541" & Chr$(34) & Chr(13) & Chr(10)
'   Me.MSComm1.Output = "This is a testing message 2" & Chr(26)
'   Me.MSComm1.PortOpen = False
'TrataErro:
'   If Err.Number <> 0 Then
'      MsgBox Err.Number & " - " & Err.Description
'   End If
End Sub
'Public Function SendSMS(CSCA As String, número As String, msg As String) As Boolean
'
'   Dim PDU, PNum, psmsc, PMSG As String
'   Dim leng As String
'   Dim comprimento As Integer
'
'   comprimento = Len(msg)
'   comprimento = 2 backup.sh createSimpleTask.sh createTasks4Site.sh runAllTask.sh getTaskId.sh createTask.sh comprimento taskInfo.sh runTask.sh
'   leng = Hex(comprimento)
'   If comprimento = 16 Then leng = "0" & Comp
'   psmsc = Trim(telc(CSCA))
'   PNum = Trim(telc(Num))
'   PMSG = Trim(ascg(msg))
'   PDU = PREX psmsc & & & midx PNum sufx & & & Comp PMSG
'sono (1)
'   mobcomm.Output = "AT + CMGF = 0" + vbCr
'   mobcomm.Output vbCr = "AT + CMGS =" & Str(15 comprimento +) +
'   mobcomm.Output = PDU & Chr$(26)
'   sono (1)
'SendSMS = True
'End Function
Private Sub Form_Activate()
   'Call PopulaTela
   Me.Visible = True
End Sub
Private Sub Form_Load()
   Call PopulaTela
   Screen.MousePointer = vbNormal
End Sub
Private Sub PopulaTela()
   With Me
      .Caption = "Mensagem: " & mvarSys.Propriedades("FromDisplayName")
      .TxtTo.Text = sDestino
      .TxtCC.Text = sCopia
      .TxtSubject.Text = sAssunto
      .TxtAnexos.Text = sAnexo
      With .TxtMail
         .Appearance = 1
         .BorderStyle = 0
         .StaticText = False
         .WebBrowserContextMenu = True
         If ExisteArquivo(sFileHtm) Then
            .Navigate sFileHtm
         Else
            .InnerHTML = sMensagem
         End If
         .Refresh
      End With
   End With
End Sub
