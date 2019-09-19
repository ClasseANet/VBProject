VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "Msinet.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.Controls.v11.2.2.ocx"
Begin VB.Form FrmCAFTP 
   Caption         =   "FTP "
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   8535
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   5055
      Left            =   120
      TabIndex        =   20
      Top             =   1200
      Width           =   4215
      _Version        =   720898
      _ExtentX        =   7435
      _ExtentY        =   8916
      _StockProps     =   68
      Appearance      =   9
      Color           =   4
      ItemCount       =   2
      Item(0).Caption =   "Diretório"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "TabControlPage1"
      Item(1).Caption =   "Arquivos"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "TabControlPage2"
      Begin XtremeSuiteControls.TabControlPage TabControlPage2 
         Height          =   4695
         Left            =   -69970
         TabIndex        =   22
         Top             =   330
         Visible         =   0   'False
         Width           =   4155
         _Version        =   720898
         _ExtentX        =   7329
         _ExtentY        =   8281
         _StockProps     =   1
         Page            =   1
         Begin VB.CommandButton CmdEnviarLista 
            Caption         =   "Enviar Lista"
            Height          =   375
            Left            =   2160
            TabIndex        =   29
            Top             =   4200
            Width           =   1335
         End
         Begin VB.CommandButton CmdRemover 
            Caption         =   "Remover"
            Height          =   375
            Left            =   240
            TabIndex        =   28
            Top             =   4200
            Width           =   1335
         End
         Begin VB.ListBox LstArquivos 
            Appearance      =   0  'Flat
            Height          =   3930
            Left            =   120
            TabIndex        =   27
            Top             =   120
            Width           =   3975
         End
      End
      Begin XtremeSuiteControls.TabControlPage TabControlPage1 
         Height          =   4695
         Left            =   30
         TabIndex        =   21
         Top             =   330
         Width           =   4155
         _Version        =   720898
         _ExtentX        =   7329
         _ExtentY        =   8281
         _StockProps     =   1
         Page            =   0
         Begin VB.Frame fraLocal 
            Caption         =   "Local Host"
            Height          =   4695
            Left            =   120
            TabIndex        =   23
            Top             =   0
            Width           =   4095
            Begin VB.CommandButton CmdAddList 
               Caption         =   "Adicionar à Lista"
               Height          =   375
               Left            =   1920
               TabIndex        =   30
               Top             =   4200
               Width           =   1935
            End
            Begin VB.DriveListBox drvList 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   120
               TabIndex        =   26
               Top             =   360
               Width           =   1695
            End
            Begin VB.DirListBox dirList 
               Appearance      =   0  'Flat
               Height          =   3240
               Left            =   120
               TabIndex        =   25
               Top             =   840
               Width           =   1695
            End
            Begin VB.FileListBox filList 
               Appearance      =   0  'Flat
               Height          =   3735
               Left            =   1920
               MultiSelect     =   2  'Extended
               TabIndex        =   24
               Top             =   360
               Width           =   2055
            End
         End
      End
   End
   Begin XtremeSuiteControls.ProgressBar PrbFTP 
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   6300
      Width           =   8295
      _Version        =   720898
      _ExtentX        =   14631
      _ExtentY        =   450
      _StockProps     =   93
      UseVisualStyle  =   -1  'True
   End
   Begin VB.CommandButton CmdRecebe 
      Appearance      =   0  'Flat
      Caption         =   "<--"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton CmdEnvia 
      Appearance      =   0  'Flat
      Caption         =   "-->"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3120
      Width           =   495
   End
   Begin VB.Frame fraRemote 
      Caption         =   "Host Remoto"
      Height          =   4935
      Left            =   5040
      TabIndex        =   10
      Top             =   1320
      Width           =   3375
      Begin VB.CommandButton CmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   4320
         Width           =   1215
      End
      Begin VB.CommandButton CmdMkDir 
         Appearance      =   0  'Flat
         Caption         =   "&MkDir"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   4320
         Width           =   1215
      End
      Begin VB.ListBox lstRemote 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   3345
         Left            =   240
         MultiSelect     =   2  'Extended
         TabIndex        =   11
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label lblRemoteDirectory 
         AutoSize        =   -1  'True
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   60
      End
   End
   Begin VB.Frame fraStatus 
      Caption         =   "Status"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   7920
      Width           =   8295
      Begin VB.Label lblStatus 
         Caption         =   "Pronto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   480
         TabIndex        =   9
         Top             =   240
         Width           =   6930
      End
   End
   Begin VB.Frame fraFTP 
      Caption         =   "FTP :"
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   6255
      Begin VB.CommandButton CmdConecta 
         Appearance      =   0  'Flat
         Caption         =   "&Conectar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1080
         PasswordChar    =   "#"
         TabIndex        =   7
         Text            =   "dolphin"
         Top             =   840
         Width           =   3615
      End
      Begin VB.TextBox txtUsername 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Text            =   "classeanet"
         Top             =   540
         Width           =   3615
      End
      Begin VB.TextBox txtAddress 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Text            =   "ftp.classeanet.com.br"
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label lblPassword 
         Caption         =   "Senha :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1000
      End
      Begin VB.Label lblUsername 
         Caption         =   "Usuário:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1000
      End
      Begin VB.Label lblAddress 
         Caption         =   "Endereço :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1000
      End
   End
   Begin VB.TextBox txtLog 
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6600
      Width           =   8295
   End
   Begin InetCtlsObjects.Inet InetFTP 
      Left            =   7320
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   2
      RemotePort      =   21
      URL             =   "ftp://"
      RequestTimeout  =   10
   End
   Begin XtremeSuiteControls.CheckBox ChkZIP 
      Height          =   255
      Left            =   4440
      TabIndex        =   19
      Top             =   2520
      Width           =   615
      _Version        =   720898
      _ExtentX        =   1085
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "ZIP"
   End
   Begin VB.Menu Mnu 
      Caption         =   "Administração"
      Index           =   0
      Begin VB.Menu MnuAdmin 
         Caption         =   "&Sair"
         Index           =   0
      End
   End
   Begin VB.Menu Mnu 
      Caption         =   "Ferramentas"
      Index           =   1
      Begin VB.Menu MnuFerr 
         Caption         =   "Verificar Versão..."
         Index           =   0
      End
      Begin VB.Menu MnuFerr 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu MnuFerr 
         Caption         =   "Liberar Tela..."
         Index           =   2
      End
   End
End
Attribute VB_Name = "FrmCAFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public gCODSIS As String
Event Activate()
Event Load()
Event Unload(Cancel As Integer)
Event CmdAddListClick()
Event CmdEnviaClick()
Event CmdConectaClick()
Event CmdDeleteClick()
Event CmdEnviarListaClick()
Event MnuFerr(Index As Integer)
Event TabControl1BeforeItemClick(ByVal Item As XtremeSuiteControls.ITabControlItem, Cancel As Variant)
Private Sub CmdAddList_Click()
   RaiseEvent CmdAddListClick
End Sub
Private Sub CmdConecta_Click()
   Screen.MousePointer = vbHourglass
   PrbFTP.Scrolling = xtpProgressBarMarquee
   RaiseEvent CmdConectaClick
   PrbFTP.Scrolling = xtpProgressBarStandard
   Screen.MousePointer = vbDefault
End Sub
Private Sub CmdDelete_Click()
   Screen.MousePointer = vbHourglass
   PrbFTP.Scrolling = xtpProgressBarMarquee
    
   RaiseEvent CmdDeleteClick
   
   PrbFTP.Scrolling = xtpProgressBarStandard
   Screen.MousePointer = vbDefault
End Sub
Private Sub CmdEnviarLista_Click()
   RaiseEvent CmdEnviarListaClick
End Sub

Private Sub cmdMkDir_Click()
    Dim dir As String, operacao As String
    
    dir = InputBox("Informe o nome da pasta", "Cria Diretório")
    If dir <> "" Then
        operacao = "mkdir " & dir
        ExecutaComando operacao, True
    End If
End Sub

Private Sub cmdRecebe_Click()
   Screen.MousePointer = vbHourglass
   PrbFTP.Scrolling = xtpProgressBarMarquee

    Dim contador As Integer
    Dim operacao As String
    Dim nomeArquivo As String, arquivoSaida As String
    
    For contador = 0 To lstRemote.ListCount - 1
        If lstRemote.Selected(contador) = True Then
            nomeArquivo = lstRemote.List(contador)
            If Len(dirList.Path) > 3 Then
                arquivoSaida = dirList.Path & "\" & nomeArquivo
            Else
                arquivoSaida = dirList.Path & nomeArquivo
            End If
            operacao = "recv " & nomeArquivo & " " & arquivoSaida
            ExecutaComando operacao, False
            lstRemote.Selected(contador) = False
        End If
    Next contador
   filList.Refresh
   PrbFTP.Scrolling = xtpProgressBarStandard
   Screen.MousePointer = vbDefault
End Sub

Private Sub CmdEnvia_Click()
   Screen.MousePointer = vbHourglass
   PrbFTP.Scrolling = xtpProgressBarMarquee
   RaiseEvent CmdEnviaClick
   PrbFTP.Scrolling = xtpProgressBarStandard
   Screen.MousePointer = vbDefault
End Sub
Private Sub dirList_Change()
    filList.Path = dirList.Path
End Sub
Private Sub drvList_Change()
    On Error GoTo driveError
    
    dirList.Path = drvList.Drive
    Exit Sub
driveError:
    MsgBox Err.Description, vbExclamation, "Drive Error"
End Sub
Private Sub Form_Activate()
   RaiseEvent Activate
End Sub
Private Sub Form_Load()
   RaiseEvent Load
End Sub

Private Sub Form_Resize()
   On Error Resume Next
   Me.fraRemote.Left = Me.Width - Me.fraRemote.Width - 120
   Me.ChkZIP.Left = Me.fraRemote.Left - Me.ChkZIP.Width - 60
   Me.CmdEnvia.Left = Me.ChkZIP.Left
   Me.CmdRecebe.Left = Me.ChkZIP.Left
   Me.TabControl1.Width = Me.ChkZIP.Left - Me.TabControl1.Left - 120
   Me.LstArquivos.Width = Me.TabControl1.Width - Me.LstArquivos.Left - 120
   
   Me.fraLocal.Left = 60
   Me.fraLocal.Width = Me.TabControl1.Width - 2 * Me.fraLocal.Left
   Me.dirList.Left = 60
   Me.dirList.Width = (Me.fraLocal.Width / 2) - Me.dirList.Left - 60
   If Me.dirList.Width > 3000 Then
      Me.dirList.Width = 3000
   End If
   Me.filList.Left = Me.dirList.Width + 180
   Me.filList.Width = Me.fraLocal.Width - Me.filList.Left - 60
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Screen.MousePointer = vbHourglass
   RaiseEvent Unload(Cancel)
   
   On Error Resume Next
   Call ExecutaComando("quit", False)
   'Me.InetFTP.Execute , "quit"
   Do While InetFTP.StillExecuting
      DoEvents
   Loop
   Screen.MousePointer = vbDefault
End Sub


'Private Sub InetFTP_StateChanged(ByVal State As Integer)
'    Select Case State
'        Case icResolvingHost
'            lblStatus.Caption = "Resolvendo Host"
'        Case icHostResolved
'            lblStatus.Caption = "Host Resolvido"
'        Case icConnecting
'            lblStatus.Caption = "Conectando ..."
'        Case icConnected
'            lblStatus.Caption = "Conectado"
'        Case icRequesting
'            lblStatus.Caption = "Requesitando ..."
'        Case icRequestSent
'            lblStatus.Caption = "Requesição enviada"
'        Case icReceivingResponse
'            lblStatus.Caption = "Recebendo ..."
'        Case icResponseReceived
'            lblStatus.Caption = "Resposta recebida"
'        Case icDisconnecting
'            lblStatus.Caption = "Desconectando ..."
'        Case icDisconnected
'            lblStatus.Caption = "Desconectado"
'        Case icError
'            lblStatus.Caption = InetFTP.ResponseInfo
'            txtLog.Text = txtLog.Text & InetFTP.ResponseInfo & vbCrLf
'        Case icResponseCompleted
'            lblStatus.Caption = "operacao Completa"
'            txtLog.Text = txtLog.Text & "operacao Completa" & vbCrLf
'    End Select
'    txtLog.SelStart = Len(txtLog.Text)
'End Sub
Private Sub lstRemote_DblClick()
    Dim operacao As String, dir As String
    ' Se o item é uma pasta muda para a pasta
    If Right(lstRemote.List(lstRemote.ListIndex), 1) = "/" Then
        dir = lstRemote.List(lstRemote.ListIndex)
        operacao = "cd " & Left(dir, Len(dir) - 1)
        ExecutaComando operacao, True
    End If
End Sub
Private Sub MnuAdmin_Click(Index As Integer)
   Unload Me
End Sub

Private Sub MnuFerr_Click(Index As Integer)
   RaiseEvent MnuFerr(Index)
End Sub
Private Sub TabControl1_BeforeItemClick(ByVal Item As XtremeSuiteControls.ITabControlItem, Cancel As Variant)
   RaiseEvent TabControl1BeforeItemClick(Item, Cancel)
End Sub
Private Sub txtLog_GotFocus()
    'Me.lblAddress.SetFocus
End Sub
Private Sub ExecutaComando(ByVal op As String, ByVal Ld As Boolean)
 
    On Error GoTo Trata_erro

    If InetFTP.StillExecuting Then
        InetFTP.Cancel
    End If
    txtLog.Text = txtLog.Text & "Comando: " & op & vbCrLf
    InetFTP.Execute , op
    TerminaComando
    If Ld = True Then
        ListaDir
        TerminaComando
    End If
    Exit Sub
Trata_erro:
    MsgBox "Não foi possivel efetuar operacao com : " & txtAddress.Text & vbCrLf & " erro : " & Err.Number
End Sub

Private Sub TerminaComando()
    Do While InetFTP.StillExecuting
        DoEvents
    Loop
End Sub

Private Sub ListaDir()
    Dim operacao As String
    Dim data As Variant, contador As Integer
    Dim inicio As Integer, length As Integer
    
    inicio = 1
    lstRemote.Clear
    operacao = "dir"
    ExecutaComando operacao, False
    Do
        data = InetFTP.GetChunk(2048, icString)
        DoEvents
        For contador = 1 To Len(data)
            If Mid(data, contador, 1) = Chr(13) Then
                If length > 0 And Mid(data, inicio, length) <> "./" Then
                    lstRemote.AddItem Mid(data, inicio, length)
                End If
                inicio = contador + 2
                length = -1
            Else
                length = length + 1
            End If
        Next contador
    Loop While LenB(data) > 0
    operacao = "pwd"
    ExecutaComando operacao, False
    lblRemoteDirectory.Caption = InetFTP.GetChunk(1024, icString)
End Sub

