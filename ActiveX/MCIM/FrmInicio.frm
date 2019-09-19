VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{8153511F-FE57-47E0-A0A1-DBA712C97332}#1.0#0"; "MCIControl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmInicio 
   Caption         =   "Inicio"
   ClientHeight    =   6045
   ClientLeft      =   5625
   ClientTop       =   1020
   ClientWidth     =   3045
   Icon            =   "FrmInicio.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   3045
   Begin MCIControls.MCIButton CmdConectar 
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   5520
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Conectar"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin MCIControls.MCIMenu MCIMenu1 
      Left            =   0
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      HighlightCheckedItems=   0   'False
      TickIconIndex   =   0
      BackColor       =   -2147483633
      HaveComboBox    =   0   'False
      HaveCheckBox    =   0   'False
      HaveTextBox     =   0   'False
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   120
      Top             =   6360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInicio.frx":0442
            Key             =   "CLOSE"
            Object.Tag             =   "CLOSE"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInicio.frx":09DC
            Key             =   "OPEN"
            Object.Tag             =   "OPEN"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInicio.frx":0F76
            Key             =   "USER"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TrwUsers 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   8281
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      Style           =   3
      SingleSel       =   -1  'True
      ImageList       =   "ImgList"
      Appearance      =   0
   End
   Begin MCIControls.MCIButton CmdSair 
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   5520
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&Sair"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin VB.Frame FrmeTituloTree 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   3015
   End
   Begin MSWinsockLib.Winsock WskMensagem 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim MCICLiente    As MCIMensagem
'Dim MCIServidor   As MCIMensagem
'Dim CollCli       As Collection
'Dim CollSrv       As Collection
Private Sub CmdConectar_Click()
   Call Conecta
End Sub
Private Sub CmdSair_Click()
   Call Form_Unload(False)
End Sub
Private Sub Form_Initialize()
   On Error Resume Next
   Sys.PathUpdate = "\\Guarani\Sistemas\Admin"
   'Call AtualizaDLL(GetWindowsSysDir() & "MSWSOCK.DLL", Sys.PathUpdate & "\" & "MSWSOCK.DLL")
   Call AtualizaDLL(GetWindowsSysDir() & "MSWINSCK.OCX", Sys.PathUpdate & "\" & "MSWINSCK.OCX")
   Call ConectarBanco
End Sub
Private Sub Form_Load()
   With Me.WskMensagem
      .LocalPort = Sys.PortaPadrao
      .Listen
   End With
   
'   Set MCIServidor = New MCIMensagem
'   With MCIServidor
'      .LocalName = Sys.APELIDO
'
'      .Conexao = MCIMConexao.Servidor
'      .Porta = Sys.PortaPadrao
'      .Show
'      .Hide
'   End With
   DefineFundo "Azul"
   Me.Left = 0
   Me.Top = 300
   
   
   Call ConectarCOOL
   Call PopulaTrwUsers
End Sub

Private Sub Form_Resize()
   Me.TrwUsers.Left = 0
   Me.TrwUsers.Width = Me.Width - 120
   Me.TrwUsers.Height = Me.Height - (3 * Me.TrwUsers.Top)
   Me.FrmeTituloTree.Left = Me.TrwUsers.Left
   Me.FrmeTituloTree.Width = Me.TrwUsers.Width
   Me.CmdConectar.Top = Me.TrwUsers.Top + Me.TrwUsers.Height + 60
   Me.CmdSair.Top = Me.CmdConectar.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Dim n As Object
   On Error Resume Next
   For Each n In Forms
      Unload n
   Next
   Call DesconectarCOOL
   XDb.SrvDesconecta
   Set XDb = Nothing
   End
End Sub
Private Sub DefineFundo(Optional pCor As String = "Padrão")
   Select Case pCor
      Case "Azul":   Me.BackColor = &HE8D7CB
      Case "Padrão": Me.BackColor = &HD8E9EC
   End Select
End Sub
Private Sub ConectarBanco()
   Dim MySenha    As SENHA
   Set XDb = New DS_BANCO
   With XDb
      .SrvConecta , "MAUA", , "USU_VERIF", "DIPLOMATA", , , "TAMOIO"
      If Not .Conectado Then
         MsgBox "Não foi possível conectar servidor!", vbCritical, "Conversa On-Line"
         End
      End If
   End With
   If Sys.IDUSU = "" Then
      Set MySenha = New SENHA
      With MySenha
         Set .XDb = XDb
         .Show
         If .Conectado Then
            Sys.IDUSU = .IDUSU
         Else
            End
         End If
      End With
   End If
   
End Sub
Private Sub PopulaTrwUsers()
   Dim Sql        As String
   Dim xNode      As Node 'MSComctlLib
   Dim QtdOnLine  As Double
   Dim QtdOffLine As Double
   Dim sCOD       As String
   Dim sNMMAQ     As String
   Dim sCODPESSOA As String
   Dim sAPELIDO   As String
   Dim sNMPESSOA  As String
   Dim nIDPESSOA  As Double

   
   Me.TrwUsers.Nodes.Clear
   Set xNode = Me.TrwUsers.Nodes.Add(, , "C0", "Conectados (0)", "CLOSE")
   xNode.ExpandedImage = "OPEN"
   xNode.Sorted = True
   xNode.Expanded = True
   xNode.ForeColor = vbBlue
   xNode.Bold = True
   Call SetTag(xNode, "CAPTION", "Conectados")
   
   Sql = "Select C.NMMAQ, P.CODPESSOA, P.APELIDO, P.NMPESSOA, P.IDPESSOA"
   Sql = Sql & " From CONEXAOMSG C, PESSOA P"
   Sql = Sql & " Where C.IDUSU = P.CODPESSOA"
   Sql = Sql & " And C.IDUSU <>" & SqlStr(Sys.IDUSU)
   
   If XDb.AbreTabela(Sql) Then
      QtdOnLine = XDb.RSAux.RecordCount
      While Not XDb.RSAux.EOF
         sNMMAQ = Trim(XDb.RSAux("NMMAQ") & "")
         sCODPESSOA = Trim(XDb.RSAux("CODPESSOA") & "")
         sAPELIDO = Trim(XDb.RSAux("APELIDO") & "")
         sNMPESSOA = Trim(XDb.RSAux("NMPESSOA") & "")
         nIDPESSOA = Val(XDb.RSAux("IDPESSOA") & "")
         If sAPELIDO = "" Then sAPELIDO = sNMPESSOA
         
         Set xNode = Me.TrwUsers.Nodes.Add("C0", tvwChild, "k" & nIDPESSOA, sNMPESSOA, "USER", "USER")
         
         Call SetTag(xNode, "NMMAQ", sNMMAQ)
         Call SetTag(xNode, "CODPESSOA", sCODPESSOA)
         Call SetTag(xNode, "APELIDO", sAPELIDO)
         Call SetTag(xNode, "NMPESSOA", sNMPESSOA)
         Call SetTag(xNode, "IDPESSOA", nIDPESSOA)
         XDb.RSAux.MoveNext
      Wend
      Set xNode = Me.TrwUsers.Nodes("C0")
      xNode.Text = GetTag(xNode, "CAPTION", "Conectados") & " (" & CStr(QtdOnLine) & ")"
   End If
   
   Set xNode = Me.TrwUsers.Nodes.Add(, , "D0", "Desconectados (0)", "CLOSE")
   xNode.ExpandedImage = "OPEN"
   xNode.Sorted = True
   xNode.Expanded = False
   xNode.ForeColor = vbBlue
   xNode.Bold = True
   Call SetTag(xNode, "CAPTION", "Desconectados")

   Sql = "Select DISTINCT P.CODPESSOA, P.APELIDO, P.NMPESSOA, P.IDPESSOA"
   Sql = Sql & " From PESSOA P, USUARIO U, USU_GRUPOS UG, GRPACESSO GA, GRPUSU_SISTEMA GS"
   Sql = Sql & " Where P.CODPESSOA = U.IDUSU"
   Sql = Sql & " And U.IDUSU = UG.IDUSU"
   Sql = Sql & " And UG.IDGRUPO= GA.IDGRUPO"
   Sql = Sql & " And GA.IDGRUPO=GS.IDGRUPO"
   Sql = Sql & " And GS.CODSIS = " & SqlStr("PMINFO")
   Sql = Sql & " And NOT U.IDUSU IN (SELECT DISTINCT IDUSU FROM CONEXAOMSG)"
   Sql = Sql & " Order By P.NMPESSOA"
   If XDb.AbreTabela(Sql) Then
      QtdOffLine = XDb.RSAux.RecordCount
      sNMMAQ = ""
      While Not XDb.RSAux.EOF
         sCODPESSOA = Trim(XDb.RSAux("CODPESSOA") & "")
         sAPELIDO = Trim(XDb.RSAux("APELIDO") & "")
         sNMPESSOA = Trim(XDb.RSAux("NMPESSOA") & "")
         nIDPESSOA = Val(XDb.RSAux("IDPESSOA") & "")
         If sAPELIDO = "" Then sAPELIDO = sNMPESSOA
         
         Set xNode = Me.TrwUsers.Nodes.Add("D0", tvwChild, "k" & nIDPESSOA, sNMPESSOA, "USER", "USER")
         
'         Call SetTag(xNode, "NMMAQ", sNMMAQ)
         Call SetTag(xNode, "CODPESSOA", sCODPESSOA)
         Call SetTag(xNode, "APELIDO", sAPELIDO)
         Call SetTag(xNode, "NMPESSOA", sNMPESSOA)
         Call SetTag(xNode, "IDPESSOA", nIDPESSOA)
         XDb.RSAux.MoveNext
      Wend
      Set xNode = Me.TrwUsers.Nodes("D0")
      xNode.Text = GetTag(xNode, "CAPTION", "Desconectados") & " (" & CStr(QtdOffLine) & ")"
   End If

End Sub
Private Sub ConectarMaquina(pRemoteName As String, pRemoteHost As String)
   If Trim(pRemoteHost) = "" Then
      pRemoteHost = UCase(InputBox("Entre com o nome ou IP da máquina.", "Conversa On-Line"))
   End If
   If Trim(pRemoteHost) = "" Then Exit Sub
   If CollCli Is Nothing Then Set CollCli = New Collection
   Set MCICLiente = New MCIMensagem
   With MCICLiente
      .Conexao = Cliente
      .LocalName = Sys.APELIDO
      
      .RemoteName = pRemoteName
      .RemoteHost = pRemoteHost
      .Porta = Sys.PortaPadrao
      
      
      .Top = Me.Top
      .Left = Me.Left + Me.Width
      .Height = Me.Height
      '.hWnd = Me.hWnd
      .Show
      CollCli.Add MCICLiente, "k" & MCICLiente.hWnd
   End With
   Set MCICLiente = Nothing

   Dim n       As Form
   Dim v       As MCIMensagem
   Dim bAchou  As Boolean
   For Each v In CollCli
      For Each n In Forms
         If CStr(n.hWnd) = CStr(v.hWnd) Then
            bAchou = True
            Exit For
         End If
      Next
      If Not bAchou Then
         If ExisteItem(CollCli, "k" & v.hWnd) Then
            CollCli.Remove "k" & v.hWnd
         End If
      End If
   Next
End Sub

Private Sub TrwUsers_DblClick()
   Call Conecta
End Sub
Private Sub Conecta()
   Dim sRemoteName   As String
   Dim sRemoteHost   As String
   Dim xNode         As Node
   
   Set xNode = Me.TrwUsers.SelectedItem
'   If Mid(xNode.Key, 1, 1) <> "k" Then
'      Exit Sub
'   End If
   
   sRemoteName = xNode.Text
   sRemoteHost = GetTag(xNode, "NMMAQ")
   If sRemoteName = "" Or sRemoteHost = "" Then
      Call ConectarMaquina("", "")
   Else
      Call ConectarMaquina(sRemoteName, sRemoteHost)
   End If
End Sub
Private Sub ConectarCOOL()
   Dim Sql As String
   
   If Trim(Sys.IDUSU) = "" Then
      Sys.IDUSU = UCase(InputBox("Entre com o seu código de usuário.", "Conversa On-Line"))
   End If
   
   Sql = "Select * "
   Sql = Sql & " From CONEXAOMSG "
   Sql = Sql & " Where IDUSU=" & SqlStr(Sys.IDUSU)
   If XDb.AbreTabela(Sql) Then
      'If XDb.RSAux("NMMAQ") <> MCIServidor.LocalHost Then
      If XDb.RSAux("NMMAQ") <> Me.WskMensagem.LocalHostName Then
         Call DesconectarCOOL
         Call ConectarCOOL
      End If
   Else
      Sql = "Insert Into CONEXAOMSG "
      Sql = Sql & "(IDUSU, NMMAQ) "
      Sql = Sql & " Values "
      Sql = Sql & "( " & SqlStr(Sys.IDUSU)
      Sql = Sql & ", " & SqlStr(Me.WskMensagem.LocalHostName)
      'Sql = Sql & ", " & SqlStr(MCIServidor.LocalHost)
      Sql = Sql & ")"
      Call XDb.Executa(Sql)
   End If
End Sub
Private Sub DesconectarCOOL()
   Dim Sql As String
   
   If Trim(Sys.IDUSU) = "" Then
      Sys.IDUSU = UCase(InputBox("Entre com o seu código de usuário.", "Conversa On-Line"))
   End If
   
   Sql = "Delete CONEXAOMSG "
   Sql = Sql & " Where IDUSU = " & SqlStr(Sys.IDUSU)
   Call XDb.Executa(Sql)
End Sub

Private Sub WskMensagem_ConnectionRequest(ByVal RequestID As Long)
   Set MCIServidor = New MCIMensagem
   With MCIServidor
      .LocalName = Sys.APELIDO

      .Conexao = MCIMConexao.Servidor
      .Porta = Sys.PortaPadrao
      .RequestID = RequestID
      .Show
   End With
   
   If CollSrv Is Nothing Then Set CollSrv = New Collection
   CollSrv.Add MCIServidor, CStr(MCIServidor.hWnd)
   Set MCIServidor = Nothing
End Sub
'**********************************************
'**********************************************
Option Explicit
Dim MCICLiente    As MCIMensagem
Dim MCIServidor   As MCIMensagem
Dim CollCli       As Collection
Dim CollSrv       As Collection
Private Sub CmdConectar_Click()
   Call Conecta
End Sub
Private Sub CmdSair_Click()
   Call Form_Unload(False)
End Sub
Private Sub Form_Initialize()
   On Error Resume Next
   Sys.PathUpdate = "\\Guarani\Sistemas\Admin"
   'Call AtualizaDLL(GetWindowsSysDir() & "MSWSOCK.DLL", Sys.PathUpdate & "\" & "MSWSOCK.DLL")
   Call AtualizaDLL(GetWindowsSysDir() & "MSWINSCK.OCX", Sys.PathUpdate & "\" & "MSWINSCK.OCX")
   Call ConectarBanco
End Sub
Private Sub Form_Load()
   With Me.WskMensagem
      .LocalPort = Sys.PortaPadrao
      .Listen
   End With
   
'   Set MCIServidor = New MCIMensagem
'   With MCIServidor
'      .LocalName = Sys.APELIDO
'
'      .Conexao = MCIMConexao.Servidor
'      .Porta = Sys.PortaPadrao
'      .Show
'      .Hide
'   End With
   DefineFundo "Azul"
   Me.Left = 0
   Me.Top = 300
   
   
   Call ConectarCOOL
   Call PopulaTrwUsers
End Sub

Private Sub Form_Resize()
   Me.TrwUsers.Left = 0
   Me.TrwUsers.Width = Me.Width - 120
   Me.TrwUsers.Height = Me.Height - (3 * Me.TrwUsers.Top)
   Me.FrmeTituloTree.Left = Me.TrwUsers.Left
   Me.FrmeTituloTree.Width = Me.TrwUsers.Width
   Me.CmdConectar.Top = Me.TrwUsers.Top + Me.TrwUsers.Height + 60
   Me.CmdSair.Top = Me.CmdConectar.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Dim n As Object
   On Error Resume Next
   For Each n In Forms
      Unload n
   Next
   Call DesconectarCOOL
   XDb.SrvDesconecta
   Set XDb = Nothing
   End
End Sub
Private Sub DefineFundo(Optional pCor As String = "Padrão")
   Select Case pCor
      Case "Azul":   Me.BackColor = &HE8D7CB
      Case "Padrão": Me.BackColor = &HD8E9EC
   End Select
End Sub
Private Sub ConectarBanco()
   Dim MySenha    As SENHA
   Set XDb = New DS_BANCO
   With XDb
      .SrvConecta , "MAUA", , "USU_VERIF", "DIPLOMATA", , , "TAMOIO"
      If Not .Conectado Then
         MsgBox "Não foi possível conectar servidor!", vbCritical, "Conversa On-Line"
         End
      End If
   End With
   If Sys.IDUSU = "" Then
      Set MySenha = New SENHA
      With MySenha
         Set .XDb = XDb
         .Show
         If .Conectado Then
            Sys.IDUSU = .IDUSU
         Else
            End
         End If
      End With
   End If
   
End Sub
Private Sub PopulaTrwUsers()
   Dim Sql        As String
   Dim xNode      As Node 'MSComctlLib
   Dim QtdOnLine  As Double
   Dim QtdOffLine As Double
   Dim sCOD       As String
   Dim sNMMAQ     As String
   Dim sCODPESSOA As String
   Dim sAPELIDO   As String
   Dim sNMPESSOA  As String
   Dim nIDPESSOA  As Double

   
   Me.TrwUsers.Nodes.Clear
   Set xNode = Me.TrwUsers.Nodes.Add(, , "C0", "Conectados (0)", "CLOSE")
   xNode.ExpandedImage = "OPEN"
   xNode.Sorted = True
   xNode.Expanded = True
   xNode.ForeColor = vbBlue
   xNode.Bold = True
   Call SetTag(xNode, "CAPTION", "Conectados")
   
   Sql = "Select C.NMMAQ, P.CODPESSOA, P.APELIDO, P.NMPESSOA, P.IDPESSOA"
   Sql = Sql & " From CONEXAOMSG C, PESSOA P"
   Sql = Sql & " Where C.IDUSU = P.CODPESSOA"
   Sql = Sql & " And C.IDUSU <>" & SqlStr(Sys.IDUSU)
   
   If XDb.AbreTabela(Sql) Then
      QtdOnLine = XDb.RSAux.RecordCount
      While Not XDb.RSAux.EOF
         sNMMAQ = Trim(XDb.RSAux("NMMAQ") & "")
         sCODPESSOA = Trim(XDb.RSAux("CODPESSOA") & "")
         sAPELIDO = Trim(XDb.RSAux("APELIDO") & "")
         sNMPESSOA = Trim(XDb.RSAux("NMPESSOA") & "")
         nIDPESSOA = Val(XDb.RSAux("IDPESSOA") & "")
         If sAPELIDO = "" Then sAPELIDO = sNMPESSOA
         
         Set xNode = Me.TrwUsers.Nodes.Add("C0", tvwChild, "k" & nIDPESSOA, sNMPESSOA, "USER", "USER")
         
         Call SetTag(xNode, "NMMAQ", sNMMAQ)
         Call SetTag(xNode, "CODPESSOA", sCODPESSOA)
         Call SetTag(xNode, "APELIDO", sAPELIDO)
         Call SetTag(xNode, "NMPESSOA", sNMPESSOA)
         Call SetTag(xNode, "IDPESSOA", nIDPESSOA)
         XDb.RSAux.MoveNext
      Wend
      Set xNode = Me.TrwUsers.Nodes("C0")
      xNode.Text = GetTag(xNode, "CAPTION", "Conectados") & " (" & CStr(QtdOnLine) & ")"
   End If
   
   Set xNode = Me.TrwUsers.Nodes.Add(, , "D0", "Desconectados (0)", "CLOSE")
   xNode.ExpandedImage = "OPEN"
   xNode.Sorted = True
   xNode.Expanded = False
   xNode.ForeColor = vbBlue
   xNode.Bold = True
   Call SetTag(xNode, "CAPTION", "Desconectados")

   Sql = "Select DISTINCT P.CODPESSOA, P.APELIDO, P.NMPESSOA, P.IDPESSOA"
   Sql = Sql & " From PESSOA P, USUARIO U, USU_GRUPOS UG, GRPACESSO GA, GRPUSU_SISTEMA GS"
   Sql = Sql & " Where P.CODPESSOA = U.IDUSU"
   Sql = Sql & " And U.IDUSU = UG.IDUSU"
   Sql = Sql & " And UG.IDGRUPO= GA.IDGRUPO"
   Sql = Sql & " And GA.IDGRUPO=GS.IDGRUPO"
   Sql = Sql & " And GS.CODSIS = " & SqlStr("PMINFO")
   Sql = Sql & " And NOT U.IDUSU IN (SELECT DISTINCT IDUSU FROM CONEXAOMSG)"
   Sql = Sql & " Order By P.NMPESSOA"
   If XDb.AbreTabela(Sql) Then
      QtdOffLine = XDb.RSAux.RecordCount
      sNMMAQ = ""
      While Not XDb.RSAux.EOF
         sCODPESSOA = Trim(XDb.RSAux("CODPESSOA") & "")
         sAPELIDO = Trim(XDb.RSAux("APELIDO") & "")
         sNMPESSOA = Trim(XDb.RSAux("NMPESSOA") & "")
         nIDPESSOA = Val(XDb.RSAux("IDPESSOA") & "")
         If sAPELIDO = "" Then sAPELIDO = sNMPESSOA
         
         Set xNode = Me.TrwUsers.Nodes.Add("D0", tvwChild, "k" & nIDPESSOA, sNMPESSOA, "USER", "USER")
         
'         Call SetTag(xNode, "NMMAQ", sNMMAQ)
         Call SetTag(xNode, "CODPESSOA", sCODPESSOA)
         Call SetTag(xNode, "APELIDO", sAPELIDO)
         Call SetTag(xNode, "NMPESSOA", sNMPESSOA)
         Call SetTag(xNode, "IDPESSOA", nIDPESSOA)
         XDb.RSAux.MoveNext
      Wend
      Set xNode = Me.TrwUsers.Nodes("D0")
      xNode.Text = GetTag(xNode, "CAPTION", "Desconectados") & " (" & CStr(QtdOffLine) & ")"
   End If

End Sub
Private Sub ConectarMaquina(pRemoteName As String, pRemoteHost As String)
   If Trim(pRemoteHost) = "" Then
      pRemoteHost = UCase(InputBox("Entre com o nome ou IP da máquina.", "Conversa On-Line"))
   End If
   If Trim(pRemoteHost) = "" Then Exit Sub
   If CollCli Is Nothing Then Set CollCli = New Collection
   Set MCICLiente = New MCIMensagem
   With MCICLiente
      .Conexao = Cliente
      .LocalName = Sys.APELIDO
      
      .RemoteName = pRemoteName
      .RemoteHost = pRemoteHost
      .Porta = Sys.PortaPadrao
      
      
      .Top = Me.Top
      .Left = Me.Left + Me.Width
      .Height = Me.Height
      '.hWnd = Me.hWnd
      .Show
      CollCli.Add MCICLiente, "k" & MCICLiente.hWnd
   End With
   Set MCICLiente = Nothing

   Dim n       As Form
   Dim v       As MCIMensagem
   Dim bAchou  As Boolean
   For Each v In CollCli
      For Each n In Forms
         If CStr(n.hWnd) = CStr(v.hWnd) Then
            bAchou = True
            Exit For
         End If
      Next
      If Not bAchou Then
         If ExisteItem(CollCli, "k" & v.hWnd) Then
            CollCli.Remove "k" & v.hWnd
         End If
      End If
   Next
End Sub

Private Sub TrwUsers_DblClick()
   Call Conecta
End Sub
Private Sub Conecta()
   Dim sRemoteName   As String
   Dim sRemoteHost   As String
   Dim xNode         As Node
   
   Set xNode = Me.TrwUsers.SelectedItem
'   If Mid(xNode.Key, 1, 1) <> "k" Then
'      Exit Sub
'   End If
   
   sRemoteName = xNode.Text
   sRemoteHost = GetTag(xNode, "NMMAQ")
   If sRemoteName = "" Or sRemoteHost = "" Then
      Call ConectarMaquina("", "")
   Else
      Call ConectarMaquina(sRemoteName, sRemoteHost)
   End If
End Sub
Private Sub ConectarCOOL()
   Dim Sql As String
   
   If Trim(Sys.IDUSU) = "" Then
      Sys.IDUSU = UCase(InputBox("Entre com o seu código de usuário.", "Conversa On-Line"))
   End If
   
   Sql = "Select * "
   Sql = Sql & " From CONEXAOMSG "
   Sql = Sql & " Where IDUSU=" & SqlStr(Sys.IDUSU)
   If XDb.AbreTabela(Sql) Then
      'If XDb.RSAux("NMMAQ") <> MCIServidor.LocalHost Then
      If XDb.RSAux("NMMAQ") <> Me.WskMensagem.LocalHostName Then
         Call DesconectarCOOL
         Call ConectarCOOL
      End If
   Else
      Sql = "Insert Into CONEXAOMSG "
      Sql = Sql & "(IDUSU, NMMAQ) "
      Sql = Sql & " Values "
      Sql = Sql & "( " & SqlStr(Sys.IDUSU)
      Sql = Sql & ", " & SqlStr(Me.WskMensagem.LocalHostName)
      'Sql = Sql & ", " & SqlStr(MCIServidor.LocalHost)
      Sql = Sql & ")"
      Call XDb.Executa(Sql)
   End If
End Sub
Private Sub DesconectarCOOL()
   Dim Sql As String
   
   If Trim(Sys.IDUSU) = "" Then
      Sys.IDUSU = UCase(InputBox("Entre com o seu código de usuário.", "Conversa On-Line"))
   End If
   
   Sql = "Delete CONEXAOMSG "
   Sql = Sql & " Where IDUSU = " & SqlStr(Sys.IDUSU)
   Call XDb.Executa(Sql)
End Sub

Private Sub WskMensagem_ConnectionRequest(ByVal RequestID As Long)
   Set MCIServidor = New MCIMensagem
   With MCIServidor
      .LocalName = Sys.APELIDO

      .Conexao = MCIMConexao.Servidor
      .Porta = Sys.PortaPadrao
      .RequestID = RequestID
      .Show
   End With
   
   If CollSrv Is Nothing Then Set CollSrv = New Collection
   CollSrv.Add MCIServidor, CStr(MCIServidor.hWnd)
   Set MCIServidor = Nothing
End Sub

