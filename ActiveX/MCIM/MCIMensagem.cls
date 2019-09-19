VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MCIMensagem"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mvarMe  As FrmMensagem
Attribute mvarMe.VB_VarHelpID = -1
Private WithEvents mvarMDI As FrmInicio
Attribute mvarMDI.VB_VarHelpID = -1

Public Enum MCIMConexao
   Servidor
   Cliente
End Enum

Private mvarConexao        As MCIMConexao
Private mvarLocalHost      As String
Private mvarRemoteHost     As String
Private mvarLocalName      As String
Private mvarRemoteName     As String
Private mvarPorta          As Double
Private mvarMensagem       As String
Private mvarTop            As Double
Private mvarLeft           As Double
Private mvarHeight         As Double
Private mvarhWnd           As Double
Private mvarRequestID      As Long
Private mvarisConferencia  As Boolean

Dim kServidor              As Integer
Dim kCliente               As Integer
Dim CollServidores         As Collection

'* Propriedades Usuadas exclusivamente pelo MDI
Dim MCICLiente    As MCIMensagem
Dim MCIServidor   As MCIMensagem
Dim CollCli       As Collection
Dim CollSrv       As Collection
Public Property Let isConferencia(ByVal vData As Boolean)
   mvarisConferencia = vData
End Property
Public Property Get isConferencia() As Boolean
   isConferencia = mvarisConferencia
End Property
Public Property Let RequestID(ByVal vData As Long)
   mvarRequestID = vData
End Property
Public Property Get RequestID() As Long
   RequestID = mvarRequestID
End Property
Public Property Get hWnd() As Double
   mvarhWnd = mvarMe.hWnd
   hWnd = mvarhWnd
End Property
Public Property Let Top(ByVal vData As Double)
   mvarTop = vData
End Property
Public Property Get Top() As Double
   Top = mvarTop
End Property
Public Property Let Left(ByVal vData As Double)
   mvarLeft = vData
End Property
Public Property Get Left() As Double
   Left = mvarLeft
End Property
Public Property Let Height(ByVal vData As Double)
   If vData < 6450 Then mvarHeight = 6450
   mvarHeight = vData
End Property
Public Property Get Height() As Double
   Height = mvarHeight
End Property
Public Property Get Mensagem() As String
   Mensagem = mvarMensagem
End Property
Public Property Let Porta(ByVal vData As Double)
   mvarPorta = vData
End Property
Public Property Get Porta() As Double
   Porta = mvarPorta
End Property
Public Property Let Conexao(ByVal vData As MCIMConexao)
   mvarConexao = vData
End Property
Public Property Get Conexao() As MCIMConexao
   Conexao = mvarConexao
End Property
'Public Property Let LocalHost(ByVal vData As String)
'   Set mvarLocalHost = vData
'End Property
Public Property Get LocalHost() As String
   LocalHost = mvarLocalHost
End Property
Public Property Let RemoteHost(ByVal vData As String)
   mvarRemoteHost = vData
End Property
Public Property Get RemoteHost() As String
   RemoteHost = mvarRemoteHost
End Property
Public Property Let LocalName(ByVal vData As String)
   mvarLocalName = vData
End Property
Public Property Get LocalName() As String
   LocalName = mvarLocalName
End Property
Public Property Let RemoteName(ByVal vData As String)
   mvarRemoteName = vData
End Property
Public Property Get RemoteName() As String
   RemoteName = mvarRemoteName
End Property
Private Sub Class_Initialize()
   Set mvarMe = New FrmMensagem
   Set mvarMDI = New FrmInicio
   mvarHeight = 6450
End Sub
Private Sub Class_Terminate()
   On Error Resume Next
   Unload mvarMe
   Set mvarMe = Nothing
   Set CollServidores = Nothing
End Sub
Public Sub Show()
   mvarMe.Show
End Sub
Public Sub Hide()
   mvarMe.Hide
End Sub
Private Sub mvarMe_CmdEnviarClick()
   If mvarConexao = Servidor Then
      Call Enviar(kServidor)
   ElseIf mvarConexao = Cliente Then
      Call Enviar(kCliente)
   End If
End Sub
Private Sub Enviar(Index As Integer)
   If Trim(mvarMe.TxtEnviar.Text) = "" Then Exit Sub
   
   Screen.MousePointer = vbHourglass
   
   mvarMe.TxtEnviar.Text = EliminarString(mvarMe.TxtEnviar.Text, vbNewLine)
   mvarMensagem = Trim(mvarMe.TxtEnviar.Text)
   With mvarMe.WskMensagem(Index)
      If Not .State = sckConnected Then
         Call Conectar
      End If
      If Not .State = sckConnected Then
'         If MsgBox("Você não esta conectado ainda . Deseja conectar a " & mvarRemoteName & " ? ", vbYesNo + vbQuestion) = vbYes Then
'            Call Conectar
'         Else
            mvarMe.LblAviso.Caption = "Desconectado!"
            mvarMe.LblEspera.Caption = ""
'            MsgBox "A conexão não foi realizada !", vbInformation, "Usando Winsock"
'         End If
      End If
      
      If .State = sckConnected Then
         Call SendData(Index)
         Call Exibir_Mensagem_Chat(mvarLocalHost, mvarLocalName, mvarMensagem)
         mvarMe.TxtEnviar.Text = ""
         mvarMe.TxtEnviar.SetFocus
      End If
      DoEvents
'      .Close
   End With
   SendKeys "^{End}"
   Screen.MousePointer = vbDefault
End Sub
Private Sub Conectar()
   Dim Hr1 As Double
   Dim i    As Integer
      
   mvarMe.LblAviso.Visible = True
   With mvarMe.WskMensagem(0)
      .Close
      .RemotePort = mvarPorta
      '.RemotePort = 1224
      .RemoteHost = mvarRemoteHost
      .Connect
      
      Hr1 = Int(Timer)
      While .State <> StateConstants.sckConnected
         DoEvents
         If mvarConexao = Cliente Then
            If .State = StateConstants.sckError Or (Int(Timer) - Hr1 > 10) Then
'               mvarPorta = mvarPorta + 1
'               Call Conectar
'               Exit Sub
            End If
         End If
         Select Case .State
            Case StateConstants.sckClosed: mvarMe.LblAviso = "Conexão Fechada!"
            Case StateConstants.sckClosing: mvarMe.LblAviso = "Fechando Conexão!"
            Case StateConstants.sckConnected: mvarMe.LblAviso = "Conetado!"
            Case StateConstants.sckConnecting: mvarMe.LblAviso = "Conectando..."
            Case StateConstants.sckConnectionPending: mvarMe.LblAviso = "Conexão Pendente..."
            Case StateConstants.sckError: mvarMe.LblAviso = "Erro!!"
            Case StateConstants.sckHostResolved: mvarMe.LblAviso = "Servidor encontrado!"
            Case StateConstants.sckListening: mvarMe.LblAviso = "Esperando Resposta..."
            Case StateConstants.sckOpen: mvarMe.LblAviso = "Conexão Aberta"
            Case StateConstants.sckResolvingHost: mvarMe.LblAviso = "Procurando Servidor..."
         End Select
         
         If .State <> StateConstants.sckConnected Then
            mvarMe.LblEspera.Left = mvarMe.LblAviso.Left + mvarMe.LblAviso.Width + 20
            mvarMe.LblEspera = Mid("..........", 1, i)
            i = i + 1
            If i = 11 Then i = 0
         End If
         If .State <> StateConstants.sckConnected And .State <> StateConstants.sckConnecting And .State <> StateConstants.sckResolvingHost Then
            If .State = StateConstants.sckError Or (Int(Timer) - Hr1 > 10) Then
               Call Exibir_Mensagem_Chat("", "Sistema", "Problemas com a conexão !")
'               MsgBox "Há problemas com a conexão !"
               Exit Sub
            End If
         End If
      Wend
   End With
   mvarMe.LblCliente.Caption = mvarLocalName
   'mvarMe.Caption = "Usando Winsock - Cliente " & mvarLocalName
   'mvarMe.LblConexao.Caption = " Conectado "
End Sub
Private Sub mvarMe_Load()
   Dim Tentativas  As Integer
  
   On Error GoTo Trata_Erro
   kServidor = 0
   kCliente = 0
   With mvarMe.WskMensagem(0)
      If mvarConexao = Servidor Then
         If mvarRequestID <> 0 Then
            .LocalPort = 0
            If .State <> sckClosed Then
               .Close
            End If
            .Accept RequestID
         Else
            .LocalPort = mvarPorta
            '.Bind mvarPorta, .LocalIP
            .Listen
         End If
         mvarMe.Caption = mvarMe.Caption & " (s)"
         
      ElseIf mvarConexao = Cliente Then
         .RemotePort = mvarPorta
         .RemoteHost = mvarRemoteHost
         mvarMe.Caption = mvarMe.Caption & " (c)"
      End If
      
      mvarLocalHost = mvarMe.WskMensagem(0).LocalHostName
   End With
   mvarMe.LblAviso.Caption = "Conectado."
   mvarMe.LblEspera.Caption = ""
   mvarMe.LblCliente.Caption = mvarLocalName
   mvarMe.LblLocalHost.Caption = "<" & mvarLocalHost & ">"
   
   mvarMe.LblServidor.Caption = mvarRemoteName
   mvarMe.LblRemoteHost.Caption = "<" & mvarRemoteHost & ">"
   
   DefineFundo '"Azul"
   
   mvarMe.Top = mvarTop
   mvarMe.Left = mvarLeft
   mvarMe.Height = mvarHeight
      
   Exit Sub
Trata_Erro:
   Tentativas = Tentativas + 1
   If Tentativas > 10 Then
      MsgBox " Outra aplicação esta usando a porta " & mvarMe.WskMensagem(0).LocalPort, vbCritical, "Usando Winsock "
   Else
      mvarPorta = mvarPorta + 1
      mvarMe.WskMensagem(0).LocalPort = mvarPorta
      Resume
   End If
   Unload mvarMe
End Sub
Private Sub mvarMe_Resize()
   Call AjustaNomes
   With mvarMe
      .TxtReceber.Left = 120
   End With
End Sub
Private Sub mvarMe_Timer01()
   Dim i             As Integer
   Static ConnHide   As Integer
   
   With mvarMe
      ConnHide = ConnHide + 1
      If ConnHide >= 60 Then
         If .Visible = False Then
            For i = .WskMensagem.LBound To .WskMensagem.UBound
               .WskMensagem(i).Close
            Next
         End If
         ConnHide = 0
      End If
   End With
End Sub
Private Sub mvarMe_TxtEnviarKeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      DoEvents
      KeyCode = 0
      If Asc(mvarMe.TxtEnviar.Text) = vbKeyReturn Then mvarMe.TxtEnviar.Text = ""
      
      If mvarConexao = Servidor Then
         Call Enviar(kServidor)
      ElseIf mvarConexao = Cliente Then
         Call Enviar(kCliente)
      End If
   End If
End Sub
Private Sub mvarMe_Unload(Cancel As Integer)
   Dim i             As Integer
   Dim bPortaAberta  As Boolean

   bPortaAberta = False
   For i = mvarMe.WskMensagem.LBound To mvarMe.WskMensagem.UBound
      If mvarMe.WskMensagem(i).State <> sckClosed Then
         bPortaAberta = True
         Exit For
      End If
   Next
   If bPortaAberta Then
      Cancel = True
      mvarMe.Hide
   End If
End Sub
Private Sub mvarMe_WskMensagemConnectionRequest(Index As Integer, ByVal RequestID As Long)
   Dim ClServidor As MCIMensagem
   If Index = 0 Then
      If CollServidores Is Nothing Then Set CollServidores = New Collection
      
      If isConferencia Then
         kServidor = kServidor + 1
         Load mvarMe.WskMensagem(kServidor)
      
         mvarMe.WskMensagem(kServidor).LocalPort = 0
         If mvarMe.WskMensagem(kServidor).State <> sckClosed Then
            mvarMe.WskMensagem(kServidor).Close
         End If
         mvarMe.WskMensagem(kServidor).Accept RequestID
      Else
         Set ClServidor = New MCIMensagem
         ClServidor.RequestID = RequestID
         ClServidor.Show
         CollServidores.Add ClServidor, CStr(ClServidor.hWnd)
      End If
   End If
'   If mvarMe.WskMensagem(0).State <> sckClosed Then
'      mvarMe.WskMensagem(0).Close
'   End If
'   mvarMe.WskMensagem(0).Accept requestID
End Sub

Private Sub mvarMe_WskMensagemDataArrival(Index As Integer, ByVal bytesTotal As Long)
   Dim StrDados As String
   
   If Not mvarMe.Visible Then mvarMe.Show
   
   With mvarMe.WskMensagem(Index)
      If Not .State = sckConnected Then
         MsgBox "Não existe conexão ativa , aguarde ... "
         If MsgBox("Você não esta conectado ainda . Deseja conectar a " & mvarRemoteName & " ? ", vbYesNo + vbQuestion) = vbYes Then
            Call Conectar
         Else
            MsgBox "A conexão não foi realizada !", vbInformation, "Usando Winsock"
         End If
      End If
         
      If .State = sckConnected Then
         Call GetData(Index)
         If mvarMe.LblRemoteHost.Caption = "<>" Then
            mvarMe.LblServidor.Caption = mvarRemoteName
            mvarMe.LblRemoteHost.Caption = "<" & mvarRemoteHost & ">"
            Call AjustaNomes
            mvarMe.Show
         End If
         Call Exibir_Mensagem_Chat(mvarRemoteHost, mvarRemoteName, mvarMensagem)
      End If
   End With
End Sub
Private Sub Exibir_Mensagem_Chat(Maquina As String, Remetente As String, ByRef Mensagem As String)
   Dim MyCtrl  As Object
'   Set MyCtrl = mvarMe.ActiveControl
   
   mvarMe.TxtReceber.Text = mvarMe.TxtReceber.Text & Remetente & " diz : " & vbNewLine
   mvarMe.TxtReceber.Text = mvarMe.TxtReceber.Text & "   " & Mensagem & vbNewLine
   Mensagem = ""
  ' mvarMe.TxtReceber.SetFocus
   mvarMe.TxtReceber.SelStart = Len(mvarMe.TxtReceber.Text)
   'SendKeys "^{End}"
   
   On Error Resume Next
'   MyCtrl.SetFocus
End Sub
Private Sub SendData(Index As Integer)
   Dim sMsg As String
   
   sMsg = ""
   Call SetTag(sMsg, "TYPE", "MESSAGE")
   Call SetTag(sMsg, "HOST", mvarLocalHost)
   Call SetTag(sMsg, "NAME", mvarLocalName)
   Call SetTag(sMsg, "MESSAGE", mvarMensagem)
   Call SetTag(sMsg, "DATE", Now)
   mvarMe.WskMensagem(Index).SendData sMsg
End Sub
Private Sub GetData(Index As Integer)
   Dim sMsg As String
   
   mvarMe.WskMensagem(Index).GetData sMsg
   
   mvarMensagem = ""
   If GetTag(sMsg, "TYPE") = "MESSAGE" Then
      mvarRemoteHost = GetTag(sMsg, "HOST", mvarRemoteHost)
      mvarRemoteName = GetTag(sMsg, "NAME", mvarRemoteName)
      mvarMensagem = GetTag(sMsg, "MESSAGE", "")
   End If
End Sub
Private Sub AjustaNomes()
   With mvarMe
      .LblPara.Left = mvarMe.TxtReceber.Left + 120
      .LblServidor.Left = .LblPara.Left + .LblPara.Width + 60
      .LblRemoteHost.Left = .LblServidor.Left + .LblServidor.Width + 60
      
      .LblDe.Left = mvarMe.TxtEnviar.Left + 120
      .LblCliente.Left = .LblDe.Left + .LblDe.Width + 60
      .LblLocalHost.Left = .LblCliente.Left + .LblCliente.Width + 60
   End With
End Sub
Private Sub DefineFundo(Optional pCor As String = "Padrão")
   Select Case pCor
      Case "Azul":   mvarMe.BackColor = &HE8D7CB
      Case "Padrão": mvarMe.BackColor = &HD8E9EC
   End Select
End Sub

