VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SENHA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mvarMe  As FrmSenha
Attribute mvarMe.VB_VarHelpID = -1

Private mvarxDb            As DS_BANCO
Private mvarConectado      As Boolean
Private mvarCancelado      As Boolean
Private mvarIDUSU          As String
'Public Property Let IDUSU(ByVal vData As String)
'   mvarIDUSU = vData
'End Property
Public Property Get IDUSU() As String
   IDUSU = mvarIDUSU
End Property
Public Property Let Conectado(ByVal vData As Boolean)
   mvarConectado = vData
End Property
Public Property Get Conectado() As Boolean
   Conectado = mvarConectado
End Property
Public Property Let Cancelado(ByVal vData As Boolean)
   mvarConectado = vData
End Property
Public Property Get Cancelado() As Boolean
   Cancelado = mvarCancelado
End Property
Public Property Set XDb(ByVal vData As Object)
   Set mvarxDb = vData
End Property
Public Property Get XDb() As Object
   Set XDb = mvarxDb
End Property
Public Sub Show()
   mvarMe.Show vbModal
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmSenha
   mvarCancelado = False
   mvarConectado = False
End Sub
Private Sub Class_Terminate()
   Set mvarMe = Nothing
End Sub
Private Sub mvarMe_Activate()
   If Not Splash Is Nothing Then
      If Splash.Flood.Visible Then
         Splash.LblMsg = ""
         Splash.Flood.Visible = False
      End If
      Splash.Flood.Visible = False
   End If
   If Trim(GetSetting(Sys.CODSIS, "Config", "UsuárioPadrão")) = "" Then
      mvarMe.TxtIDUSU.Text = UCase(NetworkUserID)
   Else
      mvarMe.TxtIDUSU.Text = GetSetting(Sys.CODSIS, "Config", "UsuárioPadrão")
   End If
   mvarMe.TxtSENHA.Text = ""
   mvarMe.TxtSENHA.SetFocus
   mvarMe.TxtSERVIDOR.Text = UCase(XDb.Server)
   mvarMe.TxtBANCO.Text = UCase(XDb.dbName)
   
   mvarMe.TxtSERVIDOR.Enabled = False
   mvarMe.TxtBANCO.Enabled = False
  
   DoEvents
   Screen.MousePointer = vbDefault
'   Call SetTopMostWindow(mvarMe.hWnd, True)
   'Forms(0).Top = 0
   'Forms(0).Left = 0
End Sub
Private Sub mvarMe_CmdCancelClick()
   mvarMe.CmdCancel.MousePointer = vbHourglass
   Screen.MousePointer = vbHourglass
   
   DoEvents
   If mvarConectado Then mvarMe.Hide
   mvarConectado = False
   mvarCancelado = True
   mvarIDUSU = ""
   Unload mvarMe
End Sub
Private Sub mvarMe_CmdOKClick()
   On Error GoTo TrataErro
   mvarMe.CmdOK.MousePointer = vbHourglass
   Screen.MousePointer = vbArrowHourglass
   mvarMe.LblError.Visible = False
   
   Call SplashFlood(10, "Conectando...")
   
   Call Conectar(mvarMe.TxtSERVIDOR.Text, mvarMe.TxtBANCO.Text, mvarMe.TxtIDUSU.Text, mvarMe.TxtSENHA.Text, True)
   If mvarConectado Then
      Call SaveConfig
   Else
      mvarMe.CmdOK.MousePointer = vbCustom
   End If
   Screen.MousePointer = vbDefault
   
   Exit Sub
TrataErro:
   MsgBox "5"
   Call UnloadIni
   Call ShowError
End Sub
Public Function Conectar(pServidor As String, pBanco As String, pUsuario As String, pSenha As String, pExibir As Boolean) As Boolean
   On Error GoTo TrataErro
   Screen.MousePointer = vbHourglass
   
   Call SplashFlood(20, "Conectando...")
   mvarConectado = True
   mvarCancelado = False
   
   If mvarxDb.Server <> pServidor And mvarxDb.Conectado Then
      Call mvarxDb.SrvDesconecta
   End If
   If mvarxDb.dbName <> pBanco And mvarxDb.Conectado Then
      Call mvarxDb.SrvDesconecta
   End If
   mvarxDb.Server = pServidor
   mvarxDb.dbName = pBanco
   
   Call SplashFlood(30, "Verificando Usuário..")
   If Not mvarxDb.Conectado Then
      Call mvarxDb.SrvConecta(ExibeErro:=False)
   End If

   If mvarxDb.Conectado Then
      If mvarxDb.AbreTabela("Select * From USUARIO Where IDUSU=" & SqlStr(pUsuario)) Then
         If pSenha = mvarxDb.RSAux("SENHAUSU") & "" Then
            If Trim(pUsuario) <> "" Then
               Call SaveSetting(Sys.CODSIS, "Config", "UsuárioPadrão", pUsuario)
            End If
            mvarConectado = True
            mvarIDUSU = mvarxDb.RSAux("IDUSU") & ""
                        
            Unload mvarMe
         Else
            If pExibir Then
               mvarMe.Visible = True
               Call SplashFlood(0, "", False)
               Call MsgBox("Senha Inválida!!", vbDefaultButton1 + vbInformation, "Segurança")
               mvarMe.TxtSENHA.SetFocus
            End If
         End If
      Else
         If pExibir Then
            Call SplashFlood(0, "", False)
            
            If mvarxDb.AbreTabela("Select * From USUARIO") Then
               mvarMe.Visible = True
               Call MsgBox("Usuário Inválido!!", vbDefaultButton1 + vbInformation, "Segurança")
               mvarMe.TxtIDUSU.SetFocus
            Else
               '* Caso não exista nenhum usuário cadastrado
               Call ShowCadUsuario
            End If
         End If
      End If
   End If
   Conectar = mvarConectado
   Screen.MousePointer = vbDefault
   Exit Function
TrataErro:
   Call UnloadIni
   MsgBox Error
   'Call ShowError
End Function
Private Sub mvarMe_KeyPress(KeyAscii As Integer)
   KeyAscii = SendTab(mvarMe, KeyAscii)
End Sub
Private Sub mvarMe_Load()
   Call MontaTela
   Call GetConfig
End Sub
Private Sub mvarMe_TxtBANCOGotFocus()
   Call SelecionarTexto(mvarMe.ActiveControl)
End Sub
Private Sub mvarMe_TxtBANCOKeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      Call mvarMe_CmdOKClick
   End If
End Sub
Private Sub mvarMe_TxtIDUSUGotFocus()
   Call SelecionarTexto(mvarMe.ActiveControl)
End Sub
Private Sub mvarMe_TxtSENHAGotFocus()
   Call SelecionarTexto(mvarMe.ActiveControl)
End Sub
Private Sub mvarMe_TxtSENHAKeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn And Not mvarMe.TxtSERVIDOR.Enabled And Not mvarMe.TxtBANCO.Enabled Then
      Call mvarMe_CmdOKClick
   End If
End Sub
Private Sub mvarMe_TxtSERVIDORGotFocus()
   Call SelecionarTexto(mvarMe.ActiveControl)
End Sub
Private Sub MontaTela()
'   Forms(0).Left = (Screen.Width - (Forms(0).Width + mvarMe.Width)) / 2
   mvarMe.Left = Forms(0).Left + Forms(0).Width
   mvarMe.Top = Forms(0).Top
   mvarMe.Moldura.Top = 20
   mvarMe.Moldura.Left = 20
   mvarMe.Moldura.Width = mvarMe.Width - 35
   mvarMe.Moldura.Height = mvarMe.Height - 35
   
   mvarMe.CmdOK.MousePointer = 99
   mvarMe.CmdCancel.MousePointer = 99
   On Error Resume Next
   mvarMe.CmdOK.MouseIcon = LoadResPicture("POINT", vbResCursor)
   mvarMe.CmdCancel.MouseIcon = mvarMe.CmdOK.MouseIcon
   
   mvarMe.ImgBanco.MousePointer = 99
   mvarMe.ImgServidor.MousePointer = 99
   mvarMe.ImgUsuário.MousePointer = 99
   mvarMe.ImgBanco.MouseIcon = mvarMe.CmdOK.MouseIcon
   mvarMe.ImgServidor.MouseIcon = mvarMe.CmdOK.MouseIcon
   mvarMe.ImgUsuário.MouseIcon = mvarMe.CmdOK.MouseIcon
End Sub
Private Sub ShowCadUsuario()
   Dim nForm      As Form
   mvarMe.Visible = False
   For Each nForm In Forms
      If UCase(nForm.Name) = UCase("FrmInicio") Then
         nForm.Visible = False
         Exit For
      End If
   Next
'   FrmCadPessoa.Show vbModal
   mvarMe.Visible = True
   If Not nForm Is Nothing Then
      nForm.Visible = True
   End If
End Sub
