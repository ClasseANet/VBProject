VERSION 5.00
Begin VB.Form FrmCALic 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  Licença"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   10770
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Editor de Chave"
      Height          =   2895
      Left            =   7800
      TabIndex        =   16
      Top             =   3000
      Visible         =   0   'False
      Width           =   10455
      Begin VB.TextBox TxtLICOrig 
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   10095
      End
      Begin VB.TextBox TxtTAG 
         Height          =   1695
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   600
         Width           =   10095
      End
      Begin VB.TextBox TxtLICDest 
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   2400
         Width           =   10095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   10455
      Begin VB.CommandButton CmdSalvar 
         Caption         =   "Salvar"
         Height          =   255
         Left            =   8880
         TabIndex        =   20
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox TxtPWD 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "#"
         TabIndex        =   11
         Top             =   2040
         Width           =   2175
      End
      Begin VB.CommandButton CmdConectar 
         Caption         =   "Conetar"
         Height          =   255
         Left            =   960
         TabIndex        =   12
         Top             =   2400
         Width           =   2055
      End
      Begin VB.ComboBox CmbColigadas 
         Height          =   315
         Left            =   3240
         TabIndex        =   13
         Text            =   "Combo1"
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox TxtTAGColigada 
         Height          =   1695
         Left            =   3240
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   600
         Width           =   6975
      End
      Begin VB.ComboBox CmbConexao 
         Height          =   315
         Left            =   960
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox TxtUID 
         Height          =   285
         Left            =   960
         TabIndex        =   9
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox TxtDbName 
         Height          =   285
         Left            =   960
         TabIndex        =   7
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox TxtServer 
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox TxtArquivo 
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "Senha:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Usuário:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Banco:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Servidor:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Conexão:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Referência:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmCALic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const SecaoBase = "Conection "
Dim gLocalReg As String
Dim gConexoes As Collection
Dim gXDb As Object

Private Sub CmbColigadas_Click()
   Dim Sql As String
   
   Me.TxtTAGColigada.Text = ""
   With gXDb
      Sql = "Select * "
      Sql = Sql & " From Coligada "
      Sql = Sql & " Where IDCOLIGADA=" & Me.CmbColigadas.ItemData(Me.CmbColigadas.ListIndex)
      If .Abretabela(Sql) Then
         Me.TxtTAGColigada.Text = Decrypt2(.RsAux("TAG") & "")
      End If
   End With

End Sub

Private Sub CmbConexao_Click()
   Me.TxtServer.Text = ""
   Me.TxtDbName.Text = ""
   Me.TxtUID.Text = ""
   Me.TxtPWD.Text = ""
   
   If Not gConexoes(Me.CmbConexao.Text) Is Nothing Then
      With gConexoes(Me.CmbConexao.Text)
         Me.TxtServer.Text = .Server
         Me.TxtDbName.Text = .DbName
         Me.TxtUID.Text = .UID
         Me.TxtPWD.Text = .PWD
      End With
   End If
End Sub
Private Sub ProcuraReferencia()
   Dim sArq As String
   Dim sPath As String
   
   If Trim(Me.TxtArquivo.Text) = "" Then
      sPath = App.Path
      'sPath = "C:\Arquivos de programas\ClasseA\Projeto3R"
      sPath = ResolvePathName(sPath)
      
      Me.TxtArquivo.Tag = sPath
      sArq = Dir(sPath & "*.reg", vbArchive)
      gLocalReg = sPath & sArq
   ElseIf Mid(Trim(Me.TxtArquivo.Text), 1, 3) = "C:\" Then
      gLocalReg = Me.TxtArquivo.Text
   Else
      gLocalReg = Me.TxtArquivo.Tag & Me.TxtArquivo.Text
   End If
   If Right(gLocalReg, 4) <> ".reg" Then
      gLocalReg = Me.TxtArquivo.Tag & Me.TxtArquivo.Text & ".reg"
   End If
   
   If gLocalReg <> "" Then
      sArq = GetNameFromPath(gLocalReg)
      Me.TxtArquivo.Text = Mid(sArq, 1, Len(sArq) - 4)
      Call MontaConexoes(gLocalReg)
   End If
   
End Sub
Private Sub MontaConexoes(pLocalReg As String)
   Dim i       As Integer
   Dim sALIAS  As String
   Dim MyxDb   As Object
   Dim bObriga As Boolean
   
   Set gConexoes = Nothing
   Set gConexoes = New Collection

   If gConexoes.Count > 0 Then Exit Sub
      
   i = 0
   sALIAS = ReadIniFile(pLocalReg, SecaoBase & i, "ALIAS", "")
   Me.CmbConexao.Clear
   While sALIAS <> ""
      Me.CmbConexao.AddItem sALIAS, i
      
      Set MyxDb = Nothing
      Set MyxDb = CreateObject("XBANCO01.DS_BANCO")
      Call RegistroParaObjeto(MyxDb, sALIAS)
      If MyxDb.Alias <> "" Then
         If ExisteItem(gConexoes, MyxDb.Alias) Then
            gConexoes.Remove MyxDb.Alias
         End If
         gConexoes.Add MyxDb, MyxDb.Alias
      End If
      Set MyxDb = Nothing
      i = i + 1
            
      If bObriga Then
         'mvarConn.MCIMenu1.Enabled(mvarConn.MCIMenu1.Count) = False
         sALIAS = ""
      Else
         sALIAS = Trim(ReadIniFile(gLocalReg, SecaoBase & i, "ALIAS", ""))
      End If
   Wend
   If Me.CmbConexao.ListCount = 1 Then
      Me.CmbConexao.ListIndex = 0
      Me.CmbConexao.Enabled = False
   End If
   
   Set MyxDb = Nothing
End Sub

Private Sub CmdConectar_Click()
   Dim Sql As String
   
   Me.CmbColigadas.Clear
   Me.TxtTAGColigada.Text = ""
   Me.CmbColigadas.Enabled = False
   Me.TxtTAGColigada.Enabled = False
   
   
   Set gXDb = CriarObjeto("xBANCO01.DS_BANCO")
   With gXDb
      .Server = Me.TxtServer.Text
      .DbName = Me.TxtDbName.Text
      .UID = Me.TxtUID.Text
      .PWD = Me.TxtPWD.Text
      
      .SrvConecta
      
      If .Conectado Then
         Sql = "Select * From Coligada"
         If .Abretabela(Sql) Then
            Me.CmbColigadas.Enabled = True
            Me.TxtTAGColigada.Enabled = True
            While Not .RsAux.EOF
                Me.CmbColigadas.AddItem .RsAux("NMCOLIGADA")
                Me.CmbColigadas.ItemData(Me.CmbColigadas.NewIndex) = .RsAux("IDCOLIGADA")
               .RsAux.movenext
            Wend
            If Me.CmbColigadas.ListCount = 1 Then
               Me.CmbColigadas.ListIndex = 0
               Me.CmbColigadas.Enabled = False
            End If
         End If
      End If
   End With


End Sub

Private Sub CmdSalvar_Click()
   Dim Sql As String
   If Me.CmbColigadas.Text <> "" Then
      If InputBoxPassword("Entre com a Senha") = "dolphin" Then
         Sql = "Update COLIGADA "
         Sql = Sql & " Set TAG=" & SqlStr(Encrypt2(Me.TxtTAGColigada.Text))
         Sql = Sql & " Where IDCOLIGADA = " & SqlNum(Me.CmbColigadas.ItemData(Me.CmbColigadas.ListIndex))
         If gXDb.Executa(Sql) Then
            Call ExibirInformacao("Operação realizada com sucesso.", "Licença")
         Else
            Call ExibirStop("Erro ao salvar.", "Licença")
         End If
      End If
   End If
End Sub

Private Sub Form_Activate()
   Me.TxtArquivo.SetFocus
End Sub

Private Sub Form_Initialize()
   If Not InputBoxPassword("Entre com a Senha") = "dolphin" Then
      End
   End If
End Sub

Private Sub Form_Load()
   Me.Frame1.Top = 120
   Me.Frame1.Left = Me.Frame1.Top
   Me.Frame2.Top = Me.Frame1.Top
   Me.Frame2.Left = Me.Frame1.Top
   Me.Frame2.Visible = False
   
   Call ProcuraReferencia
   
   
End Sub

Private Sub Frame1_DblClick()
   If InputBoxPassword("Entre com a Senha") = "dolphin" Then
      Me.Frame1.Visible = False
      Me.Frame2.Visible = True
   End If
End Sub

Private Sub Frame2_DblClick()
   Me.Frame1.Visible = True
   Me.Frame2.Visible = False
End Sub

Private Sub TxtArquivo_Change()
   Me.TxtDbName.Text = ""
   Me.TxtLICDest.Text = ""
   Me.TxtLICOrig.Text = ""
   Me.TxtPWD.Text = ""
   Me.TxtServer.Text = ""
   Me.TxtTAG.Text = ""
   Me.TxtTAGColigada.Text = ""
   Me.TxtUID.Text = ""
   
   Me.CmbColigadas.Clear
   Me.CmbConexao.Clear
   Me.CmbColigadas.Enabled = False
   Me.TxtTAGColigada.Enabled = False
End Sub

Private Sub TxtArquivo_GotFocus()
   Call TxtArquivo_Change
End Sub

Private Sub TxtArquivo_LostFocus()
   Call ProcuraReferencia
End Sub

Private Sub TxtLICOrig_LostFocus()
   Me.TxtTAG.Text = Decrypt2(Me.TxtLICOrig.Text)
End Sub
Private Sub TxtTAG_LostFocus()
   Me.TxtLICDest.Text = Encrypt2(Me.TxtTAG.Text)
End Sub
Private Sub RegistroParaObjeto(ByRef MyxDb As Object, pAlias As String)
   Dim sSecao  As String
   
   sSecao = GetSecao(pAlias)
   If sSecao = "" Then Exit Sub
   With MyxDb
      .Alias = ReadIniFile(gLocalReg, sSecao, "ALIAS")
      .isODBC = ReadIniFile(gLocalReg, sSecao, "isODBC", False)
      .DbTipo = ReadIniFile(gLocalReg, sSecao, "DBTIPO", 1)
      .dbVersao = ReadIniFile(gLocalReg, sSecao, "DBVERSAO")
      .isADO = ReadIniFile(gLocalReg, sSecao, "isADO", True)
      If MyxDb.DbTipo = 0 Then
         .dbDrive = ReadIniFile(gLocalReg, sSecao, "DBDRIVE")
      Else
         .Server = ReadIniFile(gLocalReg, sSecao, "SERVER")
      End If
      .DbName = ReadIniFile(gLocalReg, sSecao, "DBNAME")
      .UID = ReadIniFile(gLocalReg, sSecao, "UID")
      .PWD = Decrypt2(ReadIniFile(gLocalReg, sSecao, "PWD"))
   End With
End Sub
Private Function GetSecao(pAlias As String) As String
   Dim i       As Integer
   Dim j       As Integer
   Dim sALIAS  As String
   Dim bAchou  As Boolean
   On Error GoTo TrataErro
   i = 0
   j = 0
   sALIAS = ReadIniFile(gLocalReg, SecaoBase & i, "Alias")
   While UCase(pAlias) <> UCase(sALIAS) And j <> 2
      If sALIAS = "" Then j = j + 1
      i = i + 1
      sALIAS = Trim(ReadIniFile(gLocalReg, SecaoBase & i, "Alias", ""))
      If sALIAS <> "" Then j = 0
   Wend
   GetSecao = SecaoBase & i - j
Exit Function
TrataErro:
   ShowError "Splash.GetSecao"
End Function

