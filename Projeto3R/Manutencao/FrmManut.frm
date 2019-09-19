VERSION 5.00
Begin VB.Form FrmManut 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  Manutenção"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   10365
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSair 
      Caption         =   "Sair"
      Height          =   495
      Left            =   8640
      TabIndex        =   31
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton CmdPause 
      Caption         =   "Parar"
      Height          =   495
      Left            =   120
      TabIndex        =   30
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Atualizar Id. Loja"
      Height          =   2895
      Left            =   120
      TabIndex        =   17
      Top             =   3120
      Width           =   9975
      Begin VB.TextBox TxtIDCOLIGADA2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         TabIndex        =   21
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox TxtIDCOLIGADA1 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         TabIndex        =   20
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox TxtLogIDLOJANot 
         Height          =   2325
         IMEMode         =   3  'DISABLE
         Left            =   6120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Top             =   360
         Width           =   3615
      End
      Begin VB.CommandButton CmdAtualizar 
         Caption         =   "Atualizar"
         Height          =   375
         Left            =   600
         TabIndex        =   18
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox TxtLogIDLOJA 
         Height          =   2325
         IMEMode         =   3  'DISABLE
         Left            =   2280
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Top             =   360
         Width           =   3615
      End
      Begin VB.TextBox TxtIDLOJA1 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         TabIndex        =   23
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox TxtIDLOJA2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         TabIndex        =   24
         Top             =   1800
         Width           =   855
      End
      Begin VB.CommandButton CmdExeIDLOJA 
         Caption         =   "Executar"
         Height          =   375
         Left            =   600
         TabIndex        =   25
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Não Ok."
         Height          =   255
         Left            =   6120
         TabIndex        =   28
         Top             =   120
         Width           =   855
      End
      Begin VB.Label LblObjetos 
         Caption         =   "Objetos: "
         Height          =   255
         Left            =   2280
         TabIndex        =   26
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Loja"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Coligada"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9975
      Begin VB.CommandButton CmdAtualizaLic 
         Caption         =   "Atualizar"
         Height          =   375
         Left            =   4080
         TabIndex        =   32
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CommandButton CmdSalvar 
         Caption         =   "Salvar"
         Height          =   375
         Left            =   8400
         TabIndex        =   16
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox TxtPWD 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "#"
         TabIndex        =   12
         Top             =   2040
         Width           =   2175
      End
      Begin VB.CommandButton CmdConectar 
         Caption         =   "Conetar"
         Height          =   375
         Left            =   960
         TabIndex        =   13
         Top             =   2400
         Width           =   2055
      End
      Begin VB.ComboBox CmbColigadas 
         Height          =   315
         Left            =   3240
         TabIndex        =   14
         Text            =   "Combo1"
         Top             =   240
         Width           =   6495
      End
      Begin VB.TextBox TxtTAGColigada 
         Height          =   1695
         Left            =   3240
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   600
         Width           =   6495
      End
      Begin VB.ComboBox CmbConexao 
         Height          =   315
         Left            =   960
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox TxtUID 
         Height          =   285
         Left            =   960
         TabIndex        =   10
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox TxtDbName 
         Height          =   285
         Left            =   960
         TabIndex        =   8
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox TxtServer 
         Height          =   285
         Left            =   960
         TabIndex        =   6
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox TxtArquivo 
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "Senha:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Usuário:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Banco:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Servidor:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Conexão:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Referência:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmManut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const SecaoBase = "Conection "
Dim gLocalReg As String
Dim gConexoes As Collection
Dim gXDb As Object

Dim bIDLOJA As Boolean
Dim bPause As Boolean
Private Sub CmbColigadas_Click()
   Call ExibirLicenca
End Sub
Private Sub CmdAtualizaLic_Click()
   Screen.MousePointer = vbHourglass
   Call ExibirLicenca
   Screen.MousePointer = vbDefault
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
      'sPath = App.Path
      sPath = ResolvePathName(Environ("programfiles") & "\ClasseA\Projeto3R\")
      
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
Private Sub CmdAtualizar_Click()
   Dim Sql As String
   
   Screen.MousePointer = vbHourglass
   Sql = "Select * From OLOJA"
   If gXDb.Abretabela(Sql) Then
      Me.TxtIDLOJA1.Text = gXDb.RsAux("IDLOJA")
      Me.TxtIDLOJA2.Text = gXDb.RsAux("IDLOJA") + 1
      Me.TxtIDLOJA2.SetFocus
      
      If xVal(gXDb.RsAux("IDLOJA") & "") = 0 Then
         Sql = "Select * From COLIGADA"
         Call gXDb.Abretabela(Sql)
      End If
      Me.TxtIDCOLIGADA1.Text = gXDb.RsAux("IDCOLIGADA")
      Me.TxtIDCOLIGADA2.Text = gXDb.RsAux("IDCOLIGADA") + 1
      
   End If
   
   Sql = "SELECT O.NAME [TABELA], O.*"
   Sql = Sql & " FROM SYSCOLUMNS C JOIN SYSOBJECTS O ON C.ID=O.ID"
   Sql = Sql & " WHERE C.NAME = 'IDLOJA'"
   Sql = Sql & " ORDER BY O.CRDATE"
   Me.LblObjetos = "Objetos: 0"
   If gXDb.Abretabela(Sql) Then
      Me.LblObjetos = "Objetos: " & xVal(gXDb.RsAux.Recordcount)
      bIDLOJA = True
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub CmdConectar_Click()
   Dim Sql As String
   
   Screen.MousePointer = vbHourglass
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
      
      If .conectado Then
         Sql = "Select * From Coligada"
         If .Abretabela(Sql) Then
            Me.CmbColigadas.Enabled = True
            Me.TxtTAGColigada.Enabled = True
            While Not .RsAux.EOF
                Me.CmbColigadas.AddItem .RsAux("NMCOLIGADA")
                Me.CmbColigadas.ItemData(Me.CmbColigadas.NewIndex) = .RsAux("IDCOLIGADA")
               .RsAux.MoveNext
            Wend
            If Me.CmbColigadas.ListCount = 1 Then
               Me.CmbColigadas.ListIndex = 0
               Me.CmbColigadas.Enabled = False
            End If
         End If
      End If
   End With
   Screen.MousePointer = vbDefault

End Sub
Private Sub ExibirLicenca()
   Dim Sql As String
   
   Me.TxtTAGColigada.Text = ""
   With gXDb
      If Me.CmbColigadas.ListIndex >= 0 Then
         Sql = "Select * From Coligada Where IDCOLIGADA=" & Me.CmbColigadas.ItemData(Me.CmbColigadas.ListIndex)
         If .Abretabela(Sql) Then
            Me.TxtTAGColigada.Text = Decrypt2(.RsAux("TAG"))
         End If
      End If
   End With
End Sub
Private Sub CmdExeIDLOJA_Click()
   Dim Sql As String
'   Dim RsTab As Object
'   Dim MyTab As Object
'   Dim i As Integer
'   Dim j As Integer
'   Dim kNOk As Integer
'   Dim sTabela As String
'   Dim nPos As Long
   
   On Error GoTo TrataErro
   Screen.MousePointer = vbHourglass
   
   If xVal(Me.TxtIDLOJA2.Text) = 0 Then
      MsgBox "Id Inválido"
      Me.TxtIDLOJA2.SetFocus
      GoTo Saida
   End If
   If xVal(Me.TxtIDCOLIGADA2.Text) = 0 Then
      MsgBox "Id Inválido"
      Me.TxtIDCOLIGADA2.SetFocus
      GoTo Saida
   End If
   If gXDb Is Nothing Then
      MsgBox "Conexão Inválida"
      Me.CmdConectar.SetFocus
      GoTo Saida
   Else
      If Not gXDb.conectado Then
         MsgBox "Banco desconectado"
         Me.CmdConectar.SetFocus
         GoTo Saida
      End If
   End If
   
   Sql = "Select * "
   Sql = Sql & " From OLOJA"
   Sql = Sql & " Where IDLOJA= " & SqlNum(Me.TxtIDLOJA2.Text)
   Sql = Sql & " Order By IDLOJA"
   If gXDb.Abretabela(Sql) And (xVal(Me.TxtIDLOJA1) <> xVal(Me.TxtIDLOJA2)) Then
      MsgBox "Loja " & StrZero(xVal(Me.TxtIDLOJA2), 2) & " já existe. Por favor, atualize-a para " & StrZero(xVal(Me.TxtIDLOJA2) + 1, 2)
      Me.TxtIDLOJA1.Text = xVal(Me.TxtIDLOJA2)
      Me.TxtIDLOJA2.Text = xVal(Me.TxtIDLOJA2) + 1
      GoTo Saida
   End If
   
   
   Sql = "Select * "
   Sql = Sql & " From COLIGADA"
   Sql = Sql & " Where IDCOLIGADA= " & SqlNum(Me.TxtIDCOLIGADA2.Text)
   Sql = Sql & " Order By IDCOLIGADA"
   If gXDb.Abretabela(Sql) And (xVal(Me.TxtIDCOLIGADA1) <> xVal(Me.TxtIDCOLIGADA2)) And xVal(Me.TxtIDCOLIGADA2) <> 1 Then
      MsgBox "Coligada " & StrZero(xVal(Me.TxtIDCOLIGADA2), 2) & " já existe. Por favor, atualize-a para " & StrZero(xVal(Me.TxtIDCOLIGADA2) + 1, 2)
      Me.TxtIDCOLIGADA1.Text = xVal(Me.TxtIDCOLIGADA2)
      Me.TxtIDCOLIGADA2.Text = xVal(Me.TxtIDCOLIGADA2) + 1
      GoTo Saida
   End If
   
   
   Sql = "sp_msforeachtable 'ALTER TABLE ? NOCHECK CONSTRAINT ALL'"
   'Call gXDb.Executa(Sql)
   Call gXDb.ADOConect.Execute(Sql)
   
   Sql = "IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_CVENDA_RFUNC]') AND parent_object_id = OBJECT_ID(N'[dbo].[CVENDA]')) ALTER TABLE [dbo].[CVENDA] DROP CONSTRAINT [FK_CVENDA_RFUNC];" & vbNewLine
   Sql = Sql & "IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_FLAN_FDESPESA]') AND parent_object_id = OBJECT_ID(N'[dbo].[FLAN]')) ALTER TABLE [dbo].[FLAN] DROP CONSTRAINT [FK_FLAN_FDESPESA];" & vbNewLine
   Sql = Sql & "IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_CPGTOSVENDA_FLAN]') AND parent_object_id = OBJECT_ID(N'[dbo].[FLAN]')) ALTER TABLE [dbo].[FLAN] DROP CONSTRAINT [FK_CPGTOSVENDA_FLAN];" & vbNewLine
   Sql = Sql & "IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[R_226]') AND parent_object_id = OBJECT_ID(N'[dbo].[OMAQDISPAROS]')) ALTER TABLE [dbo].[OMAQDISPAROS] DROP CONSTRAINT [R_226];" & vbNewLine
   Sql = Sql & "IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[R_230]') AND parent_object_id = OBJECT_ID(N'[dbo].[OMAQDISPAROS]')) ALTER TABLE [dbo].[OMAQDISPAROS] DROP CONSTRAINT [R_230];" & vbNewLine
   Sql = Sql & "IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[R_145]') AND parent_object_id = OBJECT_ID(N'[dbo].[FRECIBO]')) ALTER TABLE [dbo].[FRECIBO] DROP CONSTRAINT [R_145];" & vbNewLine
   Sql = Sql & "IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[R_160]') AND parent_object_id = OBJECT_ID(N'[dbo].[RFUNCIONARIO]')) ALTER TABLE [dbo].[RFUNCIONARIO] DROP CONSTRAINT [R_160];" & vbNewLine
   Sql = Sql & "IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[R_169]') AND parent_object_id = OBJECT_ID(N'[dbo].[OEVENTOAGENDA]')) ALTER TABLE [dbo].[OEVENTOAGENDA] DROP CONSTRAINT [R_169];" & vbNewLine
   Sql = Sql & "IF  EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[R_25]') AND parent_object_id = OBJECT_ID(N'[dbo].[OEVENTOAGENDA]')) ALTER TABLE [dbo].[OEVENTOAGENDA] DROP CONSTRAINT [R_25];"
   Call gXDb.ADOConect.Execute(Sql)
               
   Sql = "SET IDENTITY_INSERT COLIGADA on;" & vbNewLine
   Sql = Sql & "If Not Exists(Select * From COLIGADA Where IDCOLIGADA=" & Me.TxtIDCOLIGADA2.Text & ")"
   Sql = Sql & " Insert into COLIGADA (IDCOLIGADA, NMCOLIGADA, TAG)"
   Sql = Sql & " Select " & Me.TxtIDCOLIGADA2.Text & ", NMCOLIGADA, TAG From COLIGADA Where IDCOLIGADA=" & Me.TxtIDCOLIGADA1.Text & ";" & vbNewLine
   Sql = Sql & "SET IDENTITY_INSERT COLIGADA off;" & vbNewLine
   Call gXDb.ADOConect.Execute(Sql)
      
      
   Sql = "Update DELETEDROWS"
   Sql = Sql & " SET DELETEDROWS.QUERY = (Select SUBSTRING(D.QUERY,1, CHARINDEX('IDLOJA = ', D.Query)+8)"
   Sql = Sql & "             + '" & Trim(Me.TxtIDLOJA2.Text) & "'"
   Sql = Sql & "             + SUBSTRING(D.QUERY,CHARINDEX('IDLOJA = ', D.Query)+10, LEN(D.QUERY)-CHARINDEX('IDLOJA = ', D.Query)-8)"
   Sql = Sql & "          From DELETEDROWS D"
   Sql = Sql & "          Where CHARINDEX('IDLOJA = ', D.Query)>0"
   Sql = Sql & "          AND D.IDDELETED=DELETEDROWS.IDDELETED"
   Sql = Sql & "          )"
   Sql = Sql & " WHERE CHARINDEX('IDLOJA = ', DELETEDROWS.Query)>0"
   Sql = Sql & " AND SUBSTRING(DELETEDROWS.QUERY,CHARINDEX('IDLOJA = ', DELETEDROWS.Query)+8, 2)=" & Trim(Me.TxtIDLOJA1.Text)
   Call gXDb.ADOConect.Execute(Sql)
         
   Dim nCol As Integer
   Dim nIDL As Integer
   nCol = Me.TxtIDCOLIGADA2.Text
   nIDL = Me.TxtIDLOJA2.Text
   If AtualizarCampo("IDCOLIGADA", Me.TxtIDCOLIGADA1.Text, Me.TxtIDCOLIGADA2.Text) Then
      Call AtualizarCampo("IDLOJA", Me.TxtIDLOJA1.Text, Me.TxtIDLOJA2.Text)
   End If

      
   Sql = "ALTER TABLE [dbo].[CVENDA]  WITH NOCHECK ADD CONSTRAINT [FK_CVENDA_RFUNC] FOREIGN KEY([IDLOJA], [IDFUNCIONARIO]) REFERENCES [dbo].[RFUNCIONARIO] ([IDLOJA], [IDFUNCIONARIO]);" & vbNewLine
   Sql = Sql & "ALTER TABLE [dbo].[FLAN]  WITH NOCHECK ADD  CONSTRAINT [FK_FLAN_FDESPESA] FOREIGN KEY([IDLOJA], [IDDESP]) REFERENCES [dbo].[FDESPESA] ([IDLOJA],[IDDESP]) ON UPDATE CASCADE;" & vbNewLine
   Sql = Sql & "ALTER TABLE [dbo].[FLAN]  WITH NOCHECK ADD  CONSTRAINT [FK_CPGTOSVENDA_FLAN] FOREIGN KEY([IDLOJA], [IDVENDA], [IDPGTO]) REFERENCES [dbo].[CPGTOSVENDA] ([IDLOJA], [IDVENDA], [IDPGTO]) ON UPDATE CASCADE;" & vbNewLine
   Sql = Sql & "ALTER TABLE OMAQDISPAROS WITH NOCHECK ADD  CONSTRAINT R_226 FOREIGN KEY (IDLOJA,IDATENDIMENTO,IDSESSAO) REFERENCES OSESSAO(IDLOJA,IDATENDIMENTO,IDSESSAO);" & vbNewLine
   Sql = Sql & "ALTER TABLE OMAQDISPAROS WITH NOCHECK ADD  CONSTRAINT R_230 FOREIGN KEY (IDLOJA,IDAREA) REFERENCES OAREA(IDLOJA, IDAREA);" & vbNewLine
   Sql = Sql & "ALTER TABLE FRECIBO WITH NOCHECK ADD CONSTRAINT  R_145 FOREIGN KEY (IDLOJA,IDLOTE) REFERENCES FLOTERPS(IDLOJA, IDLOTE);" & vbNewLine
   Sql = Sql & "ALTER TABLE [dbo].[RFUNCIONARIO]  WITH CHECK ADD CONSTRAINT [R_160] FOREIGN KEY([IDLOJA]) REFERENCES [dbo].[OLOJA] ([IDLOJA]);" & vbNewLine
   Sql = Sql & "ALTER TABLE [dbo].[OEVENTOAGENDA]  WITH NOCHECK ADD  CONSTRAINT [R_169] FOREIGN KEY([IDLOJA], [IDCLIENTE]) REFERENCES [dbo].[OCLIENTE] ([IDLOJA], [IDCLIENTE]);" & vbNewLine
   Sql = Sql & "ALTER TABLE [dbo].[OEVENTOAGENDA]  WITH NOCHECK ADD  CONSTRAINT [R_25] FOREIGN KEY([IDLOJA], [IDEVENTOREC]) REFERENCES [dbo].[OEVENTOREC] ([IDLOJA], [IDEVENTOREC]);"
   Call gXDb.ADOConect.Execute(Sql)
   
   Sql = ""
   Sql = Sql & "ALTER TABLE [dbo].[CVENDA]         NOCHECK  CONSTRAINT [FK_CVENDA_RFUNC];" & vbNewLine
   Sql = Sql & "ALTER TABLE [dbo].[FRECIBO]        NOCHECK  CONSTRAINT [R_145];" & vbNewLine
   Sql = Sql & "ALTER TABLE [dbo].[OEVENTOAGENDA]  NOCHECK  CONSTRAINT [R_25];" & vbNewLine
   Sql = Sql & "ALTER TABLE [dbo].[FLAN]           NOCHECK  CONSTRAINT [FK_CPGTOSVENDA_FLAN];" & vbNewLine
   Sql = Sql & "ALTER TABLE [dbo].[OMAQDISPAROS]   NOCHECK  CONSTRAINT [R_226];" & vbNewLine
   Sql = Sql & "ALTER TABLE [dbo].[OMAQDISPAROS]   NOCHECK  CONSTRAINT [R_230];" & vbNewLine
   Sql = Sql & "ALTER TABLE [dbo].[OEVENTOAGENDA]  NOCHECK  CONSTRAINT [R_169];" & vbNewLine
   Sql = Sql & "ALTER TABLE [dbo].[RFUNCIONARIO]   CHECK    CONSTRAINT [R_160];" & vbNewLine
   Call gXDb.ADOConect.Execute(Sql)
   
   If (xVal(Me.TxtIDCOLIGADA1) <> xVal(Me.TxtIDCOLIGADA2)) Then
      Sql = "If Exists(Select * From COLIGADA Where IDCOLIGADA=" & Me.TxtIDCOLIGADA1.Text & ") Delete From COLIGADA Where IDCOLIGADA=" & Me.TxtIDCOLIGADA1.Text & ";"
      Call gXDb.ADOConect.Execute(Sql)
   End If
'   Sql = "DELETE FROM PARAM WHERE IDLOJA=0; " & vbNewLine
'   Sql = Sql & "UPDATE [PARAM] SET VLPARAM=(SELECT VLPARAM FROM [PARAM] P WHERE P.CODSIS=PARAM.CODSIS AND P.CODPARAM=PARAM.CODPARAM AND P.IDLOJA=1) Where PARAM.IDLOJA <> 1"
'   Call gXDb.ADOConect.Execute(Sql)
   
   Me.TxtIDCOLIGADA1.Text = Me.TxtIDCOLIGADA2.Text
   Me.TxtIDLOJA1.Text = Me.TxtIDLOJA2.Text
   Me.TxtIDLOJA2.SetFocus

   If Not bIDLOJA Then
      Call CmdAtualizar_Click
      Me.TxtIDCOLIGADA2.Text = nCol
      Me.TxtIDLOJA2.Text = nIDL
      If Me.TxtIDCOLIGADA1.Text <> nCol Or Me.TxtIDLOJA1.Text <> nIDL Then
         If gXDb.conectado Then bIDLOJA = True
         Call CmdExeIDLOJA_Click
         If Me.TxtIDCOLIGADA1.Text = nCol Or Me.TxtIDLOJA1.Text = nIDL Then
            MsgBox "OPERAÇÃO REALIZADA COM SUCESSO!!"
         End If
      Else
         MsgBox "OPERAÇÃO REALIZADA COM SUCESSO!!"
      End If
   End If
   If gXDb.conectado Then bIDLOJA = True
   GoTo Saida
TrataErro:
   MsgBox "Erro Nº: " & Err & vbNewLine & Error
   Resume Next
Saida:
   Screen.MousePointer = vbDefault
End Sub
Private Function AtualizarCampo(pNMCAMPO As String, pValor1 As String, pValor2 As String) As Boolean
   Dim Sql As String
   Dim i As Integer
   Dim j As Integer
   Dim RsTab As Object
   Dim kNOk  As Integer
   Dim MyTab As Object
   Dim sTabela As String
   Dim nPos As Long
   Dim sWhere As String

   
   On Error GoTo TrataErro
   
   If xVal(pValor1) = xVal(pValor2) Then
      AtualizarCampo = True
      Exit Function
   End If
   
   Sql = "SELECT O.NAME [TABELA], O.*"
   Sql = Sql & " FROM SYSCOLUMNS C JOIN SYSOBJECTS O ON C.ID=O.ID"
   Sql = Sql & " WHERE C.NAME = '" & pNMCAMPO & "'"
   Sql = Sql & " ORDER BY O.CRDATE"
   
   Me.TxtLogIDLOJA.Text = ""
   Me.TxtLogIDLOJANot.Text = ""
      
   i = 0
   j = 0
   
   If gXDb.Abretabela(Sql, RsTab) Then
      kNOk = xVal(RsTab.Recordcount)
      While bIDLOJA
         j = j + 1
         If j > 1 Then Me.TxtLogIDLOJA.Text = Me.TxtLogIDLOJA.Text & vbNewLine
         Me.TxtLogIDLOJA.Text = Me.TxtLogIDLOJA.Text & StrZero(kNOk, 2) & " ======= CICLO " & StrZero(j, 2) & "==========" & vbNewLine
         Me.TxtLogIDLOJA.Text = Me.TxtLogIDLOJA.Text & "============================" & vbNewLine
         Me.TxtLogIDLOJA.SelStart = Len(Me.TxtLogIDLOJA.Text)
         Me.TxtLogIDLOJA.SelLength = 1
         
         Me.TxtLogIDLOJANot.Text = Me.TxtLogIDLOJANot.Text & StrZero(kNOk, 2) & " ======= CICLO " & StrZero(j, 2) & "==========" & vbNewLine
         Me.TxtLogIDLOJANot.Text = Me.TxtLogIDLOJANot.Text & "============================" & vbNewLine
         Me.TxtLogIDLOJANot.SelStart = Len(Me.TxtLogIDLOJANot.Text)
         Me.TxtLogIDLOJANot.SelLength = 1
         
'         Wait 1
         
         While bPause
            DoEvents
         Wend
         
         bIDLOJA = False
         i = 0
         RsTab.MoveFirst
         kNOk = 0
         While Not RsTab.EOF
            sTabela = RsTab("TABELA")
            If InStr(sTabela, "TRAT") <> 0 Or InStr(sTabela, "PARAM") <> 0 Then
               sTabela = sTabela
            End If
            nPos = RsTab.AbsolutePosition
            Set MyTab = CriarObjeto("BANCO_3R.TB_" & sTabela)
            If MyTab Is Nothing Then
               Set MyTab = CriarObjeto("BANCO.TB_" & sTabela)
            End If
            Me.TxtLogIDLOJA.Text = Me.TxtLogIDLOJA.Text & nPos & "." & sTabela
            If MyTab Is Nothing Then
               kNOk = kNOk + 1
               Me.TxtLogIDLOJA.Text = Me.TxtLogIDLOJA.Text & " Is Nothing" & vbNewLine
            Else
               sWhere = pNMCAMPO & "=" & pValor1 & " Or " & pNMCAMPO & " Is Null Or IDLOJA<=1"
               Set MyTab.XDb = gXDb
               If MyTab.Pesquisar(Ch_Where:=sWhere) Then
                  Err = 0
                  bIDLOJA = True
                  Call AtualizaCampoExtra(sTabela)
                  Sql = "Update " & sTabela
                  Sql = Sql & " Set " & pNMCAMPO & "=" & pValor2
                  Sql = Sql & " Where " & sWhere & ";"
                  If sTabela <> "COLIGADA" Then Call gXDb.ADOConect.Execute(Sql)
                  'If gXDb.Executa(Sql) Then
                  If Err = 0 Then
                     Me.TxtLogIDLOJA.Text = Me.TxtLogIDLOJA.Text & ": " & xVal(MyTab.Rs.Recordcount) & " Is Ok" & vbNewLine
                     bIDLOJA = bIDLOJA Or False
                  Else
                     Me.TxtLogIDLOJA.Text = Me.TxtLogIDLOJA.Text & ": " & xVal(MyTab.Rs.Recordcount) & " Not Ok" & vbNewLine
                     Me.TxtLogIDLOJANot.Text = Me.TxtLogIDLOJANot.Text & nPos & "." & sTabela
                     Me.TxtLogIDLOJANot.Text = Me.TxtLogIDLOJANot.Text & ": " & xVal(MyTab.Rs.Recordcount) & " Not Ok" & vbNewLine
                     bIDLOJA = True
                     kNOk = kNOk + 1
                  End If
               Else
                  sWhere = pNMCAMPO & "=" & pValor2
                  If MyTab.Pesquisar(Ch_Where:=sWhere) Then
                     Me.TxtLogIDLOJA.Text = Me.TxtLogIDLOJA.Text & ": Already Ok (" & MyTab.Rs.Recordcount & ")" & vbNewLine
                  Else
                     Sql = "Select Count(*) [QTD]"
                     Sql = Sql & " From " & sTabela
                     Sql = Sql & " Where " & sWhere
                     If gXDb.Abretabela(Sql) Then
                        If xVal(gXDb.RsAux("QTD")) = 0 Then
                           Me.TxtLogIDLOJA.Text = Me.TxtLogIDLOJA.Text & ": Already Ok (0)" & vbNewLine
                        Else
                           Me.TxtLogIDLOJA.Text = Me.TxtLogIDLOJA.Text & ": Not Ok (" & gXDb.RsAux("QTD") & ")" & vbNewLine
                           Me.TxtLogIDLOJANot.Text = Me.TxtLogIDLOJANot.Text & nPos & "." & sTabela
                           Me.TxtLogIDLOJANot.Text = Me.TxtLogIDLOJANot.Text & ": " & xVal(0) & " Not Ok" & vbNewLine
                           bIDLOJA = True
                           kNOk = kNOk + 1
                        End If
                     Else
                        Me.TxtLogIDLOJA.Text = Me.TxtLogIDLOJA.Text & ": ***Erro Count" & vbNewLine
                        Me.TxtLogIDLOJANot.Text = Me.TxtLogIDLOJANot.Text & nPos & "." & sTabela
                        Me.TxtLogIDLOJANot.Text = Me.TxtLogIDLOJANot.Text & ": " & xVal(MyTab.Rs.Recordcount) & " Not Ok" & vbNewLine
                        bIDLOJA = True
                        kNOk = kNOk + 1
                     End If
                  End If
               End If
            End If
            Me.TxtLogIDLOJA.SelStart = Len(Me.TxtLogIDLOJA.Text)
            Me.TxtLogIDLOJA.SelLength = 1
            
            Set MyTab = Nothing
            RsTab.MoveNext
            DoEvents
         Wend
         If kNOk = 0 Then
         'If i > RsTab.Recordcount Then
            bIDLOJA = False
         End If
      Wend
      sTabela = ""
      Err = 0
      Sql = "sp_msforeachtable 'ALTER TABLE ? WITH CHECK CHECK CONSTRAINT ALL'"
      Call gXDb.ADOConect.Execute(Sql)
      
      'Call gXDb.Executa(Sql)
      Me.LblObjetos = "Objetos: " & xVal(gXDb.RsAux.Recordcount)
   End If
   AtualizarCampo = (Err = 0)
   Exit Function
TrataErro:
   If sTabela = "" Then
      MsgBox "Erro Nº: " & Err & vbNewLine & Error
   Else
      MsgBox "Erro Nº: " & Err & " / " & sTabela & vbNewLine & Error
   End If
   Resume Next
End Function
Private Sub AtualizaCampoExtra(sTabela As String)
   Dim Sql As String
   If sTabela = "OEVENTOAGENDA" Then
      Sql = "Update OEVENTOAGENDA "
      Sql = Sql & " Set ScheduleID = (IDLOJA*1000) + IDSALA"
      Call gXDb.ADOConect.Execute(Sql)
   ElseIf sTabela = "PARAM" Then
      Sql = "Delete From PARAM"
      Sql = Sql & " Where CODPARAM in (select codparam from param group by codparam having count(codparam)>1)"
      Sql = Sql & " And IDLOJA<=1"
      Call gXDb.ADOConect.Execute(Sql)
   ElseIf sTabela = "OSALA_MAQUINA" Then
      Sql = "Delete From OSALA_MAQUINA"
      Sql = Sql & " Where IDMAQUINA Not In (Select m.idmaquina from omaquina m)"
      Call gXDb.ADOConect.Execute(Sql)
   End If
End Sub
Private Sub cMDpAUSE_Click()
   If bPause Then
      Me.CmdPause.Caption = "Parar"
   Else
      Me.CmdPause.Caption = "Continuar"
   End If
   bPause = Not bPause
End Sub

Private Sub CmdSair_Click()
  End
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

Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
   Me.TxtArquivo.SetFocus
End Sub

Private Sub Form_Initialize()
   If Not LCase(InputBoxPassword("Entre com a Senha")) = "dolphin" Then
      End
   End If
End Sub

Private Sub Form_Load()
   Me.Frame1.Top = 120
   Me.Frame1.Left = Me.Frame1.Top
   
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
   'Me.TxtLICDest.Text = ""
   'Me.TxtLICOrig.Text = ""
   Me.TxtPWD.Text = ""
   Me.TxtServer.Text = ""
   'Me.TxtTAG.Text = ""
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

Private Sub TxtIDLOJA1_LostFocus()
   Me.TxtIDLOJA2.Text = xVal(Me.TxtIDLOJA1.Text) + 1
End Sub
