VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.Controls.v11.2.2.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#11.2#0"; "Codejock.ReportControl.v11.2.2.ocx"
Begin VB.Form FrmRestoreP3R 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "  Restauração 3R"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   7695
   Icon            =   "FrmRestoreP3R.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl TabControl 
      Height          =   4695
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   7455
      _Version        =   720898
      _ExtentX        =   13150
      _ExtentY        =   8281
      _StockProps     =   68
      Color           =   8
      PaintManager.Layout=   2
      PaintManager.BoldSelected=   -1  'True
      ItemCount       =   2
      SelectedItem    =   1
      Item(0).Caption =   "Log"
      Item(0).ControlCount=   3
      Item(0).Control(0)=   "TxtLog"
      Item(0).Control(1)=   "ProgBar"
      Item(0).Control(2)=   "LblProg"
      Item(1).Caption =   "Restore"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "GrdSinc"
      Begin XtremeReportControl.ReportControl GrdSinc 
         Height          =   3975
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   7215
         _Version        =   720898
         _ExtentX        =   12726
         _ExtentY        =   7011
         _StockProps     =   64
      End
      Begin XtremeSuiteControls.FlatEdit TxtLog 
         Height          =   3975
         Left            =   -69880
         TabIndex        =   8
         Top             =   360
         Visible         =   0   'False
         Width           =   7215
         _Version        =   720898
         _ExtentX        =   12726
         _ExtentY        =   7011
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "Log..."
         MultiLine       =   -1  'True
         ScrollBars      =   2
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ProgressBar ProgBar 
         Height          =   195
         Left            =   -69880
         TabIndex        =   10
         Top             =   4400
         Visible         =   0   'False
         Width           =   6735
         _Version        =   720898
         _ExtentX        =   11880
         _ExtentY        =   353
         _StockProps     =   93
         Scrolling       =   1
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.Label LblProg 
         Height          =   255
         Left            =   -63040
         TabIndex        =   12
         Top             =   4360
         Visible         =   0   'False
         Width           =   375
         _Version        =   720898
         _ExtentX        =   661
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "100%"
         Alignment       =   1
         Transparent     =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox GrpDetalhes 
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   7455
      _Version        =   720898
      _ExtentX        =   13150
      _ExtentY        =   1296
      _StockProps     =   79
      Caption         =   "Detalhes"
      Begin XtremeSuiteControls.RadioButton OptMerge 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   2055
         _Version        =   720898
         _ExtentX        =   3625
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Exporta e Importa dados"
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton OptMerge 
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   2055
         _Version        =   720898
         _ExtentX        =   3625
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Apenas Exporta"
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton OptMerge 
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   2055
         _Version        =   720898
         _ExtentX        =   3625
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Apenas Importa"
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.PushButton CmdRestore 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   5880
      Width           =   1335
      _Version        =   720898
      _ExtentX        =   2355
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Restaurar"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5280
      Top             =   5880
   End
   Begin XtremeSuiteControls.PushButton CmdSair 
      Height          =   495
      Left            =   6120
      TabIndex        =   1
      Top             =   5880
      Width           =   1335
      _Version        =   720898
      _ExtentX        =   2355
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Sair"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton CmdParar 
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   5880
      Width           =   1455
      _Version        =   720898
      _ExtentX        =   2566
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Parar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.Label LblConect 
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   5520
      Width           =   7335
      _Version        =   720898
      _ExtentX        =   12938
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Label1"
   End
   Begin XtremeSuiteControls.TrayIcon TrayIcon 
      Left            =   2280
      Top             =   5880
      _Version        =   720898
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   16
      Text            =   "Restauração de Base 3R"
      Picture         =   "FrmRestoreP3R.frx":0442
   End
   Begin VB.Menu MnuTray 
      Caption         =   "MnuTray"
      Visible         =   0   'False
      Begin VB.Menu MnuSinc 
         Caption         =   "&Abrir Restauração 3R"
         Index           =   0
      End
      Begin VB.Menu MnuSinc 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu MnuSinc 
         Caption         =   "&Parar"
         Index           =   2
      End
      Begin VB.Menu MnuSinc 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu MnuSinc 
         Caption         =   "Sai&r"
         Index           =   4
      End
   End
End
Attribute VB_Name = "FrmRestoreP3R"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bTeste     As Boolean
Dim MyRestore  As Object
Dim xDbLocal   As Object
Dim xDbRemoto  As Object
Private Sub CmdParar_Click()
   If Not MyRestore Is Nothing Then
      MyRestore.Pause = Not MyRestore.Pause
      If MyRestore.Pause Then
         Me.CmdParar.Caption = "Continuar"
         Me.CmdSair.Enabled = True
      Else
         Me.CmdParar.Caption = "Parar"
         Me.CmdSair.Enabled = False
      End If
   Else
      Me.CmdSair.Enabled = Not Me.CmdSair.Enabled
      Me.CmdParar.Caption = IIf(Me.CmdSair.Enabled, "Continuar", "Parar")
      Me.TrayIcon.Text = Trim(Me.Tag) & IIf(Me.CmdSair.Enabled, " [Processo Parado]", "")
      While Me.CmdSair.Enabled: DoEvents: Wend
   End If
End Sub
Private Sub CmdSair_Click()
   End
End Sub
Private Sub CmdRestore_Click()
   Dim sLocalTag  As String
   Dim sRemoteTag As String
   Dim Sql        As String
   Dim sLojasIn   As String
   Dim sLojasLike As String
   Dim MyRs       As Object
      
   On Error GoTo TrataErro
   'Screen.MousePointer = vbHourglass
   
   If Me.CmdRestore.Caption = "Resincronizar" Then
      Call Resinc
      Exit Sub
   End If
   
   Me.Timer1.Enabled = False

   Me.CmdRestore.Enabled = False
   Me.CmdParar.Enabled = True
   Me.CmdSair.Enabled = False
   
   sLocalTag = LocalTagServ()
   sRemoteTag = RemoteTagServ()
      
   Me.LblConect.Caption = "[" & GetTag(sLocalTag, "SERVER", "") & "].[" & GetTag(sLocalTag, "DBNAME", "") & "]"
   Me.LblConect.Caption = Me.LblConect.Caption & " <-> [Remote].[" & GetTag(sRemoteTag, "DBNAME", "") & "]"
   Me.LblConect.Caption = UCase(Me.LblConect.Caption)
   
   While Not IsWebConnected
      DoEvents
   Wend
   If gDebug Then MsgBox sLocalTag
   Call ConectarDbRemoto(sRemoteTag) 'mssql.classeaconsultoria.com.br
   If ConectarDbLocal(sLocalTag) Then
      Sql = "Select IDLOJA " & vbNewLine
      Sql = Sql & " From OLOJA" & vbNewLine
      Sql = Sql & " Where IDCOLIGADA = (Select Min(IDCOLIGADA)" & vbNewLine
      Sql = Sql & "                      From COLIGADA" & vbNewLine
      Sql = Sql & "                      Where IDCOLIGADA<>1)" & vbNewLine
      If xDbLocal.AbreTabela(Sql, MyRs) Then
         While Not MyRs.EOF
            sLojasIn = sLojasIn & IIf(Trim(sLojasIn) = "", "", ",") & MyRs("IDLOJA") & ""
            sLojasLike = sLojasLike & IIf(Trim(sLojasLike) = "", "", " Or ") & "QUERY Like '%IDLOJA = " & MyRs("IDLOJA") & "%'"
            
            MyRs.MoveNext
         Wend
      Else
         Me.CmdRestore.Enabled = True
         Me.CmdParar.Enabled = False
         Me.CmdSair.Enabled = True
         Me.LblConect.Caption = Me.LblConect.Caption & " - Sem conexão com Banco de Dados local."
         Exit Sub
      End If
   Else
     Me.CmdRestore.Enabled = True
     Me.CmdParar.Enabled = False
     Me.CmdSair.Enabled = True
     Me.LblConect.Caption = Me.LblConect.Caption & " - Sem conexão com Banco de Dados local."
     Exit Sub
   End If
   
   
   Set MyRestore = Nothing
   Set MyRestore = CriarObjeto("RestoreBd.NG_Restore")
   
   With MyRestore
      Set .FrmObj = Me
      .IniFile = Environ("PROGRAMFILES") & "\ClasseA\Projeto3R\RESTOREBD.ini"
      
      .ArrayNoSync = Array("DELETEDROWS", "USUARIO", "COLIGADA", "OLOJA", "GSINC")
      .FieldsOnTab = "'IDLOJA','TIMESTAMP'"
      .DelScriptTab = "DELETEDROWS"
   
      .LojasIn = sLojasIn
      .LojasLike = sLojasLike
      
      .SincFilter = ""
      '.SincFilter = .SincFilter & "(ALTERSTAMP=1 Or ALTERSTAMP Is Null) And"
      .SincFilter = .SincFilter & " IDLOJA In (" & sLojasIn & ")"
      .SincFilter = .SincFilter & " And (TIMESTAMP>=("
      .SincFilter = .SincFilter & " Select IsNull(Min(DTSINC),0)-(0.00208)"
      .SincFilter = .SincFilter & " From GSINC"
      .SincFilter = .SincFilter & " Where IDLOJA In (" & sLojasIn & ")"
      .SincFilter = .SincFilter & " And CODMAQ=" & SqlStr(Environ("COMPUTERNAME"))
      .SincFilter = .SincFilter & " And TABELA='@@TABELA'"
      .SincFilter = .SincFilter & "))"
      
      
      .DeletedFilter = ""
      '.DeletedFilter = .DeletedFilter & "(ALTERSTAMP=1 Or ALTERSTAMP Is Null) And "
      .DeletedFilter = .DeletedFilter & " (" & sLojasLike & ")"
      .DeletedFilter = .DeletedFilter & " And (TIMESTAMP>=("
      .DeletedFilter = .DeletedFilter & " Select IsNull(Min(DTSINC),0)"
      .DeletedFilter = .DeletedFilter & " From GSINC"
      .DeletedFilter = .DeletedFilter & " Where IDLOJA In (" & sLojasIn & ")"
      .DeletedFilter = .DeletedFilter & " And CODMAQ=" & SqlStr(Environ("COMPUTERNAME"))
      .DeletedFilter = .DeletedFilter & " And TABELA=" & SqlStr("DELETEDROWS")
      .DeletedFilter = .DeletedFilter & "))"
      
      .LocalServer = GetTag(sLocalTag, "SERVER", "")
      .LocaldbName = GetTag(sLocalTag, "DBNAME", "")
      .LocalUID = GetTag(sLocalTag, "UID", "")
      .LocalPWD = GetTag(sLocalTag, "PWD", "")
      
      .RemoteServer = GetTag(sRemoteTag, "SERVER", "")
      .RemotedbName = GetTag(sRemoteTag, "DBNAME", "")
      .RemoteUID = GetTag(sRemoteTag, "UID", "")
      .RemotePWD = GetTag(sRemoteTag, "PWD", "")
            
      
      .Import = True 'InArray(GetTag(sRemoteTag, "MERGE", "False"), Array("0", "2")) ' (Me.ChkMerge.Value = xtpChecked)
      .Export = True
      
      
      .WebConnected = IsWebConnected
      .Run
      
      If Not .WebConnected Then
         Me.Timer1.Enabled = True
      End If
   End With
   Me.CmdRestore.Enebled = True
   Me.CmdParar.Enebled = False
   Me.CmdSair.Enebled = True
   
   Exit Sub
TrataErro:
   Me.Timer1.Enabled = True
End Sub
Private Function ConectarDbLocal(pTag As String) As Boolean
   Set xDbLocal = Nothing
   Set xDbLocal = CriarObjeto("XBANCO01.DS_BANCO")
   With xDbLocal
      .Server = GetTag(pTag, "SERVER", "")
      .dbName = GetTag(pTag, "DBNAME", "")
      .UID = GetTag(pTag, "UID", "")
      .PWD = GetTag(pTag, "PWD", "")
      .SrvConecta
      If .Conectado Then
         ConectarDbLocal = True
      Else
         ConectarDbLocal = False
      End If
   End With
End Function
Private Function ConectarDbRemoto(pTag As String) As Boolean
   Set xDbRemoto = Nothing
   Set xDbRemoto = CriarObjeto("XBANCO01.DS_BANCO")
   With xDbRemoto
      .Server = GetTag(pTag, "SERVER", "")
      .dbName = GetTag(pTag, "DBNAME", "")
      .UID = GetTag(pTag, "UID", "")
      .PWD = GetTag(pTag, "PWD", "")
      .SrvConecta
      If .Conectado Then
         ConectarDbRemoto = True
      Else
         ConectarDbRemoto = False
      End If
   End With
End Function

Private Function RemoteTagServ() As String
   Dim sTag       As String
   Dim iLast      As Integer
   Dim sIniFile   As String
   
   sIniFile = gLocalPath & "RESTOREBD.ini"
   
   sTag = ""
   sTag = SetTag(sTag, "INIFILE", sIniFile)
   
   sTag = ""
   sTag = sTag & "|SERVER=[Remote]"
   sTag = sTag & "|DBNAME=G3RTESTE"
   sTag = sTag & "|UID=USU_VERIF"
   sTag = sTag & "|PWD=MINOTAURO"
   sTag = sTag & "|MERGE=0"
   sTag = sTag & "|"
        
   If ExisteArquivo(sIniFile) Then
      sTag = ReadIniFile(sIniFile, "CONFIG", "TAG", Encrypt2(sTag))
   Else
      sTag = Encrypt2(sTag)
      Call WriteIniFile(sIniFile, "CONFIG", "TAG", sTag)
      Call WriteIniFile(sIniFile, "CONFIG", "MERGE", sTag)
      sTag = Decrypt2(ReadIniFile(sIniFile, "CONFIG", "TAG", sTag))
   End If
   
   If Trim(sTag) = "" Then
      Exit Function
   Else
      sTag = Decrypt2(ReadIniFile(sIniFile, "CONFIG", "TAG", ""))
   End If
   If ReadIniFile(sIniFile, "CONFIG", "SERVER", "") <> "" Then sTag = SetTag(sTag, "SERVER", ReadIniFile(sIniFile, "CONFIG", "SERVER", ""))
   If ReadIniFile(sIniFile, "CONFIG", "DBNAME", "") <> "" Then sTag = SetTag(sTag, "DBNAME", ReadIniFile(sIniFile, "CONFIG", "DBNAME", ""))
   If ReadIniFile(sIniFile, "CONFIG", "UID", "") <> "" Then sTag = SetTag(sTag, "UID", ReadIniFile(sIniFile, "CONFIG", "UID", ""))
   If ReadIniFile(sIniFile, "CONFIG", "PWD", "") <> "" Then sTag = SetTag(sTag, "PWD", ReadIniFile(sIniFile, "CONFIG", "PWD", ""))
   If ReadIniFile(sIniFile, "CONFIG", "MERGE", "") <> "" Then sTag = SetTag(sTag, "MERGE", ReadIniFile(sIniFile, "CONFIG", "MERGE", ""))
   
   If bTeste Then
      Dim sDbName As String
      sDbName = UCase(InputBox("Banco Remoto", "Sinc3R", "G3RTESTE"))
      sTag = SetTag(sTag, "DBNAME", sDbName)
   End If
   
   RemoteTagServ = sTag
End Function
Private Function LocalTagServ() As String
   Dim sTag As String
   Dim iLast As Integer
   Dim sDbName As String
   
   If sTag = "" Then
      gLocalReg = gLocalPath & "P3R.reg"
      gSetupFile = "SETUP.INI"
      
      iLast = ReadIniFile(gLocalReg, "Conections", "Last", "0")
      Call WriteIniFile(gLocalReg, "Conections", "Last", "0")
      bTeste = False
      If InStr(App.Path, "\Sistemas\Dsr\Projeto3R\") <> 0 Then
         bTeste = (vbYes = ExibirPergunta("Continuar em teste?", "Sincronização"))
         If bTeste Then
            sDbName = InputBox("Banco de Teste.", "Sinc3R", "G3R_Freguesia")
            Call WriteIniFile(gLocalReg, "Conections", "Last", "1")
            'Call WriteIniFile(gLocalReg, "Conections", "Last", "2")
         End If
      End If
      Call MyLoadgCODSIS
      Call WriteIniFile(gLocalReg, "Conections", "Last", CStr(iLast))
      If bTeste Then gDBNAME = UCase(sDbName)
      
      sTag = ""
      sTag = SetTag(sTag, "EXEPATH", gLocalPath)
      'sTag= SetTag(sTag, "EXEPATH", App.Path & "\")
      sTag = SetTag(sTag, "CODSIS", gCODSIS)
      sTag = SetTag(sTag, "SERVER", gSERVER)
      sTag = SetTag(sTag, "DBNAME", gDBNAME)
      sTag = SetTag(sTag, "UID", gDBUSER)
      sTag = SetTag(sTag, "PWD", gDBPWD)
   End If
   
   LocalTagServ = sTag
End Function
Private Sub Form_Activate()
   Me.MousePointer = vbDefault
   Screen.MousePointer = vbDefault
   'Me.Hide
   'Me.Show

End Sub

Private Sub Form_Load()
   Screen.MousePointer = vbHourglass
   Me.Tag = Me.Caption
   Me.Move 0, 0
   Me.OptMerge(1).Enabled = False
   Me.OptMerge(2).Enabled = False
   'Call CmdRestore_Click
   Me.TrayIcon.MinimizeToTray Me.hwnd
   
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If UnloadMode = 0 Then
      TrayIcon.MinimizeToTray Me.hwnd
      Cancel = True
   End If
End Sub
Private Sub Resinc()
   Dim Sql As String
   Dim sDate As String
   Dim sLojasIn As String
   Dim sLojasLike As String
   Dim bPause As Boolean
   Dim MyRs As Object
   
   If Not MyRestore Is Nothing Then
      bPause = MyRestore.Pause
      MyRestore.Pause = True
   End If
   
   sDate = InputBox("Resincronizar a partir de:", "Resincronizar", Format(Now() - 1, "dd/mm/yyyy"))
   If IsDate(sDate) Then
      sDate = sDate & " 00:00:00"
      
      If xDbLocal.Conectado Then
         Sql = "Select IDLOJA " & vbNewLine
         Sql = Sql & " From OLOJA" & vbNewLine
         Sql = Sql & " Where IDCOLIGADA = (Select Min(IDCOLIGADA)" & vbNewLine
         Sql = Sql & "                      From COLIGADA" & vbNewLine
         Sql = Sql & "                      Where IDCOLIGADA<>1)" & vbNewLine
         sLojasIn = ""
         sLojasLike = ""
         If xDbLocal.AbreTabela(Sql, MyRs) Then
            While Not MyRs.EOF
               sLojasIn = sLojasIn & IIf(Trim(sLojasIn) = "", "", ",") & MyRs("IDLOJA") & ""
               sLojasLike = sLojasLike & IIf(Trim(sLojasLike) = "", "", " Or ") & "QUERY Like '%IDLOJA = " & MyRs("IDLOJA") & "%'"
               
               MyRs.MoveNext
            Wend
      
            Sql = "Update GSINC"
            Sql = Sql & " Set TIMESTAMP=" & SqlDate(sDate)
            Sql = Sql & " , DTSINC=" & SqlDate(sDate)
            Sql = Sql & " Where CODMAQ=" & SqlStr(Environ("COMPUTERNAME"))
            Sql = Sql & " And DTSINC>=" & SqlDate(sDate)
            Sql = Sql & " And IDLOJA in (" & sLojasIn & ");"
            If xDbRemoto.Executa(Sql) Then
               If xDbLocal.Executa(Sql) Then
                  Sql = "Update DELETEDROWS"
                  Sql = Sql & " Set ALTERSTAMP=1"
                  Sql = Sql & " Where TIMESTAMP>=" & SqlDate(sDate)
                  Sql = Sql & " And ALTERSTAMP=0"
                  Sql = Sql & " And (" & sLojasLike & ");"
                  If xDbRemoto.Executa(Sql) Then
                     If xDbLocal.Executa(Sql) Then
                        If xDbLocal.AbreTabela("Select * From GSINC", MyRs) Then
                           If Not MyRestore Is Nothing Then
                              Call MyRestore.Resinc(sDate)
                           End If
                           Call FillRCFromRS(MyRs, Me.GrdSinc)
                           Me.GrdSinc.Populate
                        End If
                     End If
                  End If
               End If
            End If
         End If
      End If
   End If
   If Not MyRestore Is Nothing Then MyRestore.Pause = bPause
End Sub

Private Sub MnuSinc_Click(Index As Integer)
   Select Case Index
      Case 0: TrayIcon.MaximizeFromTray Me.hwnd
      Case 2: If Not MyRestore Is Nothing Then MyRestore.Pause = True
      Case 4: End
   End Select
End Sub

Private Sub OptMerge_Click(Index As Integer)
   If Not MyRestore Is Nothing Then
      If Me.OptMerge(0).Value Then
         MyRestore.Import = True
         MyRestore.Export = True
      ElseIf Me.OptMerge(1).Value Then
         MyRestore.Import = False
         MyRestore.Export = True
      ElseIf Me.OptMerge(2).Value Then
         MyRestore.Import = True
         MyRestore.Export = False
      End If
   End If
End Sub

Private Sub ReportControl1_ColumnClick(ByVal Column As XtremeReportControl.IReportColumn)

End Sub

Private Sub TabControl_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
   Dim MyRs As Object
   If Item.Index = 0 Then
      Me.CmdRestore.Caption = "Sincronizar"
      Me.CmdRestore.Enabled = Me.CmdSair.Enabled
   
   ElseIf Item.Index = 1 Then
      Me.CmdRestore.Caption = "Resincronizar"
      Me.CmdRestore.Enabled = True
      If xDbLocal.Conectado Then
         If xDbLocal.AbreTabela("Select * from GSINC", MyRs) Then
            Call FillRCFromRS(MyRs, Me.GrdSinc)
         End If
      End If
   End If
End Sub

Private Sub Timer1_Timer()
   DoEvents
   If IsWebConnected Then
      If Me.CmdRestore.Caption = "Sincronizar" Then
         Call CmdRestore_Click
      End If
   End If
End Sub

Private Sub TrayIcon_DblClick()
   TrayIcon.MaximizeFromTray Me.hwnd
End Sub
Private Sub TrayIcon_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (Button = 2) Then Me.PopupMenu MnuTray
End Sub

