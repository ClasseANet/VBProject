VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.Controls.v11.2.2.ocx"
Begin VB.Form FrmDownload 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Atenção!"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5640
   ForeColor       =   &H80000001&
   Icon            =   "Dowload.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   5640
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicTit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   -240
      ScaleHeight     =   2625
      ScaleWidth      =   9225
      TabIndex        =   1
      Top             =   -120
      Width           =   9255
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   $"Dowload.frx":1582
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   360
         TabIndex        =   16
         Top             =   1080
         Width           =   5415
      End
      Begin VB.Label LblTit2 
         BackStyle       =   0  'Transparent
         Caption         =   "Para instalar os componentes clique em [Continuar...]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   2160
         Width           =   5775
      End
      Begin VB.Label LblTit1 
         BackStyle       =   0  'Transparent
         Caption         =   $"Dowload.frx":1623
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   2
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.CommandButton CmdInicio 
      Caption         =   "CONTINUAR..."
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
      TabIndex        =   0
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   5175
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1080
         Picture         =   "Dowload.frx":16B6
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   14
         Top             =   1680
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1080
         Picture         =   "Dowload.frx":1800
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1080
         Picture         =   "Dowload.frx":194A
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   8
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1080
         Picture         =   "Dowload.frx":1A94
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   7
         Top             =   960
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1080
         Picture         =   "Dowload.frx":1BDE
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   6
         Top             =   1320
         Visible         =   0   'False
         Width           =   255
      End
      Begin XtremeSuiteControls.ProgressBar ProgressBar1 
         Height          =   180
         Left            =   240
         TabIndex        =   5
         Top             =   2280
         Visible         =   0   'False
         Width           =   4815
         _Version        =   720898
         _ExtentX        =   8493
         _ExtentY        =   317
         _StockProps     =   93
         ForeColor       =   -2147483630
         Appearance      =   1
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Criando Banco de Dados..."
         Height          =   255
         Left            =   1440
         TabIndex        =   17
         Top             =   1920
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Projeto 3R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   15
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "FrameWork"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   13
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Windows Installer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   12
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Sql Server"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   11
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Sql Management Studio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   1320
         Width           =   2655
      End
   End
End
Attribute VB_Name = "FrmDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PrgBar As Object
Dim xConn   As Object
Private Sub CmdInicio_Click()
   Me.CmdInicio.Enabled = False
   Call InstalarFerramentas
   Unload Me
   End
   Me.CmdInicio.Enabled = True
End Sub
Private Sub InstalarFerramentas()
   Dim sComando As String
   Dim bExiste  As Boolean
   Dim sArq1     As String
   Dim sArq2     As String
   Dim sPathMSI  As String
   Dim sMsg      As String
   
   'Screen.MousePointer = vbHourglass
   On Error GoTo Saida
   
   Me.ProgressBar1.Visible = True
   If Me.ProgressBar1.Scrolling <> 2 Then
      Me.ProgressBar1.Scrolling = 2
   Else
      Me.ProgressBar1.Scrolling = 0
   End If

   '*****************
   '* DESCOMPACTAR ARQUIVOS
   If gMetodo1 Then
      If Not ExisteArquivo(gLocalPathSetup & gFileMSDE) Then
         If ExisteArquivo(gLocalPath & gFileMSDE) Then
            Call CopiarArquivo(gLocalPath & gFileMSDE, gLocalPathSetup & gFileMSDE)
         End If
      End If
      If ExisteArquivo(gLocalPathSetup & gFileMSDE) Then
         bExiste = (Dir(gLocalPath & "\P3R.msi", vbArchive) <> "")
         bExiste = bExiste And (Dir(gLocalPathSetup & "01 Windows Installer", vbDirectory) <> "")
         bExiste = bExiste And (Dir(gLocalPathSetup & "02 FrameWork", vbDirectory) <> "")
         bExiste = bExiste And (Dir(gLocalPathSetup & "03 Sql2005Express", vbDirectory) <> "")
         'bExiste = bExiste And (Dir(gLocalPathSetup & "04 SqlManager", vbDirectory) <> "")
         bExiste = bExiste And (Dir(gLocalPathSetup & "05 TeamViewer", vbDirectory) <> "")
         bExiste = bExiste And (Dir(gLocalPathSetup & "99 Database", vbDirectory) <> "")
         If Not bExiste Then
            Call DescompactarArquivo(gLocalPathSetup, gFileMSDE, gLocalPathSetup)
         End If
      Else
         MsgBox "Arquivo de instalação '" & gLocalPath & gFileMSDE & "' não existe ou está corrompido." & vbNewLine & "A instalação não poderá continuar.", vbInformation + vbOKOnly, "Atenção!"
         End
      End If
   Else
      If Dir(gLocalPath & "\P3R.msi", vbArchive) = "" Then
         If ExisteArquivo(gLocalPathSetup & "MSI.zia") Then
            Call DescompactarArquivo(gLocalPathSetup, "MSI.zia", gLocalPath)
         End If
      End If
      If Dir(gLocalPath & "01 Windows Installer", vbDirectory) = "" Then
         If ExisteArquivo(gLocalPathSetup & "P01.zia") Then
            Call DescompactarArquivo(gLocalPathSetup, "P01.zia", gLocalPathSetup)
         End If
      End If
      If Dir(gLocalPath & "02 FrameWork", vbDirectory) = "" Then
         If ExisteArquivo(gLocalPathSetup & "P02.zia") Then
            Call DescompactarArquivo(gLocalPathSetup, "P02.zia", gLocalPathSetup)
         End If
      End If
      If Dir(gLocalPath & "03 Sql2005Express", vbDirectory) = "" Then
         If ExisteArquivo(gLocalPathSetup & "P03.zia") Then
            Call DescompactarArquivo(gLocalPathSetup, "P03.zia", gLocalPathSetup)
         End If
      End If
      'If Dir(gLocalPath & "04 SqlManager", vbDirectory) = "" Then
      '   If ExisteArquivo(gLocalPathSetup & "P01.zia") Then
      '      Call DescompactarArquivo(gLocalPathSetup, "P01.zia", gLocalPathSetup)
      '   End If
      'End If
      'If Dir(gLocalPath & "05 TeamViewer", vbDirectory) = "" Then
      '   If ExisteArquivo(gLocalPathSetup & "P01.zia") Then
      '      Call DescompactarArquivo(gLocalPathSetup, "P01.zia", gLocalPathSetup)
      '   End If
      'End If
      If Dir(gLocalPath & "99 Database", vbDirectory) = "" Then
         If ExisteArquivo(gLocalPathSetup & "P99.zia") Then
            Call DescompactarArquivo(gLocalPathSetup, "P99.zia", gLocalPathSetup)
         End If
      End If
   End If
   '*****************
   '* COPIAR MSI
   'sPathMSI = ResolvePathName(Mid(gLocalPath, 1, Len(gLocalPath) - 6))
   sPathMSI = ResolvePathName(gLocalPath)
   If ExisteArquivo(ResolvePathName(gLocalPathSetup) & "P3R.MSI") Then
      Call CopiarArquivo(ResolvePathName(gLocalPathSetup) & "P3R.MSI", sPathMSI & "P3R.MSI")
      Call ExcluirArquivo(ResolvePathName(gLocalPathSetup) & "P3R.MSI")
   End If
   
   '*****************
   '* INSTALAR WINDOWS INSTALLER
   If ExisteArquivo(gLocalPathSetup & "01 Windows Installer\WindowsXP-KB942288-v3-x86.exe") Then
      bExiste = False
      sArq1 = Environ("SystemRoot") & "\system32\msi.dll"
      If ExisteArquivo(sArq1) Then
         bExiste = (Mid(GetFileVersion(sArq1), 1, 7) >= "004.005")
      End If
      If Not bExiste Then
         sComando = gLocalPathSetup & "01 Windows Installer\WindowsXP-KB942288-v3-x86.exe /passive /norestart"
         Call SincShell(sComando, vbNormalFocus, True)
      End If
   End If
   'Me.Picture1.Visible = Me.Label1.Visible
   Call VerficarProgramas(False)
         
   '*****************
   '* INSTALAR FRAMEWORK 3.5
   If ExisteArquivo(gLocalPathSetup & "01 Windows Installer\WindowsXP-KB942288-v3-x86.exe") Then
      bExiste = False
      sArq1 = Environ("SystemRoot") & "\Microsoft.NET\Framework\v2*"
      bExiste = (Dir(sArq1, vbDirectory) <> "")
      sArq1 = Environ("SystemRoot") & "\Microsoft.NET\Framework\v3.5*"
      sArq2 = Environ("SystemRoot") & "\Microsoft.NET\Framework\v4*"
      bExiste = bExiste Or (Dir(sArq1, vbDirectory) <> "" Or Dir(sArq2, vbDirectory) <> "")
      If Not bExiste Then
         sComando = gLocalPathSetup & "02 FrameWork\dotnetfx35setup.exe /passive /norestart /fo"
         Call SincShell(sComando, vbNormalFocus, True)
      End If
   End If
   'Me.Picture2.Visible = Me.Label2.Visible
   'Call VerficarProgramas(False)
      
   '*****************
   '* INSTALAR SQL EXPRESS 2005
   Dim sLocal As String
   sLocal = ProcuraArquivo(gLocalPathSetup, "SQLEXPR32.exe")
   If Trim(sLocal) <> "" Then 'ExisteArquivo(gLocalPathSetup & "03 Sql2005Express\SQLEXPR32.exe") Then
      bExiste = False
      sArq1 = Environ("Programfiles") & "\Microsoft SQL Server\MSSQL.1\MSSQL\Binn\sqlservr.exe"
      If ExisteArquivo(sArq1) Then
         bExiste = (Mid(GetFileVersion(sArq1), 1, 4) = "2005")
      End If
      If Not bExiste Then
         sArq1 = Environ("Programfiles") & "\Microsoft SQL Server\MSSQL.2005\MSSQL\Binn\sqlservr.exe"
         If ExisteArquivo(sArq1) Then
            bExiste = (Mid(GetFileVersion(sArq1), 1, 4) = "2005")
         End If
      End If
      If Not bExiste Then
         sArq1 = Environ("Programfiles") & "\Microsoft SQL Server\MSSQL\MSSQL\Binn\sqlservr.exe"
         If ExisteArquivo(sArq1) Then
            bExiste = (Mid(GetFileVersion(sArq1), 1, 4) = "2005")
         End If
         If Not bExiste Then
            sArq1 = Environ("Programfiles") & "\Microsoft SQL Server\MSSQL10.MSSQLSERVER\MSSQL\Binn\sqlservr.exe"
            If ExisteArquivo(sArq1) Then
               bExiste = (Mid(GetFileVersion(sArq1), 1, 4) = "2005")
            End If
         End If
      End If
      If Not bExiste Then
         sComando = sLocal & "\SQLEXPR32.exe /passive /X:C:\11SQL2005\"
         Call SincShell(sComando, vbNormalFocus, True)
         
         sComando = "C:\11SQL2005\Setup.exe /qb"
         sComando = sComando & " ADDLOCAL=ALL"
         sComando = sComando & " INSTANCENAME=SQLEXPRESS"
         sComando = sComando & " SECURITYMODE=SQL"
         sComando = sComando & " SAPWD=sqlexpress"
         sComando = sComando & " SQLCOLLATION=SQL_Latin1_General_CP1_CI_AI"
         sComando = sComando & " SQLAUTOSTART=1"
         sComando = sComando & " DISABLENETWORKPROTOCOLS=0"
         sComando = sComando & " ADDUSERASADMIN=1"
         
         Call SincShell(sComando, vbNormalFocus, True)
         
         Call ApagarDiretorio("C:\11SQL2005")
      End If
   End If
   'Me.Picture3.Visible = Me.Label3.Visible
   Call VerficarProgramas(False)
  
   '*****************
   '* INSTALAR TEAM VIEWER
   'bExiste = ExisteArquivo(Environ("Programfiles") & "\TeamViewer\Version6\TeamViewer.exe")
   'If Not bExiste Then
   '   sComando = gLocalPath & "05 TeamViewer\TeamViewer_Setup.exe /quiet"
   '   Call SincShell(sComando, vbNormalFocus, True)
   'End If
   'Me.Picture4.Visible = True
   'Call VerficarProgramas(False)
   
   '*****************
   '* INSTALAR PROJETO 3R
   If ExisteArquivo(gLocalPath & "P3R.msi") Then
      bExiste = ExisteArquivo(Environ("Programfiles") & "\ClasseA\Projeto3R\P3R.exe")
      
      sMsg = "Existe uma cópia do Projeto 3R instalado. " & vbNewLine '& vbNewLine
      sMsg = sMsg & "Esta operação irá desinstalar e instalar o Sistema novamente." & vbNewLine & vbNewLine
      sMsg = sMsg & "Deseja reinstalar? "
      If bExiste Then
         If vbYes = MsgBox(sMsg, vbQuestion + vbYesNo + vbDefaultButton1, "Projeto 3R") Then
            '* Se existe desinstala
            sComando = "msiexec /x """ & gLocalPath & "P3R.msi"" /qb /norestart"
            Call SincShell(sComando, vbNormalFocus, True)
         End If
      End If
               
      bExiste = ExisteArquivo(Environ("Programfiles") & "\ClasseA\Projeto3R\P3R.exe")
      If Not bExiste Then
         Call RegistrarDependencia
         
         sComando = "msiexec /i """ & gLocalPath & "P3R.msi""  /qb /norestart"
         Call SincShell(sComando, vbNormalFocus, True)
         
         Dim MyObj As Object
         Set MyObj = CreateObject("VersaoFTP.TL_VerifVersao")
         MyObj.AutoStart = True
         MyObj.ShowCAVs
      End If
   End If
   
   DoEvents
   Call ApagarDiretorio("C:\11SQL2005")
   
   '* Configurar Instancia SqlExpress
   sArq1 = Environ("Programfiles") & "\ClasseA\Admin\Dll\Setup.ini"
   If ExisteArquivo(sArq1) Then
      Call WriteIniFile(sArq1, "Database Format", "SERVER", " " & Environ("COMPUTERNAME") & "\SQLEXPRESS")
      Call WriteIniFile(sArq1, "AutoInstall Files", "Path", " " & "%programfiles%\ClasseA\Admin\DLL\")
      Call WriteIniFile(sArq1, "P3R AutoInstall Files", "Path", " " & "%programfiles%\ClasseA\Admin\DLL\")
   End If
   
   Me.Label7.Visible = Me.Label5.Visible
   
   '*****************
   '* CRIAR BANCO
   Dim sArq As String
   Dim i As Integer
   If gMetodo1 Then
      bExiste = ExisteArquivo(Environ("Programfiles") & "\Microsoft SQL Server\MSSQL\MSSQL\Data\G3R.mdf")
      bExiste = bExiste Or ExisteArquivo(Environ("Programfiles") & "\Microsoft SQL Server\MSSQL.1\MSSQL\Data\G3R.mdf")
      bExiste = bExiste Or ExisteArquivo(Environ("Programfiles") & "\Microsoft SQL Server\MSSQL.2005\MSSQL\Data\G3R.mdf")
      If bExiste Then
         sMsg = "##### ATENÇÃO! #####" & vbNewLine & vbNewLine
         sMsg = sMsg & "Existe uma instância do Banco de Dados em seu computador." & vbNewLine & vbNewLine
         sMsg = sMsg & "Esta operação irá excluir seu Banco de Dados e não haverá" & vbNewLine
         sMsg = sMsg & "possibilidade de recuperação." & vbNewLine & vbNewLine
         sMsg = sMsg & "Deseja continuar? "
         If vbYes = MsgBox(sMsg, vbQuestion + vbYesNo + vbDefaultButton2, "*** ATENÇÃO! ***") Then
            sMsg = "##### ATENÇÃO! #####" & vbNewLine & vbNewLine
            sMsg = sMsg & "A instância do Banco de Dados será excluída." & vbNewLine & vbNewLine
            sMsg = sMsg & "Tem certeza que quer continuar? "
            bExiste = (vbNo = MsgBox(sMsg, vbQuestion + vbYesNo + vbDefaultButton2, "*** ATENÇÃO! ***"))
         End If
      End If
   
      If Not bExiste Then
         Call DescompactarArquivo(gLocalPathSetup & "99 Database\", "G3R.zia", gLocalPathSetup & "99 Database\", False)
         Call ExecuteoSql(gLocalPathSetup & "99 Database\3R01_Banco.sql")
         Call AbreConn(0)
         Call ExecuteScript(xConn, gLocalPathSetup & "99 Database\3R02_Users.sql")
         Call AbreConn(1)
         Call ExecuteScript2(xConn, gLocalPathSetup & "99 Database\3R03_Defaults.sql")
         Call ExecuteoSql(gLocalPathSetup & "99 Database\3R04_Tabelas.sql")
         Call ExecuteScript(xConn, gLocalPathSetup & "99 Database\3R05_Insert.sql")
         Call ExecuteScript(xConn, gLocalPathSetup & "99 Database\3R06_InsertMenu.sql")
      End If
      Call ExcluirArquivo(gLocalPathSetup & "99 Database\3R01_Banco.sql")
      Call ExcluirArquivo(gLocalPathSetup & "99 Database\3R02_Users.sql")
      Call ExcluirArquivo(gLocalPathSetup & "99 Database\3R03_Defaults.sql")
      Call ExcluirArquivo(gLocalPathSetup & "99 Database\3R04_Tabelas.sql")
      Call ExcluirArquivo(gLocalPathSetup & "99 Database\3R05_Insert.sql")
      Call ExcluirArquivo(gLocalPathSetup & "99 Database\3R06_InsertMenu.sql")
   Else
      Dim sPathSQL As String
      Dim Sql  As String
      
      If ExisteArquivo(Environ("Programfiles") & "\Microsoft SQL Server\MSSQL\MSSQL\Data\master.mdf") Then
         sPathSQL = Environ("Programfiles") & "\Microsoft SQL Server\MSSQL\MSSQL\Data\"
      End If
      If ExisteArquivo(Environ("Programfiles") & "\Microsoft SQL Server\MSSQL.1\MSSQL\Data\master.mdf") Then
         sPathSQL = Environ("Programfiles") & "\Microsoft SQL Server\MSSQL.1\MSSQL\Data\"
      End If
      If ExisteArquivo(Environ("Programfiles") & "\Microsoft SQL Server\MSSQL.2\MSSQL\Data\master.mdf") Then
         sPathSQL = Environ("Programfiles") & "\Microsoft SQL Server\MSSQL.2\MSSQL\Data\"
      End If
      If Not ExisteArquivo(sPathSQL & "G3R.mdf") Then
         sLocal = ProcuraArquivo(gLocalPathSetup, "G3R.mdf")
         If Trim(sLocal) <> "" Then
            Call CopiarArquivo(sLocal & "G3R.mdf", sPathSQL & "G3R.mdf")
         End If
      End If
      If Not ExisteArquivo(sPathSQL & "G3R.ldf") Then
         sLocal = ProcuraArquivo(gLocalPathSetup, "G3R.ldf")
         If Trim(sLocal) <> "" Then
            Call CopiarArquivo(sLocal & "G3R.ldf", sPathSQL & "G3R.ldf")
         End If
      End If
      
      Call AbreConn(0, "master")
      Sql = ""
      Sql = Sql & "USE [master];" & vbNewLine
      Sql = Sql & "CREATE DATABASE [G3R] ON"
      Sql = Sql & " ( FILENAME = N'" & sPathSQL & "G3R.mdf' ),"
      Sql = Sql & " ( FILENAME = N'" & sPathSQL & "G3R.ldf' )"
      Sql = Sql & " FOR ATTACH;" & vbNewLine
      xConn.Execute Sql
      sLocal = ProcuraArquivo(gLocalPathSetup, "3R02_Users.sql")
      If Trim(sLocal) <> "" Then
         Call ExecuteScript(xConn, sLocal & "3R02_Users.sql")
      End If
   End If
   Me.Label7.Caption = "Banco de dados criado."
   Me.Label7.Visible = True
   
   'Me.Picture5.Visible = Me.Label5.Visible
   Call VerficarProgramas(False)
   
   bExiste = ExisteArquivo(Environ("Programfiles") & "\ClasseA\Projeto3R\P3R.exe")
   If bExiste Then
      sComando = Environ("Programfiles") & "\ClasseA\Projeto3R\P3R.exe"
      Call SincShell(sComando, vbNormalFocus, False)
   End If
   
   'Me.Picture4.Visible = Me.Label4.Visible
   Call VerficarProgramas(False)
   
   Me.ProgressBar1.Scrolling = 0
   Me.ProgressBar1.Visible = False
   Screen.MousePointer = vbDefault
   
   
   '*****************
   '* INSTALAR SQL MANAGEMENT STUDIO
   bExiste = False
   sArq1 = Environ("Programfiles") & "\Microsoft SQL Server\100\Tools\Binn\VSShell\Common7\IDE\Ssms.exe"
   sArq2 = Environ("Programfiles") & "\Microsoft SQL Server\90\Tools\Binn\VSShell\Common7\IDE\ssmsee.exe"
   bExiste = ExisteArquivo(sArq1) Or ExisteArquivo(sArq2)
   If Not bExiste Then
      Call BaixarArquivo("P04.zia")
      If ExisteArquivo(gLocalPathSetup & "P04.zia") Then
         Call DescompactarArquivo(gLocalPathSetup, "P04.zia", gLocalPathSetup)
      End If
      '*****************
      '* INSTALAR SQL MANAGEMENT STUDIO
      
       #If Win32 Then
         sLocal = ProcuraArquivo(gLocalPathSetup, "SQLServer2005_SSMSEE.msi")
         If Trim(sLocal) <> "" Then
            bExiste = False
            sArq1 = Environ("Programfiles") & "\Microsoft SQL Server\100\Tools\Binn\VSShell\Common7\IDE\Ssms.exe"
            sArq2 = Environ("Programfiles") & "\Microsoft SQL Server\90\Tools\Binn\VSShell\Common7\IDE\ssmsee.exe"
            bExiste = ExisteArquivo(sArq1) Or ExisteArquivo(sArq2)
            If ExisteArquivo(sLocal & "SQLServer2005_SSMSEE.msi") Then
               If Not bExiste Then
                  sComando = "msiexec /i """ & sLocal & "SQLServer2005_SSMSEE.msi"" /qb"
                  Call SincShell(sComando, vbNormalFocus, True)
                  bExiste = False
                  sArq1 = Environ("Programfiles") & "\Microsoft SQL Server\100\Tools\Binn\VSShell\Common7\IDE\Ssms.exe"
                  sArq2 = Environ("Programfiles") & "\Microsoft SQL Server\90\Tools\Binn\VSShell\Common7\IDE\ssmsee.exe"
                  bExiste = ExisteArquivo(sArq1) Or ExisteArquivo(sArq2)
                  If ExisteArquivo(sLocal & "SQLServer2005_SSMSEE_x64.msi") Then
                     If Not bExiste Then
                        sComando = "msiexec /i """ & sLocal & "SQLServer2005_SSMSEE_x64.msi"" /qb"
                        Call SincShell(sComando, vbNormalFocus, True)
                     End If
                  End If
                  
               End If
            End If
          End If
      #Else
         sLocal = ProcuraArquivo(gLocalPathSetup, "SQLServer2005_SSMSEE_x64.msi")
         If Trim(sLocal) <> "" Then
            bExiste = False
            sArq1 = Environ("Programfiles") & "\Microsoft SQL Server\100\Tools\Binn\VSShell\Common7\IDE\Ssms.exe"
            sArq2 = Environ("Programfiles") & "\Microsoft SQL Server\90\Tools\Binn\VSShell\Common7\IDE\ssmsee.exe"
            bExiste = ExisteArquivo(sArq1) Or ExisteArquivo(sArq2)
            If ExisteArquivo(sLocal & "SQLServer2005_SSMSEE_x64.msi") Then
               If Not bExiste Then
                  sComando = "msiexec /i """ & sLocal & "SQLServer2005_SSMSEE_x64.msi"" /qb"
                  Call SincShell(sComando, vbNormalFocus, True)
               End If
            End If
         End If
      #End If
   End If
Exit Sub
Saida:
   MsgBox Err & " - " & Error, vbCritical, "Aviso/Erro em Instalar Ferramentas"
   Resume Next
End Sub
Private Sub VerficarProgramas(Optional pInicio As Boolean)
   Dim bExiste  As Boolean
   Dim sArq1     As String
   Dim sArq2     As String
   
   '*****************
   '* INSTALAR WINDOWS INSTALLER
   sArq1 = Environ("SystemRoot") & "\system32\msi.dll"
   If ExisteArquivo(sArq1) Then
      bExiste = (Mid(GetFileVersion(sArq1), 1, 7) >= "004.005")
   End If
   If bExiste Then
      Me.Picture1.Visible = Me.Label1.Visible
      Me.Picture1.Refresh
      Me.Refresh
      DoEvents
   Else
      Call BaixarArquivo("P01.zia")
   End If
         
   '*****************
   '* INSTALAR FRAMEWORK 3.5
   bExiste = False
   sArq1 = Environ("SystemRoot") & "\Microsoft.NET\Framework\v2*"
   bExiste = (Dir(sArq1, vbDirectory) <> "")
   sArq1 = Environ("SystemRoot") & "\Microsoft.NET\Framework\v3.5*"
   sArq2 = Environ("SystemRoot") & "\Microsoft.NET\Framework\v4*"
   bExiste = bExiste Or (Dir(sArq1, vbDirectory) <> "" Or Dir(sArq2, vbDirectory) <> "")
   If bExiste Then
      Me.Picture2.Visible = Me.Label2.Visible
      Me.Picture2.Refresh
      Me.Refresh
      DoEvents
   Else
      Call BaixarArquivo("P02.zia")
   End If
      
   '*****************
   '* INSTALAR SQL EXPRESS 2005
   bExiste = False
   sArq1 = Environ("Programfiles") & "\Microsoft SQL Server\MSSQL.1\MSSQL\Binn\sqlservr.exe"
   If ExisteArquivo(sArq1) Then
      bExiste = (Mid(GetFileVersion(sArq1), 1, 4) = "2005")
   End If
   If Not bExiste Then
      sArq1 = Environ("Programfiles") & "\Microsoft SQL Server\MSSQL.2005\MSSQL\Binn\sqlservr.exe"
      If ExisteArquivo(sArq1) Then
         bExiste = (Mid(GetFileVersion(sArq1), 1, 4) = "2005")
      End If
   End If
   If Not bExiste Then
      sArq1 = Environ("Programfiles") & "\Microsoft SQL Server\MSSQL\MSSQL\Binn\sqlservr.exe"
      If ExisteArquivo(sArq1) Then
         bExiste = (Mid(GetFileVersion(sArq1), 1, 4) = "2005")
      End If
      If Not bExiste Then
         sArq1 = Environ("Programfiles") & "\Microsoft SQL Server\MSSQL10.MSSQLSERVER\MSSQL\Binn\sqlservr.exe"
         If ExisteArquivo(sArq1) Then
            bExiste = (Mid(GetFileVersion(sArq1), 1, 4) = "2005")
         End If
      End If
   End If
   If bExiste Then
      Me.Picture3.Visible = Me.Label3.Visible
      Me.Picture3.Refresh
      Me.Refresh
   Else
      Call BaixarArquivo("P03.zia")
   End If
      
   '*****************
   '* INSTALAR SQL MANAGEMENT STUDIO
   bExiste = False
   sArq1 = Environ("Programfiles") & "\Microsoft SQL Server\100\Tools\Binn\VSShell\Common7\IDE\Ssms.exe"
   sArq2 = Environ("Programfiles") & "\Microsoft SQL Server\90\Tools\Binn\VSShell\Common7\IDE\ssmsee.exe"
   bExiste = ExisteArquivo(sArq1) Or ExisteArquivo(sArq2)
   If bExiste Then
      Me.Picture4.Visible = Me.Label4.Visible
      Me.Picture4.Visible = Me.Label3.Visible
      Me.Picture4.Refresh
      Me.Refresh
   Else
      'Call BaixarArquivo("P04.zia")
   End If
     
   '*****************
   '* INSTALAR PROJETO 3R
   bExiste = ExisteArquivo(Environ("Programfiles") & "\ClasseA\Projeto3R\P3R.exe")
   If bExiste Then
      Me.Picture5.Visible = Me.Label5.Visible
   Else
      Call BaixarArquivo("MSI.zia")
   End If
   bExiste = ExisteArquivo(Environ("Programfiles") & "\Microsoft SQL Server\MSSQL\MSSQL\Data\G3R.mdf")
   bExiste = bExiste Or ExisteArquivo(Environ("Programfiles") & "\Microsoft SQL Server\MSSQL.1\MSSQL\Data\G3R.mdf")
   bExiste = bExiste Or ExisteArquivo(Environ("Programfiles") & "\Microsoft SQL Server\MSSQL.2005\MSSQL\Data\G3R.mdf")
   If bExiste Then
      Me.Label7.Visible = Me.Label5.Visible
      Me.Label7.Caption = "Banco de dados criado."
      Me.Label7.Visible = True
   Else
      Call BaixarArquivo("P99.zia")
   End If
   If pInicio Then
      Call CmdInicio_Click
   End If
End Sub
Private Sub Form_Activate()
   If Me.Tag = "" Then
      Me.Tag = "1"
      If Not gMetodo1 Then
         Me.Label6.Caption = "Os arquivos necessários serão baixados ao longo da instalação. Caso ocorra uma parada inesperada como um congelamento de tela, aguarde um pouco pois a velocidade de sua internet influenciará diretamente na sua instalação."
         Me.LblTit2.Visible = False
         Call VerficarProgramas(True)
      End If
      'Call CmdInicio_Click
   End If
End Sub

Private Sub Form_Load()
   Me.Top = 0
   Me.Left = 0
   If Not gMetodo1 Then
      Me.CmdInicio.Visible = False
      Me.ProgressBar1.Visible = True
      Me.ProgressBar1.Scrolling = 2
   End If
End Sub
Private Sub AbreConn(pTipo As Integer, Optional pDbName As String)
   Dim sConect As String
   Dim Sql     As String
    
   Dim sDbName As String
   Dim sServer As String
    
   If pDbName = "" Then
      sDbName = "G3R"
   Else
      sDbName = pDbName
   End If
   sServer = Environ("computername") & "\SQLEXPRESS"
   
   If pTipo = 0 Then
      sConect = "Provider=SQLOLEDB;"
      sConect = sConect & "Initial Catalog=" & sDbName & ";"
      sConect = sConect & "Data Source=" & sServer & ";"
      sConect = sConect & "Integrated Security=SSPI;"
   ElseIf pTipo = 1 Then
      sConect = "Provider=SQLOLEDB;"
      sConect = sConect & "Initial Catalog=" & sDbName & ";"
      sConect = sConect & "Data Source=" & sServer & ";"
      sConect = sConect & "User Id=USU_VERIF;"
      sConect = sConect & "Password=MINOTAURO;"
   End If
   
   If Not xConn Is Nothing Then
      If xConn.state = 1 Then
         xConn.Close
      End If
   End If
   Set xConn = CreateObject("ADODB.Connection") ' New ADODB.Connection
   With xConn
      .CommandTimeout = 300
      .CursorLocation = 3
      .ConnectionString = sConect
      .Open
   End With
'Sql = "CREATE PROCEDURE dbo.uspSetSQLServerAuthenticationMode" & vbNewLine
'Sql = Sql & "(" & vbNewLine
'Sql = Sql & "       @MixedMode BIT" & vbNewLine
'Sql = Sql & ")" & vbNewLine
'Sql = Sql & "AS" & vbNewLine
'Sql = Sql & " " & vbNewLine
'Sql = Sql & "SET NOCOUNT ON" & vbNewLine
'Sql = Sql & " " & vbNewLine
'Sql = Sql & "DECLARE @InstanceName NVARCHAR(1000)," & vbNewLine
'Sql = Sql & "       @Key NVARCHAR(4000)," & vbNewLine
'Sql = Sql & "       @NewLoginMode INT," & vbNewLine
'Sql = Sql & "       @OldLoginMode INT" & vbNewLine
'Sql = Sql & " " & vbNewLine
'Sql = Sql & "EXEC master..xp_regread    N'HKEY_LOCAL_MACHINE'," & vbNewLine
'Sql = Sql & "                     n 'Software\Microsoft\Microsoft SQL Server\Instance Names\SQL\'," & vbNewLine
'Sql = Sql & "                     n 'MSSQLSERVER'," & vbNewLine
'Sql = Sql & "                     @InstanceName OUTPUT" & vbNewLine
'Sql = Sql & " " & vbNewLine
'Sql = Sql & "IF @@ERROR <> 0 OR @InstanceName IS NULL" & vbNewLine
'Sql = Sql & "       BEGIN" & vbNewLine
'Sql = Sql & "              RAISERROR('Could not read SQL Server instance name.', 18, 1)" & vbNewLine
'Sql = Sql & "              RETURN -100" & vbNewLine
'Sql = Sql & "       End" & vbNewLine
'Sql = Sql & " " & vbNewLine
'Sql = Sql & "SET    @Key = N'Software\Microsoft\Microsoft SQL Server\' + @InstanceName + N'\MSSQLServer\'" & vbNewLine
'Sql = Sql & " " & vbNewLine
'Sql = Sql & "EXEC master..xp_regread    N'HKEY_LOCAL_MACHINE'," & vbNewLine
'Sql = Sql & "                     @Key," & vbNewLine
'Sql = Sql & "                     n 'LoginMode'," & vbNewLine
'Sql = Sql & "                     @OldLoginMode OUTPUT" & vbNewLine
'Sql = Sql & " " & vbNewLine
'Sql = Sql & "IF @@ERROR <> 0" & vbNewLine
'Sql = Sql & "       BEGIN" & vbNewLine
'Sql = Sql & "              RAISERROR('Could not read login mode for SQL Server instance %s.', 18, 1, @InstanceName)" & vbNewLine
'Sql = Sql & "              RETURN -110" & vbNewLine
'Sql = Sql & "       End" & vbNewLine
'Sql = Sql & " " & vbNewLine
'Sql = Sql & "IF @MixedMode IS NULL" & vbNewLine
'Sql = Sql & "       BEGIN" & vbNewLine
'Sql = Sql & "              RAISERROR('No change to authentication mode was made. Login mode is %d.', 10, 1, @OldLoginMode)" & vbNewLine
'Sql = Sql & "              RETURN -120" & vbNewLine
'Sql = Sql & "       End" & vbNewLine
'Sql = Sql & " " & vbNewLine
'Sql = Sql & "IF @MixedMode = 1" & vbNewLine
'Sql = Sql & "       SET    @NewLoginMode = 2" & vbNewLine
'Sql = Sql & "Else" & vbNewLine
'Sql = Sql & "       SET    @NewLoginMode = 1" & vbNewLine
'Sql = Sql & " " & vbNewLine
'Sql = Sql & "EXEC master..xp_regwrite   N'HKEY_LOCAL_MACHINE'," & vbNewLine
'Sql = Sql & "                           @Key," & vbNewLine
'Sql = Sql & "                           n 'LoginMode'," & vbNewLine
'Sql = Sql & "                           'REG_DWORD'," & vbNewLine
'Sql = Sql & "                           @NewLoginMode" & vbNewLine
'Sql = Sql & " " & vbNewLine
'Sql = Sql & "IF @@ERROR <> 0" & vbNewLine
'Sql = Sql & "       BEGIN" & vbNewLine
'Sql = Sql & "              RAISERROR('Could not write login mode %d for SQL Server instance %s. Login mode is %d', 18, 1, @NewLoginMode, @InstanceName, @OldLoginMode)" & vbNewLine
'Sql = Sql & "              RETURN -130" & vbNewLine
'Sql = Sql & "       End" & vbNewLine
'Sql = Sql & " " & vbNewLine
'Sql = Sql & "RAISERROR('Login mode is now %d for SQL Server instance %s. Login mode was %d before.', 10, 1, @NewLoginMode, @InstanceName, @OldLoginMode)" & vbNewLine
'Sql = Sql & "RETURN 0" & vbNewLine
'xConn.Execute Sql
   
End Sub
'Private Sub MontarBanco()
'   Dim sConect As String
'   Dim Sql     As String
'
'   Dim sDbName As String
'   Dim sServer As String
'
'      If .state = 1 Then
'         'Call ExecuteScript(xConn, gLocalPath & "3R01_Banco.sql")
'         Call ExecuteScript(xConn, gLocalPath & "3R02_Users.sql")
'
'      End If
'      .Close
'   End With
'
'   sConect = "Provider=SQLOLEDB;"
'   sConect = sConect & "Initial Catalog=" & sDbName & ";"
'   sConect = sConect & "Data Source=" & sServer & ";"
'   sConect = sConect & "User Id=USU_VERIF;"
'   sConect = sConect & "Password=MINOTAURO;"
'   Set xConn = CreateObject("ADODB.Connection") ' New ADODB.Connection
'   With xConn
'      .CommandTimeout = 300
'      .CursorLocation = 3
'      .ConnectionString = sConect
'      .Open
'      'gLocalPath =Environ ("ProgramFiles") & "\ClasseA\Admin\Instalacao\Setup"
'      If .state = 1 Then
'         Call ExecuteScript2(xConn, gLocalPath & "3R03_Defaults.sql")
'         'Call ExecuteScript(xConn, gLocalPath & "3R04_Tabelas01.sql")
'         Call ExecuteScript(xConn, gLocalPath & "3R04_Tabelas00.sql")
'         Call ExecuteScript(xConn, gLocalPath & "3R04_Tabelas02.sql")
'         Call ExecuteScript(xConn, gLocalPath & "3R04_Tabelas03.sql")
'         Call ExecuteScript(xConn, gLocalPath & "3R04_Tabelas04.sql")
'         Call ExecuteScript(xConn, gLocalPath & "3R04_Tabelas05.sql")
'         Call ExecuteScript(xConn, gLocalPath & "3R04_Tabelas06.sql")
'         Call ExecuteScript(xConn, gLocalPath & "3R04_Tabelas07.sql")
'         Call ExecuteScript2(xConn, gLocalPath & "3R05_Insert.sql")
'         Call ExecuteScript2(xConn, gLocalPath & "3R06_InsertMenu.sql")
'      End If
'   End With
'End Sub
Private Sub ExecuteScript(xConn As Object, pPathFile As String)
   Dim Sql As String
   
   If ExisteArquivo(pPathFile) Then
      Sql = ReadTextFile(pPathFile)
      Sql = Replace(Sql, Chr(239), "")
      Sql = Replace(Sql, Chr(187), "")
      Sql = Replace(Sql, Chr(191), "")
      xConn.Execute Sql
   End If
End Sub
Private Sub ExecuteScript2(xConn As Object, pPathFile As String)
   Dim Sql As String
   Dim SqlAux As String
   
   If ExisteArquivo(pPathFile) Then
      Sql = ReadTextFile(pPathFile)
      Sql = Replace(Sql, Chr(239), "")
      Sql = Replace(Sql, Chr(187), "")
      Sql = Replace(Sql, Chr(191), "")
      While InStr(Sql, ";")
         SqlAux = Mid(Sql, 1, InStr(Sql, ";"))
         xConn.Execute SqlAux
         Sql = Mid(Sql, InStr(Sql, ";") + 1)
      Wend
   End If
End Sub
Private Sub ExecuteoSql(pPathFile As String)
   Dim sComando As String
   Dim i As Integer

   i = Val(Mid(Right(pPathFile, 6), 1, 2)) + 1
   If ExisteArquivo(pPathFile) Then
      sComando = "osql -E -S "
      sComando = sComando & Environ("COMPUTERNAME") & "\SQLEXPRESS"
      sComando = sComando & " -i "
      sComando = sComando & """" & pPathFile & """"
      sComando = sComando & " -o """ & gLocalPathSetup & "99 Database\" & "Result" & Right("00" & i, 2) & ".txt" & """"
   
      Call SincShell(sComando, vbHide, True)
   End If
End Sub

