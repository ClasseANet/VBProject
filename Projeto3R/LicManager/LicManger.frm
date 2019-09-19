VERSION 5.00
Object = "{3F666E72-7F79-447A-BCFF-C42C44E3DBE7}#1.0#0"; "MCIGrid.ocx"
Begin VB.Form FrmLicManger 
   Caption         =   "FrmLicManger"
   ClientHeight    =   10350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15915
   LinkTopic       =   "Form1"
   ScaleHeight     =   10350
   ScaleWidth      =   15915
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdEnviarArq 
      Caption         =   "Enviar"
      Height          =   375
      Left            =   8760
      TabIndex        =   13
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton CmdSalvarArq 
      Caption         =   "Salvar"
      Height          =   375
      Left            =   7200
      TabIndex        =   12
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton CmdEncripto 
      Caption         =   "Encriptografa"
      Height          =   375
      Left            =   7200
      TabIndex        =   11
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton CmdDecripta 
      Caption         =   "Decriptografa"
      Height          =   375
      Left            =   7200
      TabIndex        =   10
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox TxtResult 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7200
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "LicManger.frx":0000
      Top             =   5520
      Width           =   8535
   End
   Begin VB.TextBox TxtDecripto 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7200
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "LicManger.frx":0006
      Top             =   3480
      Width           =   8535
   End
   Begin VB.TextBox TxtCripto 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7200
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "LicManger.frx":000C
      Top             =   1320
      Width           =   8535
   End
   Begin VB.CommandButton CmdEnviar 
      Caption         =   "Enviar Lics"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   9840
      Width           =   975
   End
   Begin VB.CommandButton CmdRenovar 
      Caption         =   "Renovar"
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   9840
      Width           =   975
   End
   Begin VB.CommandButton CmdSair 
      Caption         =   "Sair"
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   9840
      Width           =   975
   End
   Begin VB.CommandButton CmdAtualizar 
      Caption         =   "Baixar .Lics"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   9840
      Width           =   975
   End
   Begin VB.CommandButton CmdConectar 
      Caption         =   "Conectar"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   9840
      Width           =   975
   End
   Begin VB.CommandButton CmdSalvar 
      Caption         =   "Salvar"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   9840
      Width           =   975
   End
   Begin iGrid251_75B4A91C.iGrid GrdColigadas 
      Height          =   9495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   16748
   End
   Begin VB.Label LblFTP02 
      Caption         =   "FTP 02:"
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
      Left            =   7200
      TabIndex        =   15
      Top             =   600
      Width           =   8295
   End
   Begin VB.Label LblFTP01 
      Caption         =   "FTP 01:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   7200
      TabIndex        =   14
      Top             =   240
      Width           =   8295
   End
End
Attribute VB_Name = "FrmLicManger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oFtpPadrao  As Object
Dim oFtp1       As Object
Dim oFtp2       As Object
Dim FtpBakPath As String
Dim sTagCOL As String

Private Sub CmdAtualizar_Click()
   Screen.MousePointer = vbHourglass
   Call CarregarLic(True)
   Screen.MousePointer = vbDefault
End Sub

Private Sub CmdConectar_Click()
   DoEvents
   Screen.MousePointer = vbHourglass
   Me.MousePointer = vbHourglass
   
   On Error Resume Next
   oFtp1.DesconectarFTP
   oFtp2.DesconectarFTP
   Set oFtp1 = Nothing
   Set oFtp2 = Nothing
   
   If ConectarFtp1 Or ConectarFtp2 Then
      Call CarregarLic
      Me.CmdSalvar.Enabled = True
      Me.CmdEnviar.Enabled = True
      Me.CmdAtualizar.Enabled = True
   End If
   Me.MousePointer = vbDefault
   Screen.MousePointer = vbDefault
End Sub
Private Sub CarregarLic(Optional bRefresh As Boolean = False)
   Dim sLocalPath As String
   Dim sArqLic    As String
   Dim sTag       As String
   Dim sDTLIC     As String
   Dim nNUMLIC    As Integer
   Dim i          As Integer
   Dim cArqs      As Collection
   Dim n          As Variant
   Dim sNMCOL     As String
   Dim bBaixou    As Boolean
   Dim j          As Integer
   Dim sNMCOL0    As String
   Dim nAux       As Long
     
   sLocalPath = ResolvePathName(App.Path)
   If Trim(sLocalPath) = "" Then Exit Sub
   
   Dim xCol As CellObject
   Me.GrdColigadas.Clear True
   Me.GrdColigadas.AddCol "IDCOLIGADA", "#", lWidth:=20
   Me.GrdColigadas.AddCol "NMCOLIGADA", "COLIGADAS", lWidth:=120
   Me.GrdColigadas.AddCol "DTLIC", "LICENÇAS", lWidth:=80, eHdrTextFlags:=igTextCenter
   'xCol.eType = igCellCustomDraw
   
   sNMCOL0 = GetTag(sTagCOL, "COLIGADA" & i, "")
   sArqLic = sNMCOL & ".lic"
   Me.GrdColigadas.Clear False
   
   If bRefresh Then
      bBaixou = False
      If oFtp1 Is Nothing Then ConectarFtp1
      If oFtp2 Is Nothing Then ConectarFtp2
      
      If oFtp1.Conectado Or oFtp2.Conectado Then
         Me.CmdSalvar.Enabled = True
         Me.CmdEnviar.Enabled = True
      Else
        Call CmdConectar_Click
        If oFtp1.Conectado Or oFtp2.Conectado Then
            Me.CmdSalvar.Enabled = True
            Me.CmdEnviar.Enabled = True
        End If
      End If
   End If
   If Not (oFtp1 Is Nothing And oFtp2 Is Nothing) Or Not bRefresh Then
'      If oFtp1.Conectado Or oFtp2.Conectado Then
         For j = 1 To 2
            If j = 1 Then Set oFtpPadrao = oFtp1
            If j = 2 Then Set oFtpPadrao = oFtp2
            i = 0
            sNMCOL = sNMCOL0
            sArqLic = sNMCOL & ".lic"
            
            While sNMCOL <> ""
               sDTLIC = ""
               nNUMLIC = 0
               bBaixou = True
               If bRefresh Then
                  bBaixou = False
                  If oFtpPadrao.Conectado Then
                     bBaixou = oFtpPadrao.BaixarArquivo(FtpBakPath, sArqLic, sLocalPath, sArqLic)
                  End If
               End If
               If ExisteArquivo(sLocalPath & sArqLic) Then
                  sTag = ReadTextFile(sLocalPath & sArqLic)
                  If UCase(sNMCOL) = UCase(GetTag(Decrypt2(sTag), "NMCOLIGADA", "")) Then
                     sDTLIC = GetTag(Decrypt2(sTag), "DTLINC", "")
                     nNUMLIC = xVal(GetTag(Decrypt2(sTag), "NUMLINC", 0))
                  End If
               End If
               
               nAux = Me.GrdColigadas.FindSearchMatchRow(2, sNMCOL)
               If nAux = 0 Then
                  Me.GrdColigadas.AddRow
                  Me.GrdColigadas.CellValue(i + 1, "IDCOLIGADA") = i
                  Me.GrdColigadas.CellValue(i + 1, "NMCOLIGADA") = sNMCOL
                  Me.GrdColigadas.CellValue(i + 1, "DTLIC") = sDTLIC
                  Me.GrdColigadas.CellTextFlags(i + 1, "DTLIC") = igTextCenter
               End If
               Me.GrdColigadas.CellForeColor(i + 1, "IDCOLIGADA") = IIf(bRefresh And bBaixou, vbBlack, &H808080)
               Me.GrdColigadas.CellForeColor(i + 1, "NMCOLIGADA") = IIf(bRefresh And bBaixou, vbBlack, &H808080)
               If bBaixou Then
                  Me.GrdColigadas.CellForeColor(i + 1, "DTLIC") = vbBlack
                  If IsDate(sDTLIC) Then
                     If DateDiff("d", Now(), CDate(sDTLIC)) <= 5 Then
                        Me.GrdColigadas.CellForeColor(i + 1, "DTLIC") = vbRed
                     Else
                        If Month(sDTLIC) = Month(Now()) Then
                           Me.GrdColigadas.CellForeColor(i + 1, "DTLIC") = vbBlack
                        Else
                           Me.GrdColigadas.CellForeColor(i + 1, "DTLIC") = vbBlue
                        End If
                     End If
                  End If
               Else
                  Me.GrdColigadas.CellForeColor(i + 1, "DTLIC") = &H808080
               End If
               i = i + 1
               sNMCOL = GetTag(sTagCOL, "COLIGADA" & i, "")
               sArqLic = sNMCOL & ".lic"
            Wend
            If Not oFtpPadrao Is Nothing Then
               If oFtpPadrao.Conectado Then
                  Set cArqs = oFtpPadrao.ListaDirRemoto(FtpBakPath)
                  For Each n In cArqs
                     sArqLic = n
                     sNMCOL = Mid(sArqLic, 1, Len(sArqLic) - 4)
                     If InStr(sTagCOL, sNMCOL & "|") = 0 Then
                        Call oFtpPadrao.BaixarArquivo(FtpBakPath, sArqLic, sLocalPath, sArqLic)
                        If ExisteArquivo(sLocalPath & sArqLic) Then
                           sTag = ReadTextFile(sLocalPath & sArqLic)
                           sDTLIC = GetTag(Decrypt2(sTag), "DTLINC", "")
                           nNUMLIC = xVal(GetTag(Decrypt2(sTag), "NUMLINC", 0))
                        End If
                        
                        i = Me.GrdColigadas.RowCount
                        Me.GrdColigadas.AddRow
                        Me.GrdColigadas.CellValue(i + 1, "IDCOLIGADA") = "x"
                        Me.GrdColigadas.CellValue(i + 1, "NMCOLIGADA") = sNMCOL
                        Me.GrdColigadas.CellValue(i + 1, "DTLIC") = sDTLIC
                     End If
                  Next
               End If
            End If
         Next
'      End If
   End If
End Sub
Private Function ConectarFtp1() As Boolean
   Dim bIsWeb     As Boolean
   Dim FtpBak     As String
   Dim FtpBakUID  As String
   Dim FtpBakPWD  As String
   Dim sLocalPath As String
   Dim sArqLic    As String
   Dim bOk        As Boolean
   
   bIsWeb = IsWebConnected
   If Err <> 0 Then bIsWeb = True
   If Not bIsWeb Then Exit Function
    
   If Not oFtp1 Is Nothing Then
      If oFtp1.Conectado Then
         ConectarFtp1 = True
         Exit Function
      End If
   End If
   
'   FtpBak = "ftp.classeanet.com.br"
'   FtpBakUID = "classeanet"
'   FtpBakPWD = "ramos10"
'   FtpBakPath = "/private/Cliente/Dpil/"
   
   FtpBak = "ftp.classeaconsultoria.com.br"
   FtpBakUID = "clientedpil"
   FtpBakPWD = "Dpil10!0"
   FtpBakPath = ""
   
   If oFtp1 Is Nothing Then
      Set oFtp1 = CriarObjeto("VersaoFTP.TL_VerifVersao")
   End If

   With oFtp1
      If .ConectarFtp(FtpBak, FtpBakUID, FtpBakPWD, False) Then
         ConectarFtp1 = True
         Me.Caption = FtpBakUID & "@" & FtpBak & " [Conectado]"
         Me.LblFTP01.ForeColor = &H8000&
      Else
         Me.Caption = FtpBakUID & "@" & FtpBak & " [Desconectado]"
         Me.LblFTP01.ForeColor = vbRed
      End If
      Me.LblFTP01.Caption = UCase("FPT 01: " & Me.Caption)
      Me.LblFTP01.ToolTipText = FtpBakPWD
   End With
End Function
Private Function ConectarFtp2() As Boolean
   Dim bIsWeb     As Boolean
   Dim FtpBak     As String
   Dim FtpBakUID  As String
   Dim FtpBakPWD  As String
   Dim sLocalPath As String
   Dim sArqLic    As String
   Dim bOk        As Boolean
   
   bIsWeb = IsWebConnected
   If Err <> 0 Then bIsWeb = True
   If Not bIsWeb Then Exit Function
    
   If Not oFtp2 Is Nothing Then
      If oFtp2.Conectado Then
         ConectarFtp2 = True
         Exit Function
      End If
   End If
   
'   FtpBak = "ftp.classeanet.com.br"
'   FtpBakUID = "classeanet"
'   FtpBakPWD = "ramos10"
'   FtpBakPath = "/private/Cliente/Dpil/"
   
   FtpBak = "ftp.classeanet.com.br"
   FtpBakUID = "clientedpil"
   FtpBakPWD = "@Dpil10!0"
   FtpBakPath = ""
      
   If oFtp2 Is Nothing Then
      Set oFtp2 = CriarObjeto("VersaoFTP.TL_VerifVersao")
   End If
   With oFtp2
      If .ConectarFtp(FtpBak, FtpBakUID, FtpBakPWD, False) Then
         ConectarFtp2 = True
         Me.Caption = FtpBakUID & "@" & FtpBak & " [Conectado]"
         Me.LblFTP02.ForeColor = &H8000&
      Else
         Me.Caption = FtpBakUID & "@" & FtpBak & " [Desconectado]"
         Me.LblFTP02.ForeColor = vbRed
      End If
      Me.LblFTP02.Caption = UCase("FPT 02: " & Me.Caption)
      Me.LblFTP02.ToolTipText = FtpBakPWD
   End With
End Function

Private Sub CmdDecripta_Click()
   TxtDecripto.Text = Decrypt2(Me.TxtCripto.Text)
End Sub

Private Sub CmdEncripto_Click()
   Me.TxtResult.Text = Encrypt2(Me.TxtDecripto.Text)
End Sub

Private Sub CmdEnviar_Click()
   Call SalvarLic(False)
   Call EnviarLic
End Sub
Private Sub CmdEnviarArq_Click()
   Dim sLocalPath As String
   Dim sArqLic    As String
   Dim sTag       As String
   Dim sDTLIC     As String
   Dim nNUMLIC    As Integer
   Dim i As Integer
   Dim nResult1 As Double
   Dim nResult2 As Double
   Dim sNMCOL  As String
   
   Screen.MousePointer = vbHourglass
     
   sLocalPath = ResolvePathName(App.Path)
   If Trim(sLocalPath) = "" Then GoTo Saida
            
   sNMCOL = GetTag(sTagCOL, "COLIGADA" & i, "")
   sArqLic = sNMCOL & ".lic"
   nResult1 = True
   nResult2 = True

   sNMCOL = Me.GrdColigadas.CellValue(Me.GrdColigadas.CurRow, "NMCOLIGADA")
   sArqLic = sNMCOL & ".lic"
   sDTLIC = ""
   nNUMLIC = 0
                           

   If oFtp1 Is Nothing Then Call ConectarFtp1
   If oFtp2 Is Nothing Then Call ConectarFtp2
   If vbYes = ExibirPergunta("Atualizar " & sArqLic & " ?") Then
      If oFtp1.Conectado Then
         nResult1 = oFtp1.EnviarArquivo(sLocalPath, sArqLic, FtpBakPath, sArqLic, False)
      End If
      If oFtp2.Conectado Then
         nResult2 = oFtp2.EnviarArquivo(sLocalPath, sArqLic, FtpBakPath, sArqLic, False)
      End If
   End If
   nResult1 = nResult1 And True
   nResult2 = nResult2 And True
   
   MsgBox "Realizado!" & vbNewLine & _
          "   FTP 01: " & IIf(nResult1, "OK", "FALHOU") & vbNewLine & _
          "   FTP 02: " & IIf(nResult1, "OK", "FALHOU"), vbInformation + vbOKOnly, "Licença P3R"
Saida:
   Screen.MousePointer = vbDefault
End Sub

Private Sub CmdRenovar_Click()
   Dim i As Integer
   Dim sAux As String
   Call CmdConectar_Click
   Call CmdAtualizar_Click
   For i = 1 To Me.GrdColigadas.RowCount
      If IsNumeric(Me.GrdColigadas.CellValue(i, "IDCOLIGADA")) Then
         sAux = Format(DateAdd("d", 30, CDate(Me.GrdColigadas.CellValue(i, "DTLIC"))), "mm/yyyy")
         Me.GrdColigadas.CellValue(i, "DTLIC") = "25/" & sAux
      End If
   Next
   Call CmdEnviar_Click
End Sub

Private Sub CmdSair_Click()
   Screen.MousePointer = vbHourglass
   Unload Me
   End
End Sub

Private Sub CmdSalvar_Click()
   Call SalvarLic(True)
End Sub
Private Sub SalvarLic(bExibeMsg As Boolean)
   Dim sLocalPath As String
   Dim sArqLic    As String
   Dim sTag       As String
   Dim sDTLIC     As String
   Dim nNUMLIC    As Integer
   Dim i As Integer
   Dim nResult As Double
   Dim sNMCOL  As String
   
   Screen.MousePointer = vbHourglass
     
   sLocalPath = ResolvePathName(App.Path)
   If Trim(sLocalPath) = "" Then GoTo Saida
            
   sNMCOL = GetTag(sTagCOL, "COLIGADA" & i, "")
   sArqLic = sNMCOL & ".lic"
   'While sNMCOL <> ""
   For i = 0 To Me.GrdColigadas.RowCount - 1
      sNMCOL = Me.GrdColigadas.CellValue(i + 1, "NMCOLIGADA")
      sArqLic = sNMCOL & ".lic"
      sDTLIC = ""
      nNUMLIC = 0
                              
      If IsDate(Me.GrdColigadas.CellValue(i + 1, "DTLIC")) Then
         sTag = ReadTextFile(sLocalPath & sArqLic)
         If UCase(sNMCOL) = UCase(GetTag(Decrypt2(sTag), "NMCOLIGADA", "")) Then
            sDTLIC = GetTag(Decrypt2(sTag), "DTLINC", "")
            nNUMLIC = xVal(GetTag(Decrypt2(sTag), "NUMLINC", 0))
         End If
         If sDTLIC <> Me.GrdColigadas.CellValue(i + 1, "DTLIC") Then
            sTag = ""
            sTag = sTag & "|NMCOLIGADA=" & sNMCOL
            sTag = sTag & "|IDUSU=DIO"
            sTag = sTag & "|NUMLINC=1"
            sTag = sTag & "|DTLINC=" & Me.GrdColigadas.CellValue(i + 1, "DTLIC")
            sTag = sTag & "|"
            Call WriteTextFile(sLocalPath & sArqLic, Encrypt2(sTag))
         End If
      End If
      
      'i = i + 1
      'sNMCOL = GetTag(sTagCOL, "COLIGADA" & i, "")
      'sArqLic = sNMCOL & ".lic"
   Next
   'Wend

   'sArqLic = sNMCOL & ".lic"
   'Call .BaixarArquivo(FtpBakPath, sArqLic, sLocalPath, sArqLic)
   '|NMCOLIGADA=CASA|IDUSU=DIO|NUMLINC=1|DTLINC=30/04/2012
   If bExibeMsg Then
'      If nResult Then
         MsgBox "Realizado com sucesso!", vbInformation + vbOKOnly, "Licença P3R"
'      Else
'         MsgBox "Erro ao Salvar!", vbInformation + vbOKOnly, "Licença P3R"
'      End If
   End If
Saida:
   Screen.MousePointer = vbDefault
End Sub
Private Sub EnviarLic()
   Dim sLocalPath As String
   Dim sArqLic    As String
   Dim sTag       As String
   Dim sDTLIC     As String
   Dim nNUMLIC    As Integer
   Dim i As Integer
   Dim nResult1 As Double
   Dim nResult2 As Double
   Dim sNMCOL  As String
   
   Screen.MousePointer = vbHourglass
     
   sLocalPath = ResolvePathName(App.Path)
   If Trim(sLocalPath) = "" Then GoTo Saida
            
   sNMCOL = GetTag(sTagCOL, "COLIGADA" & i, "")
   sArqLic = sNMCOL & ".lic"
   'While sNMCOL <> ""
   nResult1 = True
   nResult2 = True
   For i = 0 To Me.GrdColigadas.RowCount - 1
      sNMCOL = Me.GrdColigadas.CellValue(i + 1, "NMCOLIGADA")
      sArqLic = sNMCOL & ".lic"
      sDTLIC = ""
      nNUMLIC = 0
                              
      If IsDate(Me.GrdColigadas.CellValue(i + 1, "DTLIC")) Then
         If oFtp1 Is Nothing Then Call ConectarFtp1
         If oFtp2 Is Nothing Then Call ConectarFtp2
         If vbYes = ExibirPergunta("Atualizar " & sArqLic & " ?") Then
            If oFtp1.Conectado Then
               nResult1 = oFtp1.EnviarArquivo(sLocalPath, sArqLic, FtpBakPath, sArqLic, False)
            End If
            If oFtp2.Conectado Then
               nResult2 = oFtp2.EnviarArquivo(sLocalPath, sArqLic, FtpBakPath, sArqLic, False)
            End If
         End If

         nResult1 = nResult1 And True
         nResult2 = nResult2 And True
      End If
      'i = i + 1
      'sNMCOL = GetTag(sTagCOL, "COLIGADA" & i, "")
      'sArqLic = sNMCOL & ".lic"
   Next
   'Wend

   'sArqLic = sNMCOL & ".lic"
   'Call .BaixarArquivo(FtpBakPath, sArqLic, sLocalPath, sArqLic)
   '|NMCOLIGADA=CASA|IDUSU=DIO|NUMLINC=1|DTLINC=30/04/2012
   
   MsgBox "Realizado!" & vbNewLine & _
          "   FTP 01: " & IIf(nResult1, "OK", "FALHOU") & vbNewLine & _
          "   FTP 02: " & IIf(nResult1, "OK", "FALHOU"), vbInformation + vbOKOnly, "Licença P3R"
   
   
'   If nResult Then
'      MsgBox "Realizado com sucesso!", vbInformation + vbOKOnly, "Licença P3R"
'   Else
'      MsgBox "Erro ao Salvar!", vbInformation + vbOKOnly, "Licença P3R"
'   End If
Saida:
   Screen.MousePointer = vbDefault
End Sub
Private Sub CmdSalvarArq_Click()
   Dim sLocalPath As String
   Dim sArqLic As String
   Dim sTag As String
   Dim sNMCOL As String
   Dim sDTLIC     As String
   Dim nNUMLIC    As Integer
   
   Screen.MousePointer = vbHourglass
     
   sLocalPath = ResolvePathName(App.Path)
   If Trim(sLocalPath) = "" Then
      Me.TxtCripto.Text = ""
      Me.TxtDecripto.Text = ""
      Me.TxtResult.Text = ""
      Screen.MousePointer = vbDefault
      Exit Sub
   End If

   sNMCOL = Me.GrdColigadas.CellValue(Me.GrdColigadas.CurRow, "NMCOLIGADA")
   sArqLic = sNMCOL & ".lic"
   If ExisteArquivo(sLocalPath & sArqLic) Then
      Call WriteTextFile(sLocalPath & sArqLic, Me.TxtResult.Text)
      MsgBox "Realizado com sucesso!", vbInformation + vbOKOnly, "Licença P3R"
   End If
   Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Activate()
   Me.Top = 60
End Sub

Private Sub Form_Load()
   Me.TxtCripto.Text = ""
   Me.TxtDecripto.Text = ""
   Me.TxtResult.Text = ""

   sTagCOL = ""
   sTagCOL = ReadTextFile(App.Path & "\COLIGADAS.TXT")
   sTagCOL = Replace(sTagCOL, Chr(10), "")
   sTagCOL = Replace(sTagCOL, Chr(13), "")
   If sTagCOL = "" Then
      sTagCOL = sTagCOL & "|COLIGADA0=CASA"
      sTagCOL = sTagCOL & "|COLIGADA1=LONDRINAI"
      sTagCOL = sTagCOL & "|COLIGADA2=LONDRINAII"
      sTagCOL = sTagCOL & "|COLIGADA3=CIDADEDUTRA"
      sTagCOL = sTagCOL & "|COLIGADA4=SAOROQUE"
      sTagCOL = sTagCOL & "|COLIGADA5=PATROCINIO"
      sTagCOL = sTagCOL & "|COLIGADA6=LARANJEIRAS"
      sTagCOL = sTagCOL & "|COLIGADA7=ROLANDIA"
      sTagCOL = sTagCOL & "|COLIGADA8=BELAVISTAGO"
      sTagCOL = sTagCOL & "|COLIGADA9=JARDIMGOIAS"
      sTagCOL = sTagCOL & "|COLIGADA10=JATAI"
      sTagCOL = sTagCOL & "|COLIGADA11=UBERABA"
      sTagCOL = sTagCOL & "|COLIGADA12=UBERLANDIA"
      sTagCOL = sTagCOL & "|COLIGADA13=CIANORTE"
      sTagCOL = sTagCOL & "|COLIGADA14=JDBELAVISTA"
      sTagCOL = sTagCOL & "|COLIGADA15=PINHEIROS2"
      sTagCOL = sTagCOL & "|"
   Else
      If Right(sTagCOL, 1) <> "|" Then
         sTagCOL = sTagCOL & "|"
      End If
   End If
   Call CarregarLic
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   oFtp1.DesconectarFTP
   oFtp2.DesconectarFTP
   Set oFtp1 = Nothing
   Set oFtp2 = Nothing
   Screen.MousePointer = vbDefault
End Sub
Private Function WriteTextFile(pFile As String, sText As String) As Boolean
   Dim intLen As Integer
   
   On Error GoTo Saida
'   If ExisteArquivo(pFile) Then
      intLen = FreeFile
      Open pFile For Output As #intLen
      Print #intLen, sText
      Close #intLen
'   End If
   WriteTextFile = (intLen > 0)
Saida:
End Function
Private Sub GrdColigadas_CellSelectionChange(ByVal lRow As Long, ByVal lCol As Long, ByVal bSelected As Boolean)
   Call ExibeLic(lRow)
End Sub

Private Sub ExibeLic(Optional lRow As Long)
   Dim sLocalPath As String
   Dim sArqLic As String
   Dim sTag As String
   Dim sNMCOL As String
   
   Screen.MousePointer = vbHourglass
     
   Me.TxtCripto.Text = ""
   Me.TxtDecripto.Text = ""
   Me.TxtResult.Text = ""
     
   sLocalPath = ResolvePathName(App.Path)
   If Trim(sLocalPath) = "" Then
      Screen.MousePointer = vbDefault
      Exit Sub
   End If

   sNMCOL = Me.GrdColigadas.CellValue(lRow, 2)
   sArqLic = sNMCOL & ".lic"

   If ExisteArquivo(sLocalPath & sArqLic) Then
      sTag = ReadTextFile(sLocalPath & sArqLic)
      Me.TxtCripto.Text = sTag
      'Me.TxtDecripto.Text = Decrypt2(Me.TxtCripto.Text)
      'Me.TxtResult.Text = Encrypt2(Me.TxtDecripto.Text)
   End If
   Screen.MousePointer = vbDefault
   
     
End Sub

Private Sub TxtCripto_Change()
   Dim sTag As String
   sTag = StrReverse(Decrypt2(Me.TxtCripto.Text))
   If sTag <> "" Then
      sTag = StrReverse(Mid(sTag, InStr(sTag, "|")))
   End If
   Me.TxtDecripto.Text = sTag
End Sub

Private Sub TxtDecripto_Change()
   Me.TxtResult.Text = Encrypt2(Me.TxtDecripto.Text)
End Sub
