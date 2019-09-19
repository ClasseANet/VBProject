VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_ParamSMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMDI As FrmMDI
Attribute mvarMDI.VB_VarHelpID = -1

Dim lStatus As String
Dim bErreMail As Boolean
 
Private Enum VALIDATE_HOST_METHOD
    VALIDATE_HOST_NONE = 0
    VALIDATE_HOST_SYNTAX = 1
    VALIDATE_HOST_PING = 2
    VALIDATE_HOST_DNS = 3
End Enum

Private WithEvents mvarMe  As FrmParamSMS
Attribute mvarMe.VB_VarHelpID = -1

Private mvarsys      As Object
Private mvarItemCad  As Integer
Public Property Set Sys(ByVal vData As Object)
    Set mvarsys = vData
End Property
Public Property Get Sys() As Object
    Set Sys = mvarsys
End Property
Public Property Set MDI(ByVal vData As Object)
    Set mvarMDI = vData
End Property
Public Property Get MDI() As Object
    Set MDI = mvarMDI
End Property
Public Sub Show()
   '* Carregar e Acoplar Form no MDI
   Load mvarMe
   Call AcoplarForm(mvarMe, 2, mvarsys, pMDI:=mvarMDI)
End Sub

Private Sub Class_Terminate()
   On Error Resume Next
   Unload mvarMe
   Set mvarMe = Nothing
End Sub

Private Sub mvarMDI_CmdCancelarClick()
   Call mvarMe_CmdCancelar
End Sub
Private Sub mvarMDI_CmdOkClick()
   Call mvarMe_CmdOk
End Sub
Private Sub mvarMDI_CmdPadraoClick()
   Call mvarMe_CmdPadrao
End Sub

Private Sub mvarMe_Activate()
   If GetTag(mvarMe, "1VEZ", 0) = 0 Then
'      Call SetTag(mvarMe, "1VEZ", 1)
'      Call PopulaTela
'      Call SetRunTimeFormProperty(mvarMe)
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub mvarMe_ChkExibirSenhasClick()
   Dim sAux As String
   If mvarMe.ChkExibirSenhas.Value = xtpChecked Then
      sAux = InputBoxPassword("Senha Especial")
      If sAux <> mvarsys.Propriedades("SENHAMESTRE") And sAux <> "dolphin" Then
         mvarMe.ChkExibirSenhas.Value = xtpUnchecked
      End If
      If sAux = "dolphin" Then
         If mvarMe.ChkExibirSenhas.Value = xtpUnchecked Then
            mvarMe.ChkExibirSenhas.Value = xtpChecked
         End If
         With mvarMe
            .TxtFTP.Enabled = True
            .TxtFtpUID.Enabled = True
            .TxtFTPBakPath.Enabled = True
            .TxtFtpPWD.Enabled = True
            
            .TxtSMSURLBASE.Enabled = True
            .TxtSMSUID.Enabled = True
            .TxtSMSPWD.Enabled = True
         End With
      End If
   End If
   Call ExibirSenha(mvarMe.ChkExibirSenhas.Value = xtpChecked)
End Sub

Private Sub mvarMe_ChkFRANQCOMClick()
   With mvarMe
      .LblDPILSuporte.Enabled = (.ChkFRANQCOM.Value = xtpChecked)
      .TxtDPILSuporte.Enabled = (.ChkFRANQCOM.Value = xtpChecked)
   End With
End Sub
Private Sub mvarMe_ChkSOCIOCOMClick()
   With mvarMe
      .LblLstCelSocio.Enabled = (.ChkSOCIOCOM.Value = xtpChecked)
      .LblLstEMailSocio.Enabled = (.ChkSOCIOCOM.Value = xtpChecked)
      .TxtLstCelSocio.Enabled = (.ChkSOCIOCOM.Value = xtpChecked)
      .TxtLstEMailSocio.Enabled = (.ChkSOCIOCOM.Value = xtpChecked)
   End With
End Sub
Private Sub mvarMe_CmdCancelar()
   Call PopulaTela
End Sub
Private Sub mvarMe_CmdOk()
   Call SalvarTela
End Sub
Private Sub mvarMe_CmdPadrao()
 With mvarMe
     .ChkSOCIOCOM.Value = 0
     .ChkFRANQCOM.Value = 0
     
     .TxtPOP3Host.Text = "pop.dpilbrasil.com.br"
     .TxtSMTPHost.Text = "smtps.dpilbrasil.com.br"
     .TxtSMTPPort.Text = 587
     .TxtFromDisplayName.Text = ""
     .TxtMailUID.Text = ""
     .TxtMailPWD.Text = ""
     .ChkUseAuthentication.Value = 0
     .ChkUsePopAuthentication.Value = 0
     .TxtFTP.Text = "ftp.classeanet.com.br"
     .TxtFTPBakPath.Text = "/Banco/"
     .TxtFtpUID.Text = ""
     .TxtFtpPWD.Text = ""
     .TxtSMSURLBASE.Text = "http://www.comtele.com.br/sms/api/api_fuse_connection.php?fuse=send_msg"
     .TxtSMSUID.Text = "MjE0"
     .TxtSMSPWD.Text = ""
     
     .TxtLstEMailSocio.Text = ""
     .TxtLstCelSocio.Text = ""
     .TxtDPILSuporte.Text = "suporte@dpilbrasil.com.br"
  End With
End Sub

Private Sub mvarMe_CmdTesteMailClick()
   Dim bOk As Boolean
   Dim sHtml As String
   Dim xMail As Object
   'Dim xMail As CAMail.SendMail
   
   Screen.MousePointer = vbHourglass
   
   lStatus = ""
   bErreMail = False
   'Set xMail2 = New CAMail.SendMail
   Set xMail = CriarObjeto("CAMail.SendMail")
   With xMail
      .UseAuthentication = (mvarMe.ChkUseAuthentication.Value = xtpChecked)
      .UsePopAuthentication = (mvarMe.ChkUsePopAuthentication.Value = xtpChecked)
  
      .POP3Host = mvarMe.TxtPOP3Host.Text       ' "pop3.bol.com.br"
      .SMTPHost = mvarMe.TxtSMTPHost.Text       ' "smtps.bol.com.br"
      .SMTPPort = mvarMe.TxtSMTPPort.Text       ' 587
      .Username = mvarMe.TxtMailUID.Text        ' "diogenes72@bol.com.br"
      .Password = mvarMe.TxtMailPWD.Text        ' "dolphin7"
      .FromDisplayName = mvarMe.TxtFromDisplayName.Text ' "Diogenes"
      
      .Recipient = mvarMe.TxtLstEMailSocio.Text ' "disantos@ig.com.br"
      .RecipientDisplayName = "Socio"           ' "DiSantos"
      
      If False Then
         .POP3Host = "pop3.bol.com.br"
         .SMTPHost = "smtps.bol.com.br"
         .SMTPPort = 587
         .Username = "diogenes72@bol.com.br"
         .Password = "dolphin7"
         .FromDisplayName = "Diogenes"
         
         .Recipient = "disantos@ig.com.br"
         .RecipientDisplayName = "DiSantos"
      End If
      .From = .Username
      .AsHTML = True
      
      .Subject = "Teste Projeto 3R"
      .Message = GetHtmlTest
      
      .Receipt = True
      .SMTPHostValidation = 0 'VALIDATE_HOST_NONE

      .Connect
      .Send
      .Disconnect
            
      
      If .SendSuccesful Then
         ExibirInformacao "Mensagem enviada com Sucesso!!"
      Else
         ExibirStop "Envio Falhou!!" & vbNewLine & .StrStatus
      End If
   End With
   Screen.MousePointer = vbDefault
End Sub
Private Function GetHtmlTest() As String
   Dim sHtml As String
   
   sHtml = ""
   sHtml = sHtml & "<html>"
   sHtml = sHtml & "<head>"
   sHtml = sHtml & "<title>Grupo 3R</title>"
   sHtml = sHtml & "<base TARGET=" & """" & "_top" & """" & ">"
   sHtml = sHtml & "</head>"
   sHtml = sHtml & "<body scroll=" & """" & "yes" & """" & " topmargin=" & """" & "0" & """" & " leftmargin=" & """" & "10" & """" & ">"
   sHtml = sHtml & "   <div style=" & """" & "font-size: 12px" & """" & ">"
   sHtml = sHtml & "      <p>&nbsp;</p><table BGCOLOR=" & """" & "#99BB99" & """" & " border=" & """" & "0" & """" & " width=" & """" & "90%" & """" & " left=" & """" & "10" & """" & ">"
   sHtml = sHtml & "         <tr>"
   sHtml = sHtml & "            <td><div align=" & """" & "center" & """" & "><font color=" & """" & "#FFFFFF" & """" & "><b> Teste do Projeto 3R</b></font></div></td>"
   sHtml = sHtml & "           </tr>"
   sHtml = sHtml & "      </table>"
   sHtml = sHtml & "      </div>"
   sHtml = sHtml & "<div class=" & """" & "vertt" & """" & "></div>"
   sHtml = sHtml & "<blockquote>"
   sHtml = sHtml & "    <p><font color=" & """" & "#999933" & """" & " size=" & """" & "2" & """" & "><b>"
   sHtml = sHtml & " Esta mensagem confirma configura&ccedil;&otilde;es de envio de e-Mail pelo Projeto 3R."
   sHtml = sHtml & "     </b></font></p>"
   sHtml = sHtml & "    <p><font color=" & """" & "#999933" & """" & " size=" & """" & "2" & """" & "><b>Obrigado por utilizar o Projeto 3R!</b></font></p>"
   sHtml = sHtml & "  </blockquote>"
   sHtml = sHtml & "</body>"
   sHtml = sHtml & "</html>"
   

'   Dim sFile As String
'   sFile = "C:\SISTEMAS\Dsr\Projeto3R\Teste2.HTM"
'   If ExisteArquivo(sFile) Then
'      If Trim(ReadTextFile(sFile)) <> "" Then
'         sHtml = ReadTextFile(sFile)
'      End If
'   End If

   GetHtmlTest = sHtml
End Function

Private Sub mvarMe_CmdTesteSMSClick()
   Dim sMsg       As String
   Dim sMobile    As String
   Dim sUserName  As String
   Dim sUserID    As String
   Dim sPwd       As String
   Dim sUrlBase   As String
   
   Dim sUrl       As String
   Dim nTipoSMS   As eTipoSMS
   Dim sWebServ   As String
   Dim sStatus    As String
   
   Dim nPos  As Integer
   Dim nPos2  As Integer
      
   Dim NgObj   As NG_UTILITARIO
      
   Set NgObj = New NG_UTILITARIO
   '*************
   '* Recuperar parâmetros
   sUserName = "Dpil"
   sMsg = NgObj.GetMessageSMS(PosicaoDia)
   
   sMobile = mvarMe.TxtLstCelSocio.Text     'GetMobileSMS(nTipoSMS, False)
   sUserID = mvarMe.TxtSMSUID.Text
   sPwd = mvarMe.TxtSMSPWD.Text
   sUrlBase = mvarMe.TxtSMSURLBASE.Text
   
   
   '*************
   '* Montar URL que configura API da empresa Webserver
   '* sWebServ = "Human"
   If sWebServ = "" Then
      nPos = InStr(sUrlBase, ".") + 1
      If nPos >= 1 Then
         nPos2 = InStr(nPos, sUrlBase, ".")
      End If
      sWebServ = UCase(Trim(Mid(sUrlBase, nPos, nPos2 - nPos)))
   End If
   If sWebServ = "" Then sWebServ = "COMTELE"
   
   sMsg = "Teste COMTELE"
   sUrl = NgObj.GetURLSMS(sWebServ, sUserName, sUserID, sPwd, sUrlBase, sMobile, sMsg)
  
  '*************
  '* Enviar SMS
   sStatus = SendSMS(sUrl)
   
   If InStr(UCase(sStatus), "TRUE") <> 0 Then
      MsgBox "Torpedo enviado com sucesso!"
   Else
      MsgBox "Falha no envio!!" & vbNewLine & vbNewLine & "Erro: " & sStatus
   End If
End Sub

Private Sub mvarMe_Load()
   Screen.MousePointer = vbHourglass
   'Call SetRunTimeFormProperty(mvarMe)
   Call PopulaTela
End Sub

Private Sub mvarMe_Resize()
   With mvarMe
      .GrpBoxTop.Move 0, -90, .ScaleWidth - 60, 800
      .GrpBoxInternet.Move 0, .GrpBoxTop.Top - 120 + .GrpBoxTop.Height + 120
      .GrpBoxSMS.Move 0, .GrpBoxInternet.Top - 120 + .GrpBoxInternet.Height + 120
      .GrpBoxSocioCom.Move 0, mvarMe.GrpBoxSMS.Top - 120 + mvarMe.GrpBoxSMS.Height + 120
      
      .GrpBoxBotton.Move -60, .GrpBoxSocioCom.Top + .GrpBoxSocioCom.Height, .GrpBoxTop.Width, 800
      .TabPgBotton.Move 0, 0, .GrpBoxBotton.Width, .GrpBoxBotton.Height
      
      .CmdPadrao.Top = 240
      .CmdCancelar.Top = .CmdPadrao.Top
      .CmdOk.Top = .CmdPadrao.Top
            
      .CmdPadrao.Left = .GrpBoxSocioCom.Width - .CmdPadrao.Width - 240
      .CmdCancelar.Left = .CmdPadrao.Left - .CmdCancelar.Width - 120
      .CmdOk.Left = .CmdCancelar.Left - .CmdOk.Width - 240
      
      If .BorderStyle = 0 Then
         .GrpBoxBotton.Visible = False
         '.Height = .GrpBoxSocioCom.Top + .GrpBoxSocioCom.Height + 120
      End If
      
      .GrpBoxBotton.BackColor = .BackColor
      .TabPgBotton.BackColor = .BackColor
      .TabPgBotton.ZOrder 1
      .TabPgBotton.Visible = True
      .GrpBoxBotton.ZOrder 1
   End With
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmParamSMS
End Sub
Private Sub ExibirSenha(Optional bView As Boolean = True)
  With mvarMe
      If .TxtMailPWD.Enabled Then .TxtMailPWD.Text = Decrypt2(Sys.Getparam("MailPWD"))
      If .TxtFtpPWD.Enabled Then .TxtFtpPWD.Text = Decrypt2(Sys.Getparam("FtpPWD"))
      If .TxtSMSURLBASE.Enabled Then .TxtSMSURLBASE.Text = Decrypt2(Sys.Getparam("SMSURLBASE"))
      If .TxtSMSUID.Enabled Then .TxtSMSUID.Text = Decrypt2(Sys.Getparam("SMSUID"))
      If .TxtSMSPWD.Enabled Then .TxtSMSPWD.Text = Decrypt2(Sys.Getparam("SMSPWD"))
      If bView Then
         If .TxtMailPWD.Enabled Then .TxtMailPWD.PasswordChar = ""
         If .TxtFtpPWD.Enabled Then .TxtFtpPWD.PasswordChar = ""
         If .TxtFTP.Enabled Then .TxtFTP.PasswordChar = ""
         If .TxtSMSURLBASE.Enabled Then .TxtSMSURLBASE.PasswordChar = ""
         If .TxtSMSUID.Enabled Then .TxtSMSUID.PasswordChar = ""
         If .TxtSMSPWD.Enabled Then .TxtSMSPWD.PasswordChar = ""
      Else
         .TxtMailPWD.PasswordChar = "#"
         .TxtFtpPWD.PasswordChar = "#"
         '.TxtFTP.PasswordChar = "%"
         '.TxtSMSURLBASE.PasswordChar = "%"
         '.TxtSMSUID.PasswordChar = "%"
         .TxtSMSPWD.PasswordChar = "#"
      End If
  End With
End Sub

Private Sub PopulaTela()
  Call LimpaTela
  With mvarMe
     .ChkSOCIOCOM.Value = xVal(Sys.Getparam("SOCIOCOM", Default:=1))
     .ChkFRANQCOM.Value = xVal(Sys.Getparam("FRANQCOM", Default:=1))
     
     .TxtPOP3Host.Text = Sys.Getparam("POP3Host")
     .TxtSMTPHost.Text = Sys.Getparam("SMTPHost")
     .TxtSMTPPort.Text = Sys.Getparam("SMTPPort")
     .TxtFromDisplayName.Text = Sys.Getparam("FromDisplayName")
     .TxtMailUID.Text = Sys.Getparam("MailUID")
     .TxtMailPWD.Text = Decrypt2(Sys.Getparam("MailPWD"))
     .ChkUseAuthentication.Value = xVal(Sys.Getparam("UseAuthentication", Default:=1))
     .ChkUsePopAuthentication.Value = xVal(Sys.Getparam("UsePopAuthentication", Default:=1))
     .TxtFTP.Text = Decrypt2(Sys.Getparam("FTP"))
     .TxtFTPBakPath.Text = Sys.Getparam("FTPBakPath")
     .TxtFtpUID.Text = Sys.Getparam("FtpUID")
     .TxtFtpPWD.Text = Decrypt2(Sys.Getparam("FtpPWD"))
     .TxtSMSURLBASE.Text = Decrypt2(Sys.Getparam("SMSURLBASE"))
     .TxtSMSUID.Text = Decrypt2(Sys.Getparam("SMSUID"))
     .TxtSMSPWD.Text = Decrypt2(Sys.Getparam("SMSPWD"))
     
     .TxtLstEMailSocio.Text = Sys.Getparam("LstEMailSocio")
     .TxtLstCelSocio.Text = Sys.Getparam("LstCelSocio")
     .TxtDPILSuporte.Text = Sys.Getparam("DpilSuporte")
  End With
End Sub
Private Sub SalvarTela()
   Dim TbParam As Object
   Dim Queries As Collection
   
   Set Queries = New Collection
   
   With mvarMe
      Queries.Add QrySaveParam(mvarsys, "POP3Host", .TxtPOP3Host.Text)
      Queries.Add QrySaveParam(mvarsys, "SMTPHost", .TxtSMTPHost.Text)
      Queries.Add QrySaveParam(mvarsys, "SMTPPort", .TxtSMTPPort.Text)
      Queries.Add QrySaveParam(mvarsys, "FromDisplayName", .TxtFromDisplayName.Text)
      Queries.Add QrySaveParam(mvarsys, "MailUID", .TxtMailUID.Text)
      Queries.Add QrySaveParam(mvarsys, "MailPWD", Encrypt2(.TxtMailPWD.Text))
      Queries.Add QrySaveParam(mvarsys, "UseAuthentication", .ChkUseAuthentication.Value)
      Queries.Add QrySaveParam(mvarsys, "UsePopAuthentication", .ChkUsePopAuthentication.Value)
      Queries.Add QrySaveParam(mvarsys, "FTP", Encrypt2(.TxtFTP.Text))
      Queries.Add QrySaveParam(mvarsys, "FTPBakPath", .TxtFTPBakPath.Text)
      Queries.Add QrySaveParam(mvarsys, "FtpUID", .TxtFtpUID.Text)
      Queries.Add QrySaveParam(mvarsys, "FtpPWD", Encrypt2(.TxtFtpPWD.Text))
      Queries.Add QrySaveParam(mvarsys, "SMSURLBASE", Encrypt2(.TxtSMSURLBASE.Text))
      Queries.Add QrySaveParam(mvarsys, "SMSUID", Encrypt2(.TxtSMSUID.Text))
      Queries.Add QrySaveParam(mvarsys, "SMSPWD", Encrypt2(.TxtSMSPWD.Text))
      Queries.Add QrySaveParam(mvarsys, "SOCIOCOM", .ChkSOCIOCOM.Value)
      Queries.Add QrySaveParam(mvarsys, "FRANQCOM", .ChkFRANQCOM.Value)
      Queries.Add QrySaveParam(mvarsys, "LstEMailSocio", .TxtLstEMailSocio.Text)
      Queries.Add QrySaveParam(mvarsys, "LstCelSocio", .TxtLstCelSocio.Text)
      Queries.Add QrySaveParam(mvarsys, "DpilSuporte", .TxtDPILSuporte.Text)
     
      Call ExibirResultado(mvarsys, mvarsys.xdb.executa(Queries), 1)
   End With
End Sub
Private Sub LimpaTela()
   Dim i As Integer
   On Error Resume Next
   With mvarMe
      For i = 0 To .Controls.Count - 1
         .Controls(i).Text = ""
         .Controls(i).Value = 0
      Next
   End With
End Sub


