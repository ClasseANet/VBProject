VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TL_ParamCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private WithEvents mvarMDI As FrmMDI
Attribute mvarMDI.VB_VarHelpID = -1
Private WithEvents mvarMe  As FrmParamCom
Attribute mvarMe.VB_VarHelpID = -1

Dim lStatus As String
Dim bErreMail As Boolean
 
Private Enum VALIDATE_HOST_METHOD
    VALIDATE_HOST_NONE = 0
    VALIDATE_HOST_SYNTAX = 1
    VALIDATE_HOST_PING = 2
    VALIDATE_HOST_DNS = 3
End Enum

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
     .ChkUseAuthentication.Value = 0 'xtpUnchecked
     .ChkSSL.Value = 0 'xtpUnchecked
     .ChkUsePopAuthentication.Value = 0 'xtpUnchecked
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
      .SSL = (mvarMe.ChkSSL.Value = xtpChecked)
  
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
   Set xMail = Nothing
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
   sMsg = NgObj.GetMessageSMS(eTipoSMS.PosicaoDia)
   
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
   Call MontaTela
   Call PopulaTela
'   Call mvarMe_Resize
End Sub

Private Sub mvarMe_OptMailFreqCompraClick(Index As Integer)
   If mvarMe.CmbMailFreqCompra.Text <> "" Then mvarMe.CmbMailFreqCompra.Tag = mvarMe.CmbMailFreqCompra.ListIndex
   If mvarMe.TxtMailFreqCompra_Dia.Text <> "" Then mvarMe.TxtMailFreqCompra_Dia.Tag = mvarMe.TxtMailFreqCompra_Dia.Text
   If mvarMe.TxtMailFreqCompra_Dias.Text <> "" Then mvarMe.TxtMailFreqCompra_Dias.Tag = mvarMe.TxtMailFreqCompra_Dias.Text
   
   mvarMe.CmbMailFreqCompra.ListIndex = -1
   mvarMe.TxtMailFreqCompra_Dia.Text = ""
   mvarMe.TxtMailFreqCompra_Dias.Text = ""
   mvarMe.CmbMailFreqCompra.Enabled = False
   mvarMe.TxtMailFreqCompra_Dia.Enabled = False
   mvarMe.TxtMailFreqCompra_Dias.Enabled = False
   Select Case Index
      Case 1:
         mvarMe.CmbMailFreqCompra.Enabled = True
         mvarMe.CmbMailFreqCompra.ListIndex = xVal(mvarMe.CmbMailFreqCompra.Tag)
      Case 2:
         mvarMe.TxtMailFreqCompra_Dia.Enabled = True
         mvarMe.TxtMailFreqCompra_Dia.Text = mvarMe.TxtMailFreqCompra_Dia.Tag
      Case 3:
         mvarMe.TxtMailFreqCompra_Dias.Enabled = True
         mvarMe.TxtMailFreqCompra_Dias.Text = mvarMe.TxtMailFreqCompra_Dias.Tag
   End Select
   
End Sub

Private Sub mvarMe_OptMailFreqInvClick(Index As Integer)
   If mvarMe.CmbMailFreqInv.Text <> "" Then mvarMe.CmbMailFreqInv.Tag = mvarMe.CmbMailFreqInv.ListIndex
   If mvarMe.TxtMailFreqInv_Dia.Text <> "" Then mvarMe.TxtMailFreqInv_Dia.Tag = mvarMe.TxtMailFreqInv_Dia.Text
   If mvarMe.TxtMailFreqInv_Dias.Text <> "" Then mvarMe.TxtMailFreqInv_Dias.Tag = mvarMe.TxtMailFreqInv_Dias.Text
   
   mvarMe.CmbMailFreqInv.ListIndex = -1
   mvarMe.TxtMailFreqInv_Dia.Text = ""
   mvarMe.TxtMailFreqInv_Dias.Text = ""
   mvarMe.CmbMailFreqInv.Enabled = False
   mvarMe.TxtMailFreqInv_Dia.Enabled = False
   mvarMe.TxtMailFreqInv_Dias.Enabled = False
   Select Case Index
      Case 1:
         mvarMe.CmbMailFreqInv.Enabled = True
         mvarMe.CmbMailFreqInv.ListIndex = xVal(mvarMe.CmbMailFreqInv.Tag)
      Case 2:
         mvarMe.TxtMailFreqInv_Dia.Enabled = True
         mvarMe.TxtMailFreqInv_Dia.Text = mvarMe.TxtMailFreqInv_Dia.Tag
      Case 3:
         mvarMe.TxtMailFreqInv_Dias.Enabled = True
         mvarMe.TxtMailFreqInv_Dias.Text = mvarMe.TxtMailFreqInv_Dias.Tag
   End Select
End Sub
Private Sub mvarMe_Resize()
   With mvarMe
      .GrpBoxTop.Move 0, -90, .Width - 120, 800
      
      .GrpBoxInternet.Move 120, 120, .PageInternet.Width - 240
      .GrpBoxSMS.Move 120, 120, .PageSMS.Width - 240
      .GrpBoxServicos.Move 120, 120, .PageServico.Width - 240
      '.grpboxser
      '.GrpTela.ZOrder 1
      '.GrpTela.Move 0, 0, .ScaleWidth - 60, .ScaleHeight
      '.GrpBoxser.Move 0, 0
      
      .TabComunicacao.Height = .GrpBoxInternet.Height + .GrpBoxInternet.Top + 360
      If .GrpBoxInternet.Height < .GrpBoxSMS.Height Then .TabComunicacao.Height = .GrpBoxSMS.Height + .GrpBoxSMS.Top + 360
      If .GrpBoxSMS.Height < .GrpBoxServicos.Height Then .TabComunicacao.Height = .GrpBoxServicos.Height + .GrpBoxServicos.Top + 360
      .TabComunicacao.Move 0, 0, .Width - 120
      If .TabComunicacao.Height < 8280 Then '.Height Then
         .TabComunicacao.Height = 8280 '.Height - (60 * 26)
      End If
      .GrpBoxBotton.Move -60, .GrpBoxServicos.Top + .GrpBoxServicos.Height, .GrpBoxTop.Width, 800
      .TabPgBotton.Move 0, 0, .GrpBoxBotton.Width, .GrpBoxBotton.Height
            
      .CmdPadrao.Top = 240
      .CmdCancelar.Top = .CmdPadrao.Top
      .CmdOk.Top = .CmdPadrao.Top
            
      .CmdPadrao.Left = .GrpBoxServicos.Width - .CmdPadrao.Width - 240
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
      
      .Resizer.Visible = True
      .Resizer.Top = .GrpBoxTop.Top + .GrpBoxTop.Height + 120
      .Resizer.Left = 0
      .Resizer.Width = .Width - 120
      .Resizer.Height = mvarMDI.Height - .Resizer.Top - 1040
      .Resizer.BorderStyle = xtpBorderFrame
      .Resizer.VScrollMaximum = .TabComunicacao.Top + .TabComunicacao.Height
      .Resizer.BorderStyle = xtpBorderNone
      .Resizer.VScrollPosition = 0
   End With
End Sub
Private Sub Class_Initialize()
   Set mvarMe = New FrmParamCom
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
Private Sub MontaTela()
   mvarMe.TabComunicacao.Item(0).Selected = True
   With mvarMe.CmbMailFreqInv
      .Clear
      .AddItem "Segunda-Feira"
      .AddItem "Terça-Feira"
      .AddItem "Quarta-Feira"
      .AddItem "Quinta-Feira"
      .AddItem "Sexta-Feira"
      .AddItem "Sábado"
   End With
   With mvarMe.CmbMailFreqCompra
      .Clear
      .AddItem "Segunda-Feira"
      .AddItem "Terça-Feira"
      .AddItem "Quarta-Feira"
      .AddItem "Quinta-Feira"
      .AddItem "Sexta-Feira"
      .AddItem "Sábado"
   End With
   mvarMe.OptMailFreqCompra(0).Visible = True
   mvarMe.CmbMailFreqCompra.Enabled = False
   mvarMe.TxtMailFreqCompra_Dia.Enabled = False
   mvarMe.TxtMailFreqCompra_Dias.Enabled = False
   mvarMe.CmbMailFreqInv.Enabled = False
   mvarMe.TxtMailFreqInv_Dia.Enabled = False
   mvarMe.TxtMailFreqInv_Dias.Enabled = False
End Sub
Private Sub PopulaTela()
   Dim sAux As String
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
      .ChkSSL.Value = xVal(Sys.Getparam("SSL", Default:=0))
      .ChkUsePopAuthentication.Value = xVal(Sys.Getparam("UsePopAuthentication", Default:=1))
      
      .TxtFTP.Text = Decrypt2(Sys.Getparam("FTP"))
      .TxtFTPBakPath.Text = Sys.Getparam("FTPBakPath")
      .TxtFtpUID.Text = Sys.Getparam("FtpUID")
      .TxtFtpPWD.Text = Decrypt2(Sys.Getparam("FtpPWD"))
      .TxtSMSURLBASE.Text = Decrypt2(Sys.Getparam("SMSURLBASE"))
      .TxtSMSUID.Text = Decrypt2(Sys.Getparam("SMSUID"))
      .TxtSMSPWD.Text = Decrypt2(Sys.Getparam("SMSPWD"))
     
      .TxtLstEMailArqFin.Text = Sys.Getparam("LstEMailArqFin")
      .TxtLstEMailOcorrencia.Text = Sys.Getparam("LstEMailOcorrencia")
      .TxtLstEMailSocio.Text = Sys.Getparam("LstEMailSocio")
      .TxtLstEMailSupri.Text = Sys.Getparam("LstEMailSupri")
      .TxtLstCelSocio.Text = Sys.Getparam("LstCelSocio")
      .TxtDPILSuporte.Text = Sys.Getparam("DpilSuporte")
      
      sAux = Sys.Getparam("MailFreqCompra", Default:="D")
      Select Case Mid(sAux, 1, 1)
         Case "D"
            mvarMe.OptMailFreqCompra(0).Value = True
         Case "S"
            mvarMe.OptMailFreqCompra(1).Value = True
            mvarMe.CmbMailFreqCompra.ListIndex = xVal(Mid(sAux, 2))
         Case "M"
            mvarMe.OptMailFreqCompra(2).Value = True
            mvarMe.TxtMailFreqCompra_Dia.Text = xVal(Mid(sAux, 2))
         Case "O"
            mvarMe.OptMailFreqCompra(3).Value = True
            mvarMe.TxtMailFreqCompra_Dias.Text = Mid(sAux, 2)
      End Select
      
      sAux = Sys.Getparam("MailFreqInv", Default:="D")
      Select Case Mid(sAux, 1, 1)
         Case "D"
            mvarMe.OptMailFreqInv(0).Value = True
         Case "S"
            mvarMe.OptMailFreqInv(1).Value = True
            mvarMe.CmbMailFreqInv.ListIndex = xVal(Mid(sAux, 2))
         Case "M"
            mvarMe.OptMailFreqInv(2).Value = True
            mvarMe.TxtMailFreqInv_Dia.Text = xVal(Mid(sAux, 2))
         Case "O"
            mvarMe.OptMailFreqInv(3).Value = True
            mvarMe.TxtMailFreqInv_Dia.Text = Mid(sAux, 2)
      End Select
  End With
End Sub
Private Sub SalvarTela()
   Dim TbParam As Object
   Dim Queries As Collection
   
   Set Queries = New Collection
   
   With mvarMe
      
      Call QrySaveParam(mvarsys, "POP3Host", .TxtPOP3Host.Text, Queries)
      Call QrySaveParam(mvarsys, "SMTPHost", .TxtSMTPHost.Text, Queries)
      Call QrySaveParam(mvarsys, "SMTPPort", .TxtSMTPPort.Text, Queries)
      Call QrySaveParam(mvarsys, "FromDisplayName", .TxtFromDisplayName.Text, Queries)
      Call QrySaveParam(mvarsys, "MailUID", .TxtMailUID.Text, Queries)
      Call QrySaveParam(mvarsys, "MailPWD", Encrypt2(.TxtMailPWD.Text), Queries)
      Call QrySaveParam(mvarsys, "UseAuthentication", .ChkUseAuthentication.Value, Queries)
      Call QrySaveParam(mvarsys, "SSL", .ChkSSL.Value, Queries)
      Call QrySaveParam(mvarsys, "UsePopAuthentication", .ChkUsePopAuthentication.Value, Queries)
      Call QrySaveParam(mvarsys, "FTP", Encrypt2(.TxtFTP.Text), Queries)
      Call QrySaveParam(mvarsys, "FTPBakPath", .TxtFTPBakPath.Text, Queries)
      Call QrySaveParam(mvarsys, "FtpUID", .TxtFtpUID.Text, Queries)
      Call QrySaveParam(mvarsys, "FtpPWD", Encrypt2(.TxtFtpPWD.Text), Queries)
      Call QrySaveParam(mvarsys, "SMSURLBASE", Encrypt2(.TxtSMSURLBASE.Text), Queries)
      Call QrySaveParam(mvarsys, "SMSUID", Encrypt2(.TxtSMSUID.Text), Queries)
      Call QrySaveParam(mvarsys, "SMSPWD", Encrypt2(.TxtSMSPWD.Text), Queries)
      Call QrySaveParam(mvarsys, "SOCIOCOM", .ChkSOCIOCOM.Value, Queries)
      Call QrySaveParam(mvarsys, "FRANQCOM", .ChkFRANQCOM.Value, Queries)
      Call QrySaveParam(mvarsys, "LstEMailArqFin", .TxtLstEMailArqFin.Text, Queries)
      Call QrySaveParam(mvarsys, "LstEMailOcorrencia", .TxtLstEMailOcorrencia.Text, Queries)
      Call QrySaveParam(mvarsys, "LstEMailSocio", .TxtLstEMailSocio.Text, Queries)
      Call QrySaveParam(mvarsys, "LstEMailSupri", .TxtLstEMailSupri.Text, Queries)
      Call QrySaveParam(mvarsys, "LstCelSocio", .TxtLstCelSocio.Text, Queries)
      Call QrySaveParam(mvarsys, "DpilSuporte", .TxtDPILSuporte.Text, Queries)
      If mvarMe.OptMailFreqCompra(0).Value Then
         Call QrySaveParam(mvarsys, "MailFreqCompra", "D", Queries)
      ElseIf mvarMe.OptMailFreqCompra(1).Value Then
         Call QrySaveParam(mvarsys, "MailFreqCompra", "S" & mvarMe.CmbMailFreqCompra.ListIndex, Queries)
      ElseIf mvarMe.OptMailFreqCompra(2).Value Then
         Call QrySaveParam(mvarsys, "MailFreqCompra", "M" & mvarMe.TxtMailFreqCompra_Dia.Text, Queries)
      ElseIf mvarMe.OptMailFreqCompra(3).Value Then
         Call QrySaveParam(mvarsys, "MailFreqCompra", "O" & mvarMe.TxtMailFreqCompra_Dias.Text, Queries)
      End If
      If mvarMe.OptMailFreqInv(0).Value Then
         Call QrySaveParam(mvarsys, "MailFreqInv", "D", Queries)
      ElseIf mvarMe.OptMailFreqInv(1).Value Then
         Call QrySaveParam(mvarsys, "MailFreqInv", "S" & mvarMe.CmbMailFreqInv.ListIndex, Queries)
      ElseIf mvarMe.OptMailFreqInv(2).Value Then
         Call QrySaveParam(mvarsys, "MailFreqInv", "M" & mvarMe.TxtMailFreqInv_Dia.Text, Queries)
      ElseIf mvarMe.OptMailFreqInv(3).Value Then
         Call QrySaveParam(mvarsys, "MailFreqInv", "O" & mvarMe.TxtMailFreqInv_Dias.Text, Queries)
      End If
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


