Attribute VB_Name = "Module1"
Private Sub FuncaoTeste()
   Dim xMail As CAMail.SendMail
   Dim bOk As Boolean
   Dim sHtml As String
  
'mvarMe.WebBrowser1.Navigate "C:\teste.htm"
  
Exit Sub
  
   Set xMail = New CAMail.SendMail
   With xMail
      .UseAuthentication = True
      .UsePopAuthentication = True
  
      .POP3Host = "pop3.bol.com.br"
      .SMTPHost = "smtps.bol.com.br"
      .Username = "diogenes72@bol.com.br"
      .Password = "dolphin7"
      
      .From = .Username
      .FromDisplayName = "Diogenes"
      
      .AsHTML = True
      .Subject = "Dpil Freguesia"
      
      sHtml = ""
      sHtml = "<!DOCTYPE HTML PUBLIC " & """" & "-//W3C//DTD HTML 4.01 Transitional//EN" & """" & ">"
      sHtml = sHtml & "<html>"
      sHtml = sHtml & "<head>"
      sHtml = sHtml & "<title>Untitled Document</title>"
      sHtml = sHtml & "<meta http-equiv=" & """" & "Content-Type" & """" & " content=" & """" & "text/html; charset=iso-8859-1" & """" & ">"
      sHtml = sHtml & "</head>"
      sHtml = sHtml & ""
      sHtml = sHtml & "<body>"
      sHtml = sHtml & "<table width=" & """" & "75%" & """" & " border=" & """" & "1" & """" & ">"
      sHtml = sHtml & "  <tr>"
      sHtml = sHtml & "    <td>Teste1</td>"
      sHtml = sHtml & "    <td>Teste2</td>"
      sHtml = sHtml & "    <td>&nbsp;</td>"
      sHtml = sHtml & "  </tr>"
      sHtml = sHtml & "  <tr>"
      sHtml = sHtml & "    <td>&nbsp;</td>"
      sHtml = sHtml & "    <td>&nbsp;</td>"
      sHtml = sHtml & "    <td>&nbsp;</td>"
      sHtml = sHtml & "  </tr>"
      sHtml = sHtml & "  <tr>"
      sHtml = sHtml & "    <td>&nbsp;</td>"
      sHtml = sHtml & "    <td>&nbsp;</td>"
      sHtml = sHtml & "    <td>&nbsp;</td>"
      sHtml = sHtml & "  </tr>"
      sHtml = sHtml & "</table>"
      sHtml = sHtml & "</body>"
      sHtml = sHtml & "</html>"

      .Message = sHtml

      .Receipt = True
      .Recipient = "disantos@ig.com.br"
      .RecipientDisplayName = "DiSantos"
      .SMTPHostValidation = VALIDATE_HOST_NONE
      .SMTPPort = 587


      .Connect
      .Send
     .Disconnect
  End With
End Sub
