Imports SMSSampleVB.br.com.byjg.www

Public Class Form1

	Private Sub button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button1.Click
		Dim sms As New SMSService()

		Dim result As String = sms.enviarSMS(txtDDD.Text, txtCelular.Text, txtMensagem.Text, txtUsuario.Text, txtSenha.Text)

		MessageBox.Show(result)
	End Sub
End Class