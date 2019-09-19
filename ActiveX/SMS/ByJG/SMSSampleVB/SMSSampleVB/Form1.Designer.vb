<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
		Me.label2 = New System.Windows.Forms.Label
		Me.label3 = New System.Windows.Forms.Label
		Me.label1 = New System.Windows.Forms.Label
		Me.button1 = New System.Windows.Forms.Button
		Me.label4 = New System.Windows.Forms.Label
		Me.groupBox1 = New System.Windows.Forms.GroupBox
		Me.txtSenha = New System.Windows.Forms.TextBox
		Me.txtUsuario = New System.Windows.Forms.TextBox
		Me.txtDDD = New System.Windows.Forms.TextBox
		Me.txtMensagem = New System.Windows.Forms.TextBox
		Me.txtCelular = New System.Windows.Forms.TextBox
		Me.groupBox1.SuspendLayout()
		Me.SuspendLayout()
		'
		'label2
		'
		Me.label2.AutoSize = True
		Me.label2.Location = New System.Drawing.Point(15, 40)
		Me.label2.Name = "label2"
		Me.label2.Size = New System.Drawing.Size(59, 13)
		Me.label2.TabIndex = 12
		Me.label2.Text = "Mensagem"
		'
		'label3
		'
		Me.label3.AutoSize = True
		Me.label3.Location = New System.Drawing.Point(13, 19)
		Me.label3.Name = "label3"
		Me.label3.Size = New System.Drawing.Size(43, 13)
		Me.label3.TabIndex = 2
		Me.label3.Text = "Usuário"
		'
		'label1
		'
		Me.label1.AutoSize = True
		Me.label1.Location = New System.Drawing.Point(15, 14)
		Me.label1.Name = "label1"
		Me.label1.Size = New System.Drawing.Size(39, 13)
		Me.label1.TabIndex = 11
		Me.label1.Text = "Celular"
		'
		'button1
		'
		Me.button1.Location = New System.Drawing.Point(396, 192)
		Me.button1.Name = "button1"
		Me.button1.Size = New System.Drawing.Size(75, 23)
		Me.button1.TabIndex = 13
		Me.button1.Text = "Enviar"
		Me.button1.UseVisualStyleBackColor = True
		'
		'label4
		'
		Me.label4.AutoSize = True
		Me.label4.Location = New System.Drawing.Point(197, 22)
		Me.label4.Name = "label4"
		Me.label4.Size = New System.Drawing.Size(38, 13)
		Me.label4.TabIndex = 3
		Me.label4.Text = "Senha"
		'
		'groupBox1
		'
		Me.groupBox1.Controls.Add(Me.label4)
		Me.groupBox1.Controls.Add(Me.label3)
		Me.groupBox1.Controls.Add(Me.txtSenha)
		Me.groupBox1.Controls.Add(Me.txtUsuario)
		Me.groupBox1.Location = New System.Drawing.Point(18, 133)
		Me.groupBox1.Name = "groupBox1"
		Me.groupBox1.Size = New System.Drawing.Size(453, 53)
		Me.groupBox1.TabIndex = 10
		Me.groupBox1.TabStop = False
		Me.groupBox1.Text = "Autenticacao"
		'
		'txtSenha
		'
		Me.txtSenha.Location = New System.Drawing.Point(241, 19)
		Me.txtSenha.Name = "txtSenha"
		Me.txtSenha.Size = New System.Drawing.Size(100, 20)
		Me.txtSenha.TabIndex = 5
		'
		'txtUsuario
		'
		Me.txtUsuario.Location = New System.Drawing.Point(62, 19)
		Me.txtUsuario.Name = "txtUsuario"
		Me.txtUsuario.Size = New System.Drawing.Size(100, 20)
		Me.txtUsuario.TabIndex = 4
		'
		'txtDDD
		'
		Me.txtDDD.Location = New System.Drawing.Point(80, 11)
		Me.txtDDD.MaxLength = 2
		Me.txtDDD.Name = "txtDDD"
		Me.txtDDD.Size = New System.Drawing.Size(35, 20)
		Me.txtDDD.TabIndex = 7
		'
		'txtMensagem
		'
		Me.txtMensagem.Location = New System.Drawing.Point(80, 40)
		Me.txtMensagem.MaxLength = 160
		Me.txtMensagem.Multiline = True
		Me.txtMensagem.Name = "txtMensagem"
		Me.txtMensagem.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
		Me.txtMensagem.Size = New System.Drawing.Size(391, 87)
		Me.txtMensagem.TabIndex = 9
		'
		'txtCelular
		'
		Me.txtCelular.Location = New System.Drawing.Point(121, 11)
		Me.txtCelular.MaxLength = 8
		Me.txtCelular.Name = "txtCelular"
		Me.txtCelular.Size = New System.Drawing.Size(81, 20)
		Me.txtCelular.TabIndex = 8
		'
		'Form1
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(486, 223)
		Me.Controls.Add(Me.label2)
		Me.Controls.Add(Me.label1)
		Me.Controls.Add(Me.button1)
		Me.Controls.Add(Me.groupBox1)
		Me.Controls.Add(Me.txtDDD)
		Me.Controls.Add(Me.txtMensagem)
		Me.Controls.Add(Me.txtCelular)
		Me.Name = "Form1"
		Me.Text = "Form1"
		Me.groupBox1.ResumeLayout(False)
		Me.groupBox1.PerformLayout()
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub
	Private WithEvents label2 As System.Windows.Forms.Label
	Private WithEvents label3 As System.Windows.Forms.Label
	Private WithEvents label1 As System.Windows.Forms.Label
	Private WithEvents button1 As System.Windows.Forms.Button
	Private WithEvents label4 As System.Windows.Forms.Label
	Private WithEvents groupBox1 As System.Windows.Forms.GroupBox
	Private WithEvents txtSenha As System.Windows.Forms.TextBox
	Private WithEvents txtUsuario As System.Windows.Forms.TextBox
	Private WithEvents txtDDD As System.Windows.Forms.TextBox
	Private WithEvents txtMensagem As System.Windows.Forms.TextBox
	Private WithEvents txtCelular As System.Windows.Forms.TextBox
End Class
