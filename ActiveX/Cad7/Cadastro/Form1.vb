Public Class Form1
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Protected Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents LnkCadastro As System.Windows.Forms.LinkLabel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.LnkCadastro = New System.Windows.Forms.LinkLabel
        Me.SuspendLayout()
        '
        'LnkCadastro
        '
        Me.LnkCadastro.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.LnkCadastro.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.LnkCadastro.Location = New System.Drawing.Point(24, 40)
        Me.LnkCadastro.Name = "LnkCadastro"
        Me.LnkCadastro.TabIndex = 0
        Me.LnkCadastro.TabStop = True
        Me.LnkCadastro.Text = "Cadastro"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 273)
        Me.Controls.Add(Me.LnkCadastro)
        Me.Menu = Me.MainMenu1
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub

#End Region

    
    Private Sub LnkCadastro_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LnkCadastro.LinkClicked
        MsgBox("Oi")
    End Sub
End Class
