<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Login
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Login))
        Me.Panel6 = New System.Windows.Forms.Panel()
        Me.Panel5 = New System.Windows.Forms.Panel()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.sla = New System.Windows.Forms.Label()
        Me.txt_senha_login = New System.Windows.Forms.TextBox()
        Me.txt_email_login = New System.Windows.Forms.TextBox()
        Me.btn_entrar = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lbl_nome_func = New System.Windows.Forms.Label()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel6
        '
        Me.Panel6.Location = New System.Drawing.Point(251, 339)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(10, 120)
        Me.Panel6.TabIndex = 22
        '
        'Panel5
        '
        Me.Panel5.Location = New System.Drawing.Point(436, 365)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(11, 60)
        Me.Panel5.TabIndex = 21
        '
        'Panel4
        '
        Me.Panel4.Location = New System.Drawing.Point(259, 415)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(179, 10)
        Me.Panel4.TabIndex = 20
        '
        'Panel3
        '
        Me.Panel3.Location = New System.Drawing.Point(259, 362)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(179, 10)
        Me.Panel3.TabIndex = 19
        '
        'sla
        '
        Me.sla.AutoSize = True
        Me.sla.BackColor = System.Drawing.Color.FromArgb(CType(CType(17, Byte), Integer), CType(CType(17, Byte), Integer), CType(CType(17, Byte), Integer))
        Me.sla.Font = New System.Drawing.Font("Candara Light", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.sla.ForeColor = System.Drawing.SystemColors.Control
        Me.sla.Location = New System.Drawing.Point(226, 209)
        Me.sla.Name = "sla"
        Me.sla.Size = New System.Drawing.Size(48, 19)
        Me.sla.TabIndex = 14
        Me.sla.Text = "E-mail"
        '
        'txt_senha_login
        '
        Me.txt_senha_login.BackColor = System.Drawing.Color.FromArgb(CType(CType(17, Byte), Integer), CType(CType(17, Byte), Integer), CType(CType(17, Byte), Integer))
        Me.txt_senha_login.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txt_senha_login.Font = New System.Drawing.Font("Candara Light", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_senha_login.ForeColor = System.Drawing.SystemColors.Window
        Me.txt_senha_login.Location = New System.Drawing.Point(233, 292)
        Me.txt_senha_login.Name = "txt_senha_login"
        Me.txt_senha_login.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txt_senha_login.Size = New System.Drawing.Size(225, 19)
        Me.txt_senha_login.TabIndex = 13
        '
        'txt_email_login
        '
        Me.txt_email_login.BackColor = System.Drawing.Color.FromArgb(CType(CType(17, Byte), Integer), CType(CType(17, Byte), Integer), CType(CType(17, Byte), Integer))
        Me.txt_email_login.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txt_email_login.Font = New System.Drawing.Font("Candara Light", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_email_login.ForeColor = System.Drawing.SystemColors.Window
        Me.txt_email_login.Location = New System.Drawing.Point(233, 231)
        Me.txt_email_login.Name = "txt_email_login"
        Me.txt_email_login.Size = New System.Drawing.Size(225, 19)
        Me.txt_email_login.TabIndex = 12
        '
        'btn_entrar
        '
        Me.btn_entrar.BackColor = System.Drawing.Color.Transparent
        Me.btn_entrar.BackgroundImage = Global.VonAuto.My.Resources.Resources.Von_auto_entrar_
        Me.btn_entrar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.btn_entrar.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btn_entrar.ForeColor = System.Drawing.Color.White
        Me.btn_entrar.Location = New System.Drawing.Point(259, 362)
        Me.btn_entrar.Margin = New System.Windows.Forms.Padding(1)
        Me.btn_entrar.Name = "btn_entrar"
        Me.btn_entrar.Size = New System.Drawing.Size(179, 63)
        Me.btn_entrar.TabIndex = 18
        Me.btn_entrar.UseVisualStyleBackColor = False
        '
        'Panel1
        '
        Me.Panel1.BackgroundImage = Global.VonAuto.My.Resources.Resources.von3_transparente
        Me.Panel1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.Panel1.Location = New System.Drawing.Point(106, 4)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(457, 175)
        Me.Panel1.TabIndex = 15
        '
        'PictureBox1
        '
        Me.PictureBox1.BackgroundImage = Global.VonAuto.My.Resources.Resources.Linha
        Me.PictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.PictureBox1.Location = New System.Drawing.Point(230, 292)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(231, 41)
        Me.PictureBox1.TabIndex = 16
        Me.PictureBox1.TabStop = False
        '
        'Panel2
        '
        Me.Panel2.BackgroundImage = Global.VonAuto.My.Resources.Resources.Linha
        Me.Panel2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Location = New System.Drawing.Point(230, 219)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(231, 67)
        Me.Panel2.TabIndex = 17
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.Color.FromArgb(CType(CType(17, Byte), Integer), CType(CType(17, Byte), Integer), CType(CType(17, Byte), Integer))
        Me.Label2.Font = New System.Drawing.Font("Candara Light", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.Control
        Me.Label2.Location = New System.Drawing.Point(-4, 51)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(51, 19)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Senha"
        '
        'lbl_nome_func
        '
        Me.lbl_nome_func.AutoSize = True
        Me.lbl_nome_func.BackColor = System.Drawing.Color.FromArgb(CType(CType(17, Byte), Integer), CType(CType(17, Byte), Integer), CType(CType(17, Byte), Integer))
        Me.lbl_nome_func.Font = New System.Drawing.Font("Candara Light", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_nome_func.ForeColor = System.Drawing.SystemColors.Control
        Me.lbl_nome_func.Location = New System.Drawing.Point(609, 434)
        Me.lbl_nome_func.Name = "lbl_nome_func"
        Me.lbl_nome_func.Size = New System.Drawing.Size(50, 19)
        Me.lbl_nome_func.TabIndex = 23
        Me.lbl_nome_func.Text = "Nome"
        Me.lbl_nome_func.Visible = False
        '
        'Login
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(17, Byte), Integer), CType(CType(17, Byte), Integer), CType(CType(17, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(669, 462)
        Me.Controls.Add(Me.lbl_nome_func)
        Me.Controls.Add(Me.Panel6)
        Me.Controls.Add(Me.Panel5)
        Me.Controls.Add(Me.Panel4)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.btn_entrar)
        Me.Controls.Add(Me.sla)
        Me.Controls.Add(Me.txt_senha_login)
        Me.Controls.Add(Me.txt_email_login)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Panel2)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Login"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Von Auto - Login "
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Panel6 As System.Windows.Forms.Panel
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents btn_entrar As System.Windows.Forms.Button
    Friend WithEvents sla As System.Windows.Forms.Label
    Friend WithEvents txt_senha_login As System.Windows.Forms.TextBox
    Friend WithEvents txt_email_login As System.Windows.Forms.TextBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents lbl_nome_func As System.Windows.Forms.Label
End Class
