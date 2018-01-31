<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAuthentication
    Inherits System.Windows.Forms.Form

    'Das Formular überschreibt den Löschvorgang, um die Komponentenliste zu bereinigen.
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

    'Wird vom Windows Form-Designer benötigt.
    Private components As System.ComponentModel.IContainer

    'Hinweis: Die folgende Prozedur ist für den Windows Form-Designer erforderlich.
    'Das Bearbeiten ist mit dem Windows Form-Designer möglich.  
    'Das Bearbeiten mit dem Code-Editor ist nicht möglich.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAuthentication))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.benutzer = New System.Windows.Forms.TextBox()
        Me.maskedPwd = New System.Windows.Forms.TextBox()
        Me.messageBox = New System.Windows.Forms.Label()
        Me.abbrButton = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.chbx_remember = New System.Windows.Forms.CheckBox()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(234, 74)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(0, 24)
        Me.Label1.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(13, 125)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(55, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Username"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(13, 165)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(50, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Passwort"
        '
        'benutzer
        '
        Me.benutzer.BackColor = System.Drawing.Color.WhiteSmoke
        Me.benutzer.Location = New System.Drawing.Point(91, 122)
        Me.benutzer.Name = "benutzer"
        Me.benutzer.Size = New System.Drawing.Size(260, 20)
        Me.benutzer.TabIndex = 4
        '
        'maskedPwd
        '
        Me.maskedPwd.BackColor = System.Drawing.Color.WhiteSmoke
        Me.maskedPwd.Location = New System.Drawing.Point(91, 161)
        Me.maskedPwd.Name = "maskedPwd"
        Me.maskedPwd.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.maskedPwd.Size = New System.Drawing.Size(260, 20)
        Me.maskedPwd.TabIndex = 5
        '
        'messageBox
        '
        Me.messageBox.AutoSize = True
        Me.messageBox.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.messageBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.messageBox.ForeColor = System.Drawing.Color.Red
        Me.messageBox.Location = New System.Drawing.Point(12, 9)
        Me.messageBox.Name = "messageBox"
        Me.messageBox.Size = New System.Drawing.Size(0, 16)
        Me.messageBox.TabIndex = 7
        '
        'abbrButton
        '
        Me.abbrButton.DialogResult = System.Windows.Forms.DialogResult.Abort
        Me.abbrButton.Location = New System.Drawing.Point(254, 235)
        Me.abbrButton.Name = "abbrButton"
        Me.abbrButton.Size = New System.Drawing.Size(96, 23)
        Me.abbrButton.TabIndex = 8
        Me.abbrButton.Text = "Abbrechen"
        Me.abbrButton.UseVisualStyleBackColor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.InitialImage = Nothing
        Me.PictureBox1.Location = New System.Drawing.Point(15, 27)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(193, 73)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox1.TabIndex = 9
        Me.PictureBox1.TabStop = False
        '
        'OKButton
        '
        Me.OKButton.Location = New System.Drawing.Point(91, 235)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(102, 23)
        Me.OKButton.TabIndex = 10
        Me.OKButton.Text = "OK"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'chbx_remember
        '
        Me.chbx_remember.AutoSize = True
        Me.chbx_remember.Location = New System.Drawing.Point(14, 200)
        Me.chbx_remember.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.chbx_remember.Name = "chbx_remember"
        Me.chbx_remember.Size = New System.Drawing.Size(110, 17)
        Me.chbx_remember.TabIndex = 11
        Me.chbx_remember.Text = "Remember Me     "
        Me.chbx_remember.UseVisualStyleBackColor = True
        '
        'frmAuthentication
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(390, 266)
        Me.ControlBox = False
        Me.Controls.Add(Me.chbx_remember)
        Me.Controls.Add(Me.OKButton)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.abbrButton)
        Me.Controls.Add(Me.messageBox)
        Me.Controls.Add(Me.maskedPwd)
        Me.Controls.Add(Me.benutzer)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Name = "frmAuthentication"
        Me.Text = "LOGIN"
        Me.TopMost = True
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents benutzer As System.Windows.Forms.TextBox
    Public WithEvents maskedPwd As System.Windows.Forms.TextBox
    Friend WithEvents messageBox As System.Windows.Forms.Label
    Friend WithEvents abbrButton As System.Windows.Forms.Button
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents OKButton As System.Windows.Forms.Button
    Friend WithEvents chbx_remember As Windows.Forms.CheckBox
End Class
