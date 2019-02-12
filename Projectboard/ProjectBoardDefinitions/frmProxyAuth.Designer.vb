<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmProxyAuth
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmProxyAuth))
        Me.maskedPwd = New System.Windows.Forms.TextBox()
        Me.benutzer = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.messageBox = New System.Windows.Forms.TextBox()
        Me.domainBox = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.AbbrButton = New System.Windows.Forms.Button()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.proxyURLbox = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'maskedPwd
        '
        Me.maskedPwd.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.maskedPwd.BackColor = System.Drawing.Color.White
        Me.maskedPwd.Location = New System.Drawing.Point(112, 100)
        Me.maskedPwd.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.maskedPwd.Name = "maskedPwd"
        Me.maskedPwd.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.maskedPwd.Size = New System.Drawing.Size(279, 22)
        Me.maskedPwd.TabIndex = 9
        '
        'benutzer
        '
        Me.benutzer.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.benutzer.BackColor = System.Drawing.Color.White
        Me.benutzer.Location = New System.Drawing.Point(112, 70)
        Me.benutzer.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.benutzer.Name = "benutzer"
        Me.benutzer.Size = New System.Drawing.Size(279, 22)
        Me.benutzer.TabIndex = 8
        '
        'Label3
        '
        Me.Label3.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(16, 103)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(65, 17)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Passwort"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(16, 73)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(73, 17)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Username"
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.proxyURLbox)
        Me.Panel1.Controls.Add(Me.messageBox)
        Me.Panel1.Controls.Add(Me.domainBox)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.AbbrButton)
        Me.Panel1.Controls.Add(Me.OKButton)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.benutzer)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.maskedPwd)
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(411, 187)
        Me.Panel1.TabIndex = 10
        '
        'messageBox
        '
        Me.messageBox.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.messageBox.BackColor = System.Drawing.Color.White
        Me.messageBox.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.messageBox.Location = New System.Drawing.Point(19, 168)
        Me.messageBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.messageBox.Name = "messageBox"
        Me.messageBox.Size = New System.Drawing.Size(370, 15)
        Me.messageBox.TabIndex = 14
        '
        'domainBox
        '
        Me.domainBox.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.domainBox.BackColor = System.Drawing.Color.White
        Me.domainBox.Location = New System.Drawing.Point(113, 40)
        Me.domainBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.domainBox.Name = "domainBox"
        Me.domainBox.Size = New System.Drawing.Size(278, 22)
        Me.domainBox.TabIndex = 13
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(16, 43)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 17)
        Me.Label1.TabIndex = 12
        Me.Label1.Text = "Domain"
        '
        'AbbrButton
        '
        Me.AbbrButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.AbbrButton.Location = New System.Drawing.Point(254, 133)
        Me.AbbrButton.Name = "AbbrButton"
        Me.AbbrButton.Size = New System.Drawing.Size(135, 28)
        Me.AbbrButton.TabIndex = 11
        Me.AbbrButton.Text = "Abbrechen"
        Me.AbbrButton.UseVisualStyleBackColor = True
        '
        'OKButton
        '
        Me.OKButton.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.OKButton.Location = New System.Drawing.Point(19, 133)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(124, 28)
        Me.OKButton.TabIndex = 10
        Me.OKButton.Text = "OK"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'proxyURLbox
        '
        Me.proxyURLbox.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.proxyURLbox.BackColor = System.Drawing.Color.White
        Me.proxyURLbox.Location = New System.Drawing.Point(113, 10)
        Me.proxyURLbox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.proxyURLbox.Name = "proxyURLbox"
        Me.proxyURLbox.Size = New System.Drawing.Size(278, 22)
        Me.proxyURLbox.TabIndex = 15
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(16, 13)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(71, 17)
        Me.Label4.TabIndex = 16
        Me.Label4.Text = "ProxyURL"
        '
        'frmProxyAuth
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(120.0!, 120.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.ClientSize = New System.Drawing.Size(415, 187)
        Me.Controls.Add(Me.Panel1)
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmProxyAuth"
        Me.Text = "Proxy-Authentication"
        Me.TopMost = True
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Public WithEvents maskedPwd As Windows.Forms.TextBox
    Public WithEvents benutzer As Windows.Forms.TextBox
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents Panel1 As Windows.Forms.Panel
    Friend WithEvents AbbrButton As Windows.Forms.Button
    Friend WithEvents OKButton As Windows.Forms.Button
    Public WithEvents messageBox As Windows.Forms.TextBox
    Public WithEvents domainBox As Windows.Forms.TextBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents Label4 As Windows.Forms.Label
    Public WithEvents proxyURLbox As Windows.Forms.TextBox
End Class
