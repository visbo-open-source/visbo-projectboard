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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.benutzer = New System.Windows.Forms.TextBox()
        Me.maskedPwd = New System.Windows.Forms.TextBox()
        Me.messageBox = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(21, 38)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "LOGIN"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(21, 72)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(55, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Username"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(21, 112)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(50, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Passwort"
        '
        'benutzer
        '
        Me.benutzer.Location = New System.Drawing.Point(91, 65)
        Me.benutzer.Name = "benutzer"
        Me.benutzer.Size = New System.Drawing.Size(260, 20)
        Me.benutzer.TabIndex = 4
        '
        'maskedPwd
        '
        Me.maskedPwd.Location = New System.Drawing.Point(91, 105)
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
        Me.messageBox.Location = New System.Drawing.Point(21, 9)
        Me.messageBox.Name = "messageBox"
        Me.messageBox.Size = New System.Drawing.Size(0, 16)
        Me.messageBox.TabIndex = 7
        '
        'frmAuthentication
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(390, 158)
        Me.ControlBox = False
        Me.Controls.Add(Me.messageBox)
        Me.Controls.Add(Me.maskedPwd)
        Me.Controls.Add(Me.benutzer)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmAuthentication"
        Me.Text = "ProjectBoard LOGIN"
        Me.ResumeLayout(false)
        Me.PerformLayout

End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents benutzer As System.Windows.Forms.TextBox
    Public WithEvents maskedPwd As System.Windows.Forms.TextBox
    Friend WithEvents messageBox As System.Windows.Forms.Label
End Class
