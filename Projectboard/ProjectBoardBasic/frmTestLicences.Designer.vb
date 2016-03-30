<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTestLicences
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
        Me.LabelUser = New System.Windows.Forms.Label()
        Me.UserName = New System.Windows.Forms.TextBox()
        Me.LabelKomponente = New System.Windows.Forms.Label()
        Me.ListKomponenten = New System.Windows.Forms.ListBox()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.statusLabel = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'LabelUser
        '
        Me.LabelUser.AutoSize = True
        Me.LabelUser.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelUser.Location = New System.Drawing.Point(28, 30)
        Me.LabelUser.Name = "LabelUser"
        Me.LabelUser.Size = New System.Drawing.Size(74, 16)
        Me.LabelUser.TabIndex = 3
        Me.LabelUser.Text = "Username:"
        '
        'UserName
        '
        Me.UserName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UserName.Location = New System.Drawing.Point(212, 24)
        Me.UserName.Name = "UserName"
        Me.UserName.Size = New System.Drawing.Size(282, 22)
        Me.UserName.TabIndex = 8
        '
        'LabelKomponente
        '
        Me.LabelKomponente.AutoSize = True
        Me.LabelKomponente.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelKomponente.Location = New System.Drawing.Point(28, 69)
        Me.LabelKomponente.Name = "LabelKomponente"
        Me.LabelKomponente.Size = New System.Drawing.Size(139, 16)
        Me.LabelKomponente.TabIndex = 9
        Me.LabelKomponente.Text = "SoftwareKomponente:"
        '
        'ListKomponenten
        '
        Me.ListKomponenten.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ListKomponenten.FormattingEnabled = True
        Me.ListKomponenten.HorizontalScrollbar = True
        Me.ListKomponenten.ItemHeight = 16
        Me.ListKomponenten.Location = New System.Drawing.Point(212, 69)
        Me.ListKomponenten.Name = "ListKomponenten"
        Me.ListKomponenten.Size = New System.Drawing.Size(282, 116)
        Me.ListKomponenten.Sorted = True
        Me.ListKomponenten.TabIndex = 10
        '
        'OKButton
        '
        Me.OKButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OKButton.Location = New System.Drawing.Point(149, 204)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(197, 23)
        Me.OKButton.TabIndex = 11
        Me.OKButton.Text = "Test Licence"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'statusLabel
        '
        Me.statusLabel.AutoSize = True
        Me.statusLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.statusLabel.Location = New System.Drawing.Point(28, 243)
        Me.statusLabel.Name = "statusLabel"
        Me.statusLabel.Size = New System.Drawing.Size(51, 17)
        Me.statusLabel.TabIndex = 44
        Me.statusLabel.Text = "Label1"
        Me.statusLabel.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        Me.statusLabel.Visible = False
        '
        'frmTestLicences
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(563, 269)
        Me.Controls.Add(Me.statusLabel)
        Me.Controls.Add(Me.OKButton)
        Me.Controls.Add(Me.ListKomponenten)
        Me.Controls.Add(Me.LabelKomponente)
        Me.Controls.Add(Me.UserName)
        Me.Controls.Add(Me.LabelUser)
        Me.Name = "frmTestLicences"
        Me.Text = "Test der Lizenzen"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents LabelUser As System.Windows.Forms.Label
    Friend WithEvents UserName As System.Windows.Forms.TextBox
    Friend WithEvents LabelKomponente As System.Windows.Forms.Label
    Friend WithEvents ListKomponenten As System.Windows.Forms.ListBox
    Friend WithEvents OKButton As System.Windows.Forms.Button
    Friend WithEvents statusLabel As System.Windows.Forms.Label
End Class
