<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmProjekteSpeichern
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
        Me.JAButton = New System.Windows.Forms.Button()
        Me.NEINButton = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(10, 19)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(233, 15)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Möchten Sie alle Änderungen speichern?"
        '
        'JAButton
        '
        Me.JAButton.DialogResult = System.Windows.Forms.DialogResult.Yes
        Me.JAButton.Location = New System.Drawing.Point(10, 56)
        Me.JAButton.Margin = New System.Windows.Forms.Padding(2)
        Me.JAButton.Name = "JAButton"
        Me.JAButton.Size = New System.Drawing.Size(83, 26)
        Me.JAButton.TabIndex = 1
        Me.JAButton.Text = "JA"
        Me.JAButton.UseVisualStyleBackColor = True
        '
        'NEINButton
        '
        Me.NEINButton.DialogResult = System.Windows.Forms.DialogResult.No
        Me.NEINButton.Location = New System.Drawing.Point(134, 56)
        Me.NEINButton.Margin = New System.Windows.Forms.Padding(2)
        Me.NEINButton.Name = "NEINButton"
        Me.NEINButton.Size = New System.Drawing.Size(87, 26)
        Me.NEINButton.TabIndex = 2
        Me.NEINButton.Text = "NEIN"
        Me.NEINButton.UseVisualStyleBackColor = True
        '
        'frmProjekteSpeichern
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.ClientSize = New System.Drawing.Size(266, 92)
        Me.Controls.Add(Me.NEINButton)
        Me.Controls.Add(Me.JAButton)
        Me.Controls.Add(Me.Label1)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "frmProjekteSpeichern"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "Exit Projectboard"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents JAButton As System.Windows.Forms.Button
    Friend WithEvents NEINButton As System.Windows.Forms.Button
End Class
