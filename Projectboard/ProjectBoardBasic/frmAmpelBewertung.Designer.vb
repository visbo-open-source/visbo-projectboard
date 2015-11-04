<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAmpelBewertung
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
        Me.OKButton = New System.Windows.Forms.Button()
        Me.AbbrButton = New System.Windows.Forms.Button()
        Me.erlaeuterung = New System.Windows.Forms.TextBox()
        Me.ampelGruen = New System.Windows.Forms.RadioButton()
        Me.ampelGelb = New System.Windows.Forms.RadioButton()
        Me.ampelRot = New System.Windows.Forms.RadioButton()
        Me.SuspendLayout()
        '
        'OKButton
        '
        Me.OKButton.Location = New System.Drawing.Point(145, 250)
        Me.OKButton.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(100, 28)
        Me.OKButton.TabIndex = 0
        Me.OKButton.Text = "OK"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'AbbrButton
        '
        Me.AbbrButton.Location = New System.Drawing.Point(344, 250)
        Me.AbbrButton.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.AbbrButton.Name = "AbbrButton"
        Me.AbbrButton.Size = New System.Drawing.Size(100, 28)
        Me.AbbrButton.TabIndex = 1
        Me.AbbrButton.Text = "Abbrechen"
        Me.AbbrButton.UseVisualStyleBackColor = True
        '
        'erlaeuterung
        '
        Me.erlaeuterung.Location = New System.Drawing.Point(24, 32)
        Me.erlaeuterung.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.erlaeuterung.MaximumSize = New System.Drawing.Size(532, 122)
        Me.erlaeuterung.MinimumSize = New System.Drawing.Size(265, 122)
        Me.erlaeuterung.Multiline = True
        Me.erlaeuterung.Name = "erlaeuterung"
        Me.erlaeuterung.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.erlaeuterung.Size = New System.Drawing.Size(532, 122)
        Me.erlaeuterung.TabIndex = 2
        '
        'ampelGruen
        '
        Me.ampelGruen.AutoSize = True
        Me.ampelGruen.Location = New System.Drawing.Point(71, 188)
        Me.ampelGruen.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ampelGruen.Name = "ampelGruen"
        Me.ampelGruen.Size = New System.Drawing.Size(61, 21)
        Me.ampelGruen.TabIndex = 3
        Me.ampelGruen.TabStop = True
        Me.ampelGruen.Text = "Grün"
        Me.ampelGruen.UseVisualStyleBackColor = True
        '
        'ampelGelb
        '
        Me.ampelGelb.AutoSize = True
        Me.ampelGelb.Location = New System.Drawing.Point(267, 188)
        Me.ampelGelb.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ampelGelb.Name = "ampelGelb"
        Me.ampelGelb.Size = New System.Drawing.Size(59, 21)
        Me.ampelGelb.TabIndex = 4
        Me.ampelGelb.TabStop = True
        Me.ampelGelb.Text = "Gelb"
        Me.ampelGelb.UseVisualStyleBackColor = True
        '
        'ampelRot
        '
        Me.ampelRot.AutoSize = True
        Me.ampelRot.Location = New System.Drawing.Point(461, 188)
        Me.ampelRot.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ampelRot.Name = "ampelRot"
        Me.ampelRot.Size = New System.Drawing.Size(51, 21)
        Me.ampelRot.TabIndex = 5
        Me.ampelRot.TabStop = True
        Me.ampelRot.Text = "Rot"
        Me.ampelRot.UseVisualStyleBackColor = True
        '
        'frmAmpelBewertung
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(120.0!, 120.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.ClientSize = New System.Drawing.Size(576, 322)
        Me.Controls.Add(Me.ampelRot)
        Me.Controls.Add(Me.ampelGelb)
        Me.Controls.Add(Me.ampelGruen)
        Me.Controls.Add(Me.erlaeuterung)
        Me.Controls.Add(Me.AbbrButton)
        Me.Controls.Add(Me.OKButton)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Name = "frmAmpelBewertung"
        Me.Text = "Ampel Bewertung"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents OKButton As System.Windows.Forms.Button
    Friend WithEvents AbbrButton As System.Windows.Forms.Button
    Friend WithEvents erlaeuterung As System.Windows.Forms.TextBox
    Friend WithEvents ampelGruen As System.Windows.Forms.RadioButton
    Friend WithEvents ampelGelb As System.Windows.Forms.RadioButton
    Friend WithEvents ampelRot As System.Windows.Forms.RadioButton
End Class
