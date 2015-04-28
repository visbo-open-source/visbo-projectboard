<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmHierarchySelection
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
        Me.hryTreeView = New System.Windows.Forms.TreeView()
        Me.hryStufenLabel = New System.Windows.Forms.Label()
        Me.hryStufen = New System.Windows.Forms.NumericUpDown()
        Me.einstellungen = New System.Windows.Forms.Label()
        Me.chkbxOneChart = New System.Windows.Forms.CheckBox()
        Me.labelPPTVorlage = New System.Windows.Forms.Label()
        Me.statusLabel = New System.Windows.Forms.Label()
        Me.repVorlagenDropbox = New System.Windows.Forms.ComboBox()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.hryStufen, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'hryTreeView
        '
        Me.hryTreeView.FullRowSelect = True
        Me.hryTreeView.Location = New System.Drawing.Point(12, 50)
        Me.hryTreeView.Name = "hryTreeView"
        Me.hryTreeView.Size = New System.Drawing.Size(534, 291)
        Me.hryTreeView.TabIndex = 32
        '
        'hryStufenLabel
        '
        Me.hryStufenLabel.AutoSize = True
        Me.hryStufenLabel.Location = New System.Drawing.Point(12, 24)
        Me.hryStufenLabel.Name = "hryStufenLabel"
        Me.hryStufenLabel.Size = New System.Drawing.Size(345, 13)
        Me.hryStufenLabel.TabIndex = 35
        Me.hryStufenLabel.Text = "wie viele Hierarchie-Stufen sollen bei der Suche berücksichtigt werden?"
        '
        'hryStufen
        '
        Me.hryStufen.Location = New System.Drawing.Point(388, 24)
        Me.hryStufen.Name = "hryStufen"
        Me.hryStufen.Size = New System.Drawing.Size(57, 20)
        Me.hryStufen.TabIndex = 34
        '
        'einstellungen
        '
        Me.einstellungen.AutoSize = True
        Me.einstellungen.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, CType((System.Drawing.FontStyle.Italic Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.einstellungen.ForeColor = System.Drawing.SystemColors.MenuHighlight
        Me.einstellungen.Location = New System.Drawing.Point(476, 381)
        Me.einstellungen.Name = "einstellungen"
        Me.einstellungen.Size = New System.Drawing.Size(70, 13)
        Me.einstellungen.TabIndex = 42
        Me.einstellungen.Text = "Einstellungen"
        Me.einstellungen.Visible = False
        '
        'chkbxOneChart
        '
        Me.chkbxOneChart.AutoSize = True
        Me.chkbxOneChart.Location = New System.Drawing.Point(428, 347)
        Me.chkbxOneChart.Name = "chkbxOneChart"
        Me.chkbxOneChart.Size = New System.Drawing.Size(118, 17)
        Me.chkbxOneChart.TabIndex = 36
        Me.chkbxOneChart.Text = "Alles in einem Chart"
        Me.chkbxOneChart.UseVisualStyleBackColor = True
        Me.chkbxOneChart.Visible = False
        '
        'labelPPTVorlage
        '
        Me.labelPPTVorlage.AutoSize = True
        Me.labelPPTVorlage.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelPPTVorlage.Location = New System.Drawing.Point(12, 378)
        Me.labelPPTVorlage.Name = "labelPPTVorlage"
        Me.labelPPTVorlage.Size = New System.Drawing.Size(126, 16)
        Me.labelPPTVorlage.TabIndex = 39
        Me.labelPPTVorlage.Text = "Powerpoint Vorlage"
        Me.labelPPTVorlage.Visible = False
        '
        'statusLabel
        '
        Me.statusLabel.AutoSize = True
        Me.statusLabel.Location = New System.Drawing.Point(-52, 258)
        Me.statusLabel.Name = "statusLabel"
        Me.statusLabel.Size = New System.Drawing.Size(39, 13)
        Me.statusLabel.TabIndex = 41
        Me.statusLabel.Text = "Label1"
        '
        'repVorlagenDropbox
        '
        Me.repVorlagenDropbox.FormattingEnabled = True
        Me.repVorlagenDropbox.Location = New System.Drawing.Point(147, 376)
        Me.repVorlagenDropbox.Name = "repVorlagenDropbox"
        Me.repVorlagenDropbox.Size = New System.Drawing.Size(264, 21)
        Me.repVorlagenDropbox.TabIndex = 37
        Me.repVorlagenDropbox.Visible = False
        '
        'OKButton
        '
        Me.OKButton.Location = New System.Drawing.Point(223, 418)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(113, 23)
        Me.OKButton.TabIndex = 38
        Me.OKButton.Text = "Anzeigen"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 445)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(0, 13)
        Me.Label1.TabIndex = 43
        '
        'frmHierarchySelection
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(558, 467)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.einstellungen)
        Me.Controls.Add(Me.chkbxOneChart)
        Me.Controls.Add(Me.labelPPTVorlage)
        Me.Controls.Add(Me.statusLabel)
        Me.Controls.Add(Me.repVorlagenDropbox)
        Me.Controls.Add(Me.OKButton)
        Me.Controls.Add(Me.hryStufenLabel)
        Me.Controls.Add(Me.hryStufen)
        Me.Controls.Add(Me.hryTreeView)
        Me.Name = "frmHierarchySelection"
        Me.Text = "Auswahl über Hierarchie"
        CType(Me.hryStufen, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents hryTreeView As System.Windows.Forms.TreeView
    Friend WithEvents hryStufenLabel As System.Windows.Forms.Label
    Friend WithEvents hryStufen As System.Windows.Forms.NumericUpDown
    Friend WithEvents einstellungen As System.Windows.Forms.Label
    Friend WithEvents chkbxOneChart As System.Windows.Forms.CheckBox
    Friend WithEvents labelPPTVorlage As System.Windows.Forms.Label
    Friend WithEvents statusLabel As System.Windows.Forms.Label
    Friend WithEvents repVorlagenDropbox As System.Windows.Forms.ComboBox
    Friend WithEvents OKButton As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
