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
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.AbbrButton = New System.Windows.Forms.Button()
        Me.SelectionReset = New System.Windows.Forms.PictureBox()
        Me.expandCompletely = New System.Windows.Forms.PictureBox()
        Me.collapseCompletely = New System.Windows.Forms.PictureBox()
        CType(Me.hryStufen, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SelectionReset, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.expandCompletely, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.collapseCompletely, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.hryStufenLabel.Size = New System.Drawing.Size(300, 13)
        Me.hryStufenLabel.TabIndex = 35
        Me.hryStufenLabel.Text = "wie viele ""Eltern"" sollen bei der Suche berücksichtigt werden?"
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
        Me.statusLabel.Location = New System.Drawing.Point(12, 444)
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
        'BackgroundWorker1
        '
        Me.BackgroundWorker1.WorkerReportsProgress = True
        Me.BackgroundWorker1.WorkerSupportsCancellation = True
        '
        'AbbrButton
        '
        Me.AbbrButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.AbbrButton.Location = New System.Drawing.Point(342, 418)
        Me.AbbrButton.Name = "AbbrButton"
        Me.AbbrButton.Size = New System.Drawing.Size(113, 23)
        Me.AbbrButton.TabIndex = 44
        Me.AbbrButton.Text = "Abbrechen"
        Me.AbbrButton.UseVisualStyleBackColor = False
        Me.AbbrButton.Visible = False
        '
        'SelectionReset
        '
        Me.SelectionReset.BackColor = System.Drawing.SystemColors.Info
        Me.SelectionReset.Location = New System.Drawing.Point(12, 349)
        Me.SelectionReset.Name = "SelectionReset"
        Me.SelectionReset.Size = New System.Drawing.Size(19, 18)
        Me.SelectionReset.TabIndex = 45
        Me.SelectionReset.TabStop = False
        '
        'expandCompletely
        '
        Me.expandCompletely.BackColor = System.Drawing.Color.OliveDrab
        Me.expandCompletely.Location = New System.Drawing.Point(60, 349)
        Me.expandCompletely.Name = "expandCompletely"
        Me.expandCompletely.Size = New System.Drawing.Size(19, 18)
        Me.expandCompletely.TabIndex = 46
        Me.expandCompletely.TabStop = False
        '
        'collapseCompletely
        '
        Me.collapseCompletely.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.collapseCompletely.Location = New System.Drawing.Point(85, 349)
        Me.collapseCompletely.Name = "collapseCompletely"
        Me.collapseCompletely.Size = New System.Drawing.Size(19, 18)
        Me.collapseCompletely.TabIndex = 47
        Me.collapseCompletely.TabStop = False
        '
        'frmHierarchySelection
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(558, 466)
        Me.Controls.Add(Me.collapseCompletely)
        Me.Controls.Add(Me.expandCompletely)
        Me.Controls.Add(Me.SelectionReset)
        Me.Controls.Add(Me.AbbrButton)
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
        Me.TopMost = True
        CType(Me.hryStufen, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SelectionReset, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.expandCompletely, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.collapseCompletely, System.ComponentModel.ISupportInitialize).EndInit()
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
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents AbbrButton As System.Windows.Forms.Button
    Friend WithEvents SelectionReset As System.Windows.Forms.PictureBox
    Friend WithEvents expandCompletely As System.Windows.Forms.PictureBox
    Friend WithEvents collapseCompletely As System.Windows.Forms.PictureBox
End Class
