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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmHierarchySelection))
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
        Me.filterDropbox = New System.Windows.Forms.ComboBox()
        Me.filterLabel = New System.Windows.Forms.Label()
        Me.auswSpeichern = New System.Windows.Forms.Button()
        Me.SelectionSet = New System.Windows.Forms.PictureBox()
        Me.collapseCompletely = New System.Windows.Forms.PictureBox()
        Me.expandCompletely = New System.Windows.Forms.PictureBox()
        Me.SelectionReset = New System.Windows.Forms.PictureBox()
        Me.BackgroundWorker3 = New System.ComponentModel.BackgroundWorker()
        CType(Me.hryStufen, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SelectionSet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.collapseCompletely, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.expandCompletely, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SelectionReset, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'hryTreeView
        '
        Me.hryTreeView.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
        Me.einstellungen.Location = New System.Drawing.Point(476, 417)
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
        Me.labelPPTVorlage.Location = New System.Drawing.Point(9, 414)
        Me.labelPPTVorlage.Name = "labelPPTVorlage"
        Me.labelPPTVorlage.Size = New System.Drawing.Size(126, 16)
        Me.labelPPTVorlage.TabIndex = 39
        Me.labelPPTVorlage.Text = "Powerpoint Vorlage"
        Me.labelPPTVorlage.Visible = False
        '
        'statusLabel
        '
        Me.statusLabel.AutoSize = True
        Me.statusLabel.Location = New System.Drawing.Point(9, 470)
        Me.statusLabel.Name = "statusLabel"
        Me.statusLabel.Size = New System.Drawing.Size(39, 13)
        Me.statusLabel.TabIndex = 41
        Me.statusLabel.Text = "Label1"
        '
        'repVorlagenDropbox
        '
        Me.repVorlagenDropbox.FormattingEnabled = True
        Me.repVorlagenDropbox.Location = New System.Drawing.Point(145, 409)
        Me.repVorlagenDropbox.Name = "repVorlagenDropbox"
        Me.repVorlagenDropbox.Size = New System.Drawing.Size(264, 21)
        Me.repVorlagenDropbox.TabIndex = 37
        Me.repVorlagenDropbox.Visible = False
        '
        'OKButton
        '
        Me.OKButton.Location = New System.Drawing.Point(199, 443)
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
        Me.AbbrButton.Location = New System.Drawing.Point(344, 444)
        Me.AbbrButton.Name = "AbbrButton"
        Me.AbbrButton.Size = New System.Drawing.Size(113, 23)
        Me.AbbrButton.TabIndex = 44
        Me.AbbrButton.Text = "Abbrechen"
        Me.AbbrButton.UseVisualStyleBackColor = False
        Me.AbbrButton.Visible = False
        '
        'filterDropbox
        '
        Me.filterDropbox.FormattingEnabled = True
        Me.filterDropbox.Location = New System.Drawing.Point(145, 381)
        Me.filterDropbox.Name = "filterDropbox"
        Me.filterDropbox.Size = New System.Drawing.Size(264, 21)
        Me.filterDropbox.TabIndex = 48
        Me.filterDropbox.Visible = False
        '
        'filterLabel
        '
        Me.filterLabel.AutoSize = True
        Me.filterLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.filterLabel.Location = New System.Drawing.Point(9, 386)
        Me.filterLabel.Name = "filterLabel"
        Me.filterLabel.Size = New System.Drawing.Size(91, 16)
        Me.filterLabel.TabIndex = 49
        Me.filterLabel.Text = "Filter-Auswahl"
        Me.filterLabel.Visible = False
        '
        'auswSpeichern
        '
        Me.auswSpeichern.Location = New System.Drawing.Point(433, 379)
        Me.auswSpeichern.Name = "auswSpeichern"
        Me.auswSpeichern.Size = New System.Drawing.Size(113, 23)
        Me.auswSpeichern.TabIndex = 50
        Me.auswSpeichern.Text = "Speichern"
        Me.auswSpeichern.UseVisualStyleBackColor = True
        '
        'SelectionSet
        '
        Me.SelectionSet.BackColor = System.Drawing.SystemColors.Control
        Me.SelectionSet.ErrorImage = CType(resources.GetObject("SelectionSet.ErrorImage"), System.Drawing.Image)
        Me.SelectionSet.Image = CType(resources.GetObject("SelectionSet.Image"), System.Drawing.Image)
        Me.SelectionSet.InitialImage = Nothing
        Me.SelectionSet.Location = New System.Drawing.Point(12, 349)
        Me.SelectionSet.Name = "SelectionSet"
        Me.SelectionSet.Size = New System.Drawing.Size(16, 16)
        Me.SelectionSet.TabIndex = 51
        Me.SelectionSet.TabStop = False
        '
        'collapseCompletely
        '
        Me.collapseCompletely.BackColor = System.Drawing.SystemColors.Control
        Me.collapseCompletely.Image = CType(resources.GetObject("collapseCompletely.Image"), System.Drawing.Image)
        Me.collapseCompletely.Location = New System.Drawing.Point(68, 349)
        Me.collapseCompletely.Name = "collapseCompletely"
        Me.collapseCompletely.Size = New System.Drawing.Size(16, 16)
        Me.collapseCompletely.TabIndex = 47
        Me.collapseCompletely.TabStop = False
        '
        'expandCompletely
        '
        Me.expandCompletely.BackColor = System.Drawing.SystemColors.Control
        Me.expandCompletely.Image = CType(resources.GetObject("expandCompletely.Image"), System.Drawing.Image)
        Me.expandCompletely.Location = New System.Drawing.Point(90, 349)
        Me.expandCompletely.Name = "expandCompletely"
        Me.expandCompletely.Size = New System.Drawing.Size(16, 16)
        Me.expandCompletely.TabIndex = 46
        Me.expandCompletely.TabStop = False
        '
        'SelectionReset
        '
        Me.SelectionReset.BackColor = System.Drawing.SystemColors.Control
        Me.SelectionReset.Image = CType(resources.GetObject("SelectionReset.Image"), System.Drawing.Image)
        Me.SelectionReset.InitialImage = Nothing
        Me.SelectionReset.Location = New System.Drawing.Point(32, 349)
        Me.SelectionReset.Name = "SelectionReset"
        Me.SelectionReset.Size = New System.Drawing.Size(16, 16)
        Me.SelectionReset.TabIndex = 45
        Me.SelectionReset.TabStop = False
        '
        'BackgroundWorker3
        '
        '
        'frmHierarchySelection
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(558, 492)
        Me.Controls.Add(Me.SelectionSet)
        Me.Controls.Add(Me.auswSpeichern)
        Me.Controls.Add(Me.filterLabel)
        Me.Controls.Add(Me.filterDropbox)
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
        CType(Me.SelectionSet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.collapseCompletely, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.expandCompletely, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SelectionReset, System.ComponentModel.ISupportInitialize).EndInit()
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
    Friend WithEvents filterDropbox As System.Windows.Forms.ComboBox
    Friend WithEvents filterLabel As System.Windows.Forms.Label
    Friend WithEvents auswSpeichern As System.Windows.Forms.Button
    Friend WithEvents SelectionSet As System.Windows.Forms.PictureBox
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
End Class
