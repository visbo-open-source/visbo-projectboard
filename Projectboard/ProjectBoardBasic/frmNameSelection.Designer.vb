<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmNameSelection
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmNameSelection))
        Me.nameListBox = New System.Windows.Forms.ListBox()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.AbbrButton = New System.Windows.Forms.Button()
        Me.filterBox = New System.Windows.Forms.TextBox()
        Me.headerLine = New System.Windows.Forms.Label()
        Me.rdbPhases = New System.Windows.Forms.RadioButton()
        Me.rdbMilestones = New System.Windows.Forms.RadioButton()
        Me.rdbRoles = New System.Windows.Forms.RadioButton()
        Me.rdbCosts = New System.Windows.Forms.RadioButton()
        Me.pictureCosts = New System.Windows.Forms.PictureBox()
        Me.pictureRoles = New System.Windows.Forms.PictureBox()
        Me.picturePhasen = New System.Windows.Forms.PictureBox()
        Me.pictureMilestones = New System.Windows.Forms.PictureBox()
        Me.chkbxOneChart = New System.Windows.Forms.CheckBox()
        Me.repVorlagenDropbox = New System.Windows.Forms.ComboBox()
        Me.labelPPTVorlage = New System.Windows.Forms.Label()
        Me.statusLabel = New System.Windows.Forms.Label()
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.einstellungen = New System.Windows.Forms.Label()
        Me.selNameListBox = New System.Windows.Forms.ListBox()
        Me.pictureTyp = New System.Windows.Forms.PictureBox()
        Me.rdbTyp = New System.Windows.Forms.RadioButton()
        Me.rdbBU = New System.Windows.Forms.RadioButton()
        Me.pictureBU = New System.Windows.Forms.PictureBox()
        Me.addButton = New System.Windows.Forms.PictureBox()
        Me.removeButton = New System.Windows.Forms.PictureBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.filterDropbox = New System.Windows.Forms.ComboBox()
        Me.filterLabel = New System.Windows.Forms.Label()
        Me.auswSpeichern = New System.Windows.Forms.Button()
        CType(Me.pictureCosts, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pictureRoles, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picturePhasen, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pictureMilestones, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pictureTyp, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pictureBU, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.addButton, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.removeButton, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'nameListBox
        '
        Me.nameListBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.nameListBox.FormattingEnabled = True
        Me.nameListBox.ItemHeight = 16
        Me.nameListBox.Location = New System.Drawing.Point(12, 109)
        Me.nameListBox.Name = "nameListBox"
        Me.nameListBox.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.nameListBox.Size = New System.Drawing.Size(275, 196)
        Me.nameListBox.Sorted = True
        Me.nameListBox.TabIndex = 0
        '
        'OKButton
        '
        Me.OKButton.Location = New System.Drawing.Point(162, 404)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(113, 23)
        Me.OKButton.TabIndex = 9
        Me.OKButton.Text = "Anzeigen"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'AbbrButton
        '
        Me.AbbrButton.Location = New System.Drawing.Point(313, 404)
        Me.AbbrButton.Name = "AbbrButton"
        Me.AbbrButton.Size = New System.Drawing.Size(113, 23)
        Me.AbbrButton.TabIndex = 10
        Me.AbbrButton.Text = "Zurücksetzen"
        Me.AbbrButton.UseVisualStyleBackColor = True
        '
        'filterBox
        '
        Me.filterBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.filterBox.HideSelection = False
        Me.filterBox.Location = New System.Drawing.Point(137, 76)
        Me.filterBox.Name = "filterBox"
        Me.filterBox.Size = New System.Drawing.Size(150, 22)
        Me.filterBox.TabIndex = 11
        '
        'headerLine
        '
        Me.headerLine.AutoSize = True
        Me.headerLine.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.headerLine.Location = New System.Drawing.Point(12, 76)
        Me.headerLine.Name = "headerLine"
        Me.headerLine.Size = New System.Drawing.Size(91, 16)
        Me.headerLine.TabIndex = 12
        Me.headerLine.Text = "Label1              "
        '
        'rdbPhases
        '
        Me.rdbPhases.AutoSize = True
        Me.rdbPhases.Location = New System.Drawing.Point(15, 32)
        Me.rdbPhases.Name = "rdbPhases"
        Me.rdbPhases.Size = New System.Drawing.Size(14, 13)
        Me.rdbPhases.TabIndex = 2
        Me.rdbPhases.TabStop = True
        Me.rdbPhases.UseVisualStyleBackColor = True
        '
        'rdbMilestones
        '
        Me.rdbMilestones.AutoSize = True
        Me.rdbMilestones.Location = New System.Drawing.Point(114, 32)
        Me.rdbMilestones.Name = "rdbMilestones"
        Me.rdbMilestones.Size = New System.Drawing.Size(14, 13)
        Me.rdbMilestones.TabIndex = 3
        Me.rdbMilestones.TabStop = True
        Me.rdbMilestones.UseVisualStyleBackColor = True
        '
        'rdbRoles
        '
        Me.rdbRoles.AutoSize = True
        Me.rdbRoles.Location = New System.Drawing.Point(213, 32)
        Me.rdbRoles.Name = "rdbRoles"
        Me.rdbRoles.Size = New System.Drawing.Size(14, 13)
        Me.rdbRoles.TabIndex = 4
        Me.rdbRoles.TabStop = True
        Me.rdbRoles.UseVisualStyleBackColor = True
        '
        'rdbCosts
        '
        Me.rdbCosts.AutoSize = True
        Me.rdbCosts.Location = New System.Drawing.Point(325, 32)
        Me.rdbCosts.Name = "rdbCosts"
        Me.rdbCosts.Size = New System.Drawing.Size(14, 13)
        Me.rdbCosts.TabIndex = 5
        Me.rdbCosts.TabStop = True
        Me.rdbCosts.UseVisualStyleBackColor = True
        '
        'pictureCosts
        '
        Me.pictureCosts.Image = CType(resources.GetObject("pictureCosts.Image"), System.Drawing.Image)
        Me.pictureCosts.Location = New System.Drawing.Point(344, 21)
        Me.pictureCosts.Name = "pictureCosts"
        Me.pictureCosts.Size = New System.Drawing.Size(33, 33)
        Me.pictureCosts.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.pictureCosts.TabIndex = 17
        Me.pictureCosts.TabStop = False
        '
        'pictureRoles
        '
        Me.pictureRoles.Image = CType(resources.GetObject("pictureRoles.Image"), System.Drawing.Image)
        Me.pictureRoles.Location = New System.Drawing.Point(242, 21)
        Me.pictureRoles.Name = "pictureRoles"
        Me.pictureRoles.Size = New System.Drawing.Size(33, 33)
        Me.pictureRoles.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.pictureRoles.TabIndex = 14
        Me.pictureRoles.TabStop = False
        '
        'picturePhasen
        '
        Me.picturePhasen.Image = CType(resources.GetObject("picturePhasen.Image"), System.Drawing.Image)
        Me.picturePhasen.Location = New System.Drawing.Point(38, 21)
        Me.picturePhasen.Name = "picturePhasen"
        Me.picturePhasen.Size = New System.Drawing.Size(33, 33)
        Me.picturePhasen.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picturePhasen.TabIndex = 18
        Me.picturePhasen.TabStop = False
        '
        'pictureMilestones
        '
        Me.pictureMilestones.Image = CType(resources.GetObject("pictureMilestones.Image"), System.Drawing.Image)
        Me.pictureMilestones.Location = New System.Drawing.Point(140, 21)
        Me.pictureMilestones.Name = "pictureMilestones"
        Me.pictureMilestones.Size = New System.Drawing.Size(33, 33)
        Me.pictureMilestones.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.pictureMilestones.TabIndex = 19
        Me.pictureMilestones.TabStop = False
        '
        'chkbxOneChart
        '
        Me.chkbxOneChart.AutoSize = True
        Me.chkbxOneChart.Location = New System.Drawing.Point(469, 311)
        Me.chkbxOneChart.Name = "chkbxOneChart"
        Me.chkbxOneChart.Size = New System.Drawing.Size(118, 17)
        Me.chkbxOneChart.TabIndex = 8
        Me.chkbxOneChart.Text = "Alles in einem Chart"
        Me.chkbxOneChart.UseVisualStyleBackColor = True
        Me.chkbxOneChart.Visible = False
        '
        'repVorlagenDropbox
        '
        Me.repVorlagenDropbox.FormattingEnabled = True
        Me.repVorlagenDropbox.Location = New System.Drawing.Point(162, 364)
        Me.repVorlagenDropbox.Name = "repVorlagenDropbox"
        Me.repVorlagenDropbox.Size = New System.Drawing.Size(264, 21)
        Me.repVorlagenDropbox.TabIndex = 9
        Me.repVorlagenDropbox.Visible = False
        '
        'labelPPTVorlage
        '
        Me.labelPPTVorlage.AutoSize = True
        Me.labelPPTVorlage.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelPPTVorlage.Location = New System.Drawing.Point(12, 369)
        Me.labelPPTVorlage.Name = "labelPPTVorlage"
        Me.labelPPTVorlage.Size = New System.Drawing.Size(126, 16)
        Me.labelPPTVorlage.TabIndex = 10
        Me.labelPPTVorlage.Text = "Powerpoint Vorlage"
        Me.labelPPTVorlage.Visible = False
        '
        'statusLabel
        '
        Me.statusLabel.AutoSize = True
        Me.statusLabel.Location = New System.Drawing.Point(12, 432)
        Me.statusLabel.Name = "statusLabel"
        Me.statusLabel.Size = New System.Drawing.Size(39, 13)
        Me.statusLabel.TabIndex = 21
        Me.statusLabel.Text = "Label1"
        '
        'BackgroundWorker1
        '
        Me.BackgroundWorker1.WorkerReportsProgress = True
        Me.BackgroundWorker1.WorkerSupportsCancellation = True
        '
        'einstellungen
        '
        Me.einstellungen.AutoSize = True
        Me.einstellungen.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, CType((System.Drawing.FontStyle.Italic Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.einstellungen.ForeColor = System.Drawing.SystemColors.MenuHighlight
        Me.einstellungen.Location = New System.Drawing.Point(517, 372)
        Me.einstellungen.Name = "einstellungen"
        Me.einstellungen.Size = New System.Drawing.Size(70, 13)
        Me.einstellungen.TabIndex = 22
        Me.einstellungen.Text = "Einstellungen"
        Me.einstellungen.Visible = False
        '
        'selNameListBox
        '
        Me.selNameListBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.selNameListBox.FormattingEnabled = True
        Me.selNameListBox.ItemHeight = 16
        Me.selNameListBox.Location = New System.Drawing.Point(312, 109)
        Me.selNameListBox.Name = "selNameListBox"
        Me.selNameListBox.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.selNameListBox.Size = New System.Drawing.Size(275, 196)
        Me.selNameListBox.Sorted = True
        Me.selNameListBox.TabIndex = 23
        '
        'pictureTyp
        '
        Me.pictureTyp.Image = CType(resources.GetObject("pictureTyp.Image"), System.Drawing.Image)
        Me.pictureTyp.Location = New System.Drawing.Point(554, 21)
        Me.pictureTyp.Name = "pictureTyp"
        Me.pictureTyp.Size = New System.Drawing.Size(33, 33)
        Me.pictureTyp.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.pictureTyp.TabIndex = 24
        Me.pictureTyp.TabStop = False
        '
        'rdbTyp
        '
        Me.rdbTyp.AutoSize = True
        Me.rdbTyp.Location = New System.Drawing.Point(534, 32)
        Me.rdbTyp.Name = "rdbTyp"
        Me.rdbTyp.Size = New System.Drawing.Size(14, 13)
        Me.rdbTyp.TabIndex = 25
        Me.rdbTyp.TabStop = True
        Me.rdbTyp.UseVisualStyleBackColor = True
        '
        'rdbBU
        '
        Me.rdbBU.AutoSize = True
        Me.rdbBU.Location = New System.Drawing.Point(430, 32)
        Me.rdbBU.Name = "rdbBU"
        Me.rdbBU.Size = New System.Drawing.Size(14, 13)
        Me.rdbBU.TabIndex = 26
        Me.rdbBU.TabStop = True
        Me.rdbBU.UseVisualStyleBackColor = True
        '
        'pictureBU
        '
        Me.pictureBU.Image = Global.ProjectBoardBasic.My.Resources.Resources.branch
        Me.pictureBU.Location = New System.Drawing.Point(449, 21)
        Me.pictureBU.Name = "pictureBU"
        Me.pictureBU.Size = New System.Drawing.Size(32, 32)
        Me.pictureBU.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.pictureBU.TabIndex = 27
        Me.pictureBU.TabStop = False
        '
        'addButton
        '
        Me.addButton.Image = CType(resources.GetObject("addButton.Image"), System.Drawing.Image)
        Me.addButton.InitialImage = CType(resources.GetObject("addButton.InitialImage"), System.Drawing.Image)
        Me.addButton.Location = New System.Drawing.Point(290, 148)
        Me.addButton.Name = "addButton"
        Me.addButton.Size = New System.Drawing.Size(20, 20)
        Me.addButton.TabIndex = 28
        Me.addButton.TabStop = False
        '
        'removeButton
        '
        Me.removeButton.Image = CType(resources.GetObject("removeButton.Image"), System.Drawing.Image)
        Me.removeButton.InitialImage = CType(resources.GetObject("removeButton.InitialImage"), System.Drawing.Image)
        Me.removeButton.Location = New System.Drawing.Point(289, 209)
        Me.removeButton.Name = "removeButton"
        Me.removeButton.Size = New System.Drawing.Size(20, 20)
        Me.removeButton.TabIndex = 29
        Me.removeButton.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(312, 85)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(237, 13)
        Me.Label1.TabIndex = 30
        Me.Label1.Text = "aktuelle Auswahl, auf die sich die Aktion bezieht "
        '
        'filterDropbox
        '
        Me.filterDropbox.FormattingEnabled = True
        Me.filterDropbox.Location = New System.Drawing.Point(162, 337)
        Me.filterDropbox.Name = "filterDropbox"
        Me.filterDropbox.Size = New System.Drawing.Size(264, 21)
        Me.filterDropbox.TabIndex = 31
        Me.filterDropbox.Visible = False
        '
        'filterLabel
        '
        Me.filterLabel.AutoSize = True
        Me.filterLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.filterLabel.Location = New System.Drawing.Point(12, 342)
        Me.filterLabel.Name = "filterLabel"
        Me.filterLabel.Size = New System.Drawing.Size(91, 16)
        Me.filterLabel.TabIndex = 32
        Me.filterLabel.Text = "Filter-Auswahl"
        Me.filterLabel.Visible = False
        '
        'auswSpeichern
        '
        Me.auswSpeichern.Location = New System.Drawing.Point(474, 335)
        Me.auswSpeichern.Name = "auswSpeichern"
        Me.auswSpeichern.Size = New System.Drawing.Size(113, 23)
        Me.auswSpeichern.TabIndex = 33
        Me.auswSpeichern.Text = "Speichern"
        Me.auswSpeichern.UseVisualStyleBackColor = True
        '
        'frmNameSelection
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(599, 454)
        Me.Controls.Add(Me.auswSpeichern)
        Me.Controls.Add(Me.filterLabel)
        Me.Controls.Add(Me.filterDropbox)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.removeButton)
        Me.Controls.Add(Me.addButton)
        Me.Controls.Add(Me.pictureBU)
        Me.Controls.Add(Me.rdbBU)
        Me.Controls.Add(Me.rdbTyp)
        Me.Controls.Add(Me.pictureTyp)
        Me.Controls.Add(Me.selNameListBox)
        Me.Controls.Add(Me.einstellungen)
        Me.Controls.Add(Me.pictureMilestones)
        Me.Controls.Add(Me.chkbxOneChart)
        Me.Controls.Add(Me.picturePhasen)
        Me.Controls.Add(Me.labelPPTVorlage)
        Me.Controls.Add(Me.pictureRoles)
        Me.Controls.Add(Me.pictureCosts)
        Me.Controls.Add(Me.rdbCosts)
        Me.Controls.Add(Me.statusLabel)
        Me.Controls.Add(Me.rdbRoles)
        Me.Controls.Add(Me.rdbMilestones)
        Me.Controls.Add(Me.repVorlagenDropbox)
        Me.Controls.Add(Me.rdbPhases)
        Me.Controls.Add(Me.headerLine)
        Me.Controls.Add(Me.filterBox)
        Me.Controls.Add(Me.AbbrButton)
        Me.Controls.Add(Me.OKButton)
        Me.Controls.Add(Me.nameListBox)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmNameSelection"
        Me.Text = "Visualisieren von Plan-Objekten"
        Me.TopMost = True
        CType(Me.pictureCosts, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pictureRoles, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picturePhasen, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pictureMilestones, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pictureTyp, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pictureBU, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.addButton, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.removeButton, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents nameListBox As System.Windows.Forms.ListBox
    Friend WithEvents OKButton As System.Windows.Forms.Button
    Friend WithEvents AbbrButton As System.Windows.Forms.Button
    Friend WithEvents filterBox As System.Windows.Forms.TextBox
    Friend WithEvents headerLine As System.Windows.Forms.Label
    Friend WithEvents rdbPhases As System.Windows.Forms.RadioButton
    Friend WithEvents rdbMilestones As System.Windows.Forms.RadioButton
    Friend WithEvents rdbRoles As System.Windows.Forms.RadioButton
    Friend WithEvents rdbCosts As System.Windows.Forms.RadioButton
    Friend WithEvents pictureCosts As System.Windows.Forms.PictureBox
    Friend WithEvents pictureRoles As System.Windows.Forms.PictureBox
    Friend WithEvents picturePhasen As System.Windows.Forms.PictureBox
    Friend WithEvents pictureMilestones As System.Windows.Forms.PictureBox
    Friend WithEvents chkbxOneChart As System.Windows.Forms.CheckBox
    Friend WithEvents repVorlagenDropbox As System.Windows.Forms.ComboBox
    Friend WithEvents labelPPTVorlage As System.Windows.Forms.Label
    Friend WithEvents statusLabel As System.Windows.Forms.Label
    Friend WithEvents einstellungen As System.Windows.Forms.Label
    Public WithEvents selNameListBox As System.Windows.Forms.ListBox
    Friend WithEvents pictureTyp As System.Windows.Forms.PictureBox
    Friend WithEvents rdbTyp As System.Windows.Forms.RadioButton
    Friend WithEvents rdbBU As System.Windows.Forms.RadioButton
    Friend WithEvents pictureBU As System.Windows.Forms.PictureBox
    Friend WithEvents addButton As System.Windows.Forms.PictureBox
    Friend WithEvents removeButton As System.Windows.Forms.PictureBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents filterDropbox As System.Windows.Forms.ComboBox
    Friend WithEvents filterLabel As System.Windows.Forms.Label
    Friend WithEvents auswSpeichern As System.Windows.Forms.Button
End Class
