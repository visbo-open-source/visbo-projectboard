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
        Me.rdbPhases = New System.Windows.Forms.RadioButton()
        Me.headerLine = New System.Windows.Forms.Label()
        Me.filterBox = New System.Windows.Forms.TextBox()
        Me.picturePhasen = New System.Windows.Forms.PictureBox()
        Me.rdbMilestones = New System.Windows.Forms.RadioButton()
        Me.pictureMilestones = New System.Windows.Forms.PictureBox()
        Me.rdbRoles = New System.Windows.Forms.RadioButton()
        Me.pictureRoles = New System.Windows.Forms.PictureBox()
        Me.rdbCosts = New System.Windows.Forms.RadioButton()
        Me.pictureCosts = New System.Windows.Forms.PictureBox()
        Me.rdbBU = New System.Windows.Forms.RadioButton()
        Me.pictureBU = New System.Windows.Forms.PictureBox()
        Me.rdbTyp = New System.Windows.Forms.RadioButton()
        Me.pictureTyp = New System.Windows.Forms.PictureBox()
        Me.rdbNameList = New System.Windows.Forms.RadioButton()
        Me.rdbProjStruktProj = New System.Windows.Forms.RadioButton()
        Me.rdbProjStruktTyp = New System.Windows.Forms.RadioButton()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.rdbPhaseMilest = New System.Windows.Forms.RadioButton()
        Me.picturePhaseMilest = New System.Windows.Forms.PictureBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.auswLaden = New System.Windows.Forms.Button()
        CType(Me.hryStufen, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SelectionSet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.collapseCompletely, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.expandCompletely, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SelectionReset, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picturePhasen, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pictureMilestones, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pictureRoles, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pictureCosts, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pictureBU, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pictureTyp, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        CType(Me.picturePhaseMilest, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'hryTreeView
        '
        Me.hryTreeView.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.hryTreeView.FullRowSelect = True
        Me.hryTreeView.Location = New System.Drawing.Point(12, 146)
        Me.hryTreeView.Name = "hryTreeView"
        Me.hryTreeView.Size = New System.Drawing.Size(540, 302)
        Me.hryTreeView.TabIndex = 32
        '
        'hryStufenLabel
        '
        Me.hryStufenLabel.AutoSize = True
        Me.hryStufenLabel.Location = New System.Drawing.Point(202, 45)
        Me.hryStufenLabel.Name = "hryStufenLabel"
        Me.hryStufenLabel.Size = New System.Drawing.Size(264, 13)
        Me.hryStufenLabel.TabIndex = 35
        Me.hryStufenLabel.Text = "Anzahl Parents in der Projekt-Struktur berücksichtigen:"
        Me.hryStufenLabel.Visible = False
        '
        'hryStufen
        '
        Me.hryStufen.Location = New System.Drawing.Point(489, 43)
        Me.hryStufen.Name = "hryStufen"
        Me.hryStufen.Size = New System.Drawing.Size(57, 20)
        Me.hryStufen.TabIndex = 34
        Me.hryStufen.Value = New Decimal(New Integer() {10, 0, 0, 0})
        Me.hryStufen.Visible = False
        '
        'einstellungen
        '
        Me.einstellungen.AutoSize = True
        Me.einstellungen.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, CType((System.Drawing.FontStyle.Italic Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.einstellungen.ForeColor = System.Drawing.SystemColors.MenuHighlight
        Me.einstellungen.Location = New System.Drawing.Point(460, 550)
        Me.einstellungen.Name = "einstellungen"
        Me.einstellungen.Size = New System.Drawing.Size(70, 13)
        Me.einstellungen.TabIndex = 42
        Me.einstellungen.Text = "Einstellungen"
        Me.einstellungen.Visible = False
        '
        'chkbxOneChart
        '
        Me.chkbxOneChart.AutoSize = True
        Me.chkbxOneChart.Location = New System.Drawing.Point(428, 454)
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
        Me.labelPPTVorlage.Location = New System.Drawing.Point(12, 525)
        Me.labelPPTVorlage.Name = "labelPPTVorlage"
        Me.labelPPTVorlage.Size = New System.Drawing.Size(136, 16)
        Me.labelPPTVorlage.TabIndex = 39
        Me.labelPPTVorlage.Text = "Powerpoint Template"
        Me.labelPPTVorlage.Visible = False
        '
        'statusLabel
        '
        Me.statusLabel.AutoSize = True
        Me.statusLabel.Location = New System.Drawing.Point(9, 573)
        Me.statusLabel.Name = "statusLabel"
        Me.statusLabel.Size = New System.Drawing.Size(39, 13)
        Me.statusLabel.TabIndex = 41
        Me.statusLabel.Text = "Label1"
        '
        'repVorlagenDropbox
        '
        Me.repVorlagenDropbox.FormattingEnabled = True
        Me.repVorlagenDropbox.Location = New System.Drawing.Point(155, 520)
        Me.repVorlagenDropbox.Name = "repVorlagenDropbox"
        Me.repVorlagenDropbox.Size = New System.Drawing.Size(264, 21)
        Me.repVorlagenDropbox.TabIndex = 37
        Me.repVorlagenDropbox.Visible = False
        '
        'OKButton
        '
        Me.OKButton.Location = New System.Drawing.Point(155, 545)
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
        Me.AbbrButton.Location = New System.Drawing.Point(306, 545)
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
        Me.filterDropbox.Location = New System.Drawing.Point(155, 493)
        Me.filterDropbox.MaxDropDownItems = 10
        Me.filterDropbox.Name = "filterDropbox"
        Me.filterDropbox.Size = New System.Drawing.Size(264, 21)
        Me.filterDropbox.Sorted = True
        Me.filterDropbox.TabIndex = 48
        Me.filterDropbox.Visible = False
        '
        'filterLabel
        '
        Me.filterLabel.AutoSize = True
        Me.filterLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.filterLabel.Location = New System.Drawing.Point(12, 494)
        Me.filterLabel.Name = "filterLabel"
        Me.filterLabel.Size = New System.Drawing.Size(37, 16)
        Me.filterLabel.TabIndex = 49
        Me.filterLabel.Text = "Filter"
        Me.filterLabel.Visible = False
        '
        'auswSpeichern
        '
        Me.auswSpeichern.Location = New System.Drawing.Point(439, 492)
        Me.auswSpeichern.Name = "auswSpeichern"
        Me.auswSpeichern.Size = New System.Drawing.Size(113, 21)
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
        Me.SelectionSet.Location = New System.Drawing.Point(12, 455)
        Me.SelectionSet.Name = "SelectionSet"
        Me.SelectionSet.Size = New System.Drawing.Size(16, 16)
        Me.SelectionSet.TabIndex = 51
        Me.SelectionSet.TabStop = False
        '
        'collapseCompletely
        '
        Me.collapseCompletely.BackColor = System.Drawing.SystemColors.Control
        Me.collapseCompletely.Image = CType(resources.GetObject("collapseCompletely.Image"), System.Drawing.Image)
        Me.collapseCompletely.Location = New System.Drawing.Point(67, 455)
        Me.collapseCompletely.Name = "collapseCompletely"
        Me.collapseCompletely.Size = New System.Drawing.Size(16, 16)
        Me.collapseCompletely.TabIndex = 47
        Me.collapseCompletely.TabStop = False
        '
        'expandCompletely
        '
        Me.expandCompletely.BackColor = System.Drawing.SystemColors.Control
        Me.expandCompletely.Image = CType(resources.GetObject("expandCompletely.Image"), System.Drawing.Image)
        Me.expandCompletely.Location = New System.Drawing.Point(89, 455)
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
        Me.SelectionReset.Location = New System.Drawing.Point(34, 455)
        Me.SelectionReset.Name = "SelectionReset"
        Me.SelectionReset.Size = New System.Drawing.Size(16, 16)
        Me.SelectionReset.TabIndex = 45
        Me.SelectionReset.TabStop = False
        '
        'BackgroundWorker3
        '
        Me.BackgroundWorker3.WorkerReportsProgress = True
        Me.BackgroundWorker3.WorkerSupportsCancellation = True
        '
        'rdbPhases
        '
        Me.rdbPhases.AutoSize = True
        Me.rdbPhases.Checked = True
        Me.rdbPhases.Location = New System.Drawing.Point(6, 16)
        Me.rdbPhases.Name = "rdbPhases"
        Me.rdbPhases.Size = New System.Drawing.Size(14, 13)
        Me.rdbPhases.TabIndex = 52
        Me.rdbPhases.TabStop = True
        Me.rdbPhases.UseVisualStyleBackColor = True
        '
        'headerLine
        '
        Me.headerLine.AutoSize = True
        Me.headerLine.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.headerLine.Location = New System.Drawing.Point(9, 121)
        Me.headerLine.Name = "headerLine"
        Me.headerLine.Size = New System.Drawing.Size(91, 16)
        Me.headerLine.TabIndex = 53
        Me.headerLine.Text = "Label1              "
        '
        'filterBox
        '
        Me.filterBox.Enabled = False
        Me.filterBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.filterBox.HideSelection = False
        Me.filterBox.Location = New System.Drawing.Point(125, 118)
        Me.filterBox.Name = "filterBox"
        Me.filterBox.Size = New System.Drawing.Size(150, 22)
        Me.filterBox.TabIndex = 54
        Me.filterBox.Visible = False
        '
        'picturePhasen
        '
        Me.picturePhasen.Image = CType(resources.GetObject("picturePhasen.Image"), System.Drawing.Image)
        Me.picturePhasen.Location = New System.Drawing.Point(26, 3)
        Me.picturePhasen.Name = "picturePhasen"
        Me.picturePhasen.Size = New System.Drawing.Size(33, 33)
        Me.picturePhasen.TabIndex = 55
        Me.picturePhasen.TabStop = False
        '
        'rdbMilestones
        '
        Me.rdbMilestones.AutoSize = True
        Me.rdbMilestones.Location = New System.Drawing.Point(93, 16)
        Me.rdbMilestones.Name = "rdbMilestones"
        Me.rdbMilestones.Size = New System.Drawing.Size(14, 13)
        Me.rdbMilestones.TabIndex = 56
        Me.rdbMilestones.UseVisualStyleBackColor = True
        '
        'pictureMilestones
        '
        Me.pictureMilestones.Image = CType(resources.GetObject("pictureMilestones.Image"), System.Drawing.Image)
        Me.pictureMilestones.Location = New System.Drawing.Point(113, 3)
        Me.pictureMilestones.Name = "pictureMilestones"
        Me.pictureMilestones.Size = New System.Drawing.Size(33, 33)
        Me.pictureMilestones.TabIndex = 57
        Me.pictureMilestones.TabStop = False
        '
        'rdbRoles
        '
        Me.rdbRoles.AutoSize = True
        Me.rdbRoles.Location = New System.Drawing.Point(247, 16)
        Me.rdbRoles.Name = "rdbRoles"
        Me.rdbRoles.Size = New System.Drawing.Size(14, 13)
        Me.rdbRoles.TabIndex = 58
        Me.rdbRoles.UseVisualStyleBackColor = True
        '
        'pictureRoles
        '
        Me.pictureRoles.Image = CType(resources.GetObject("pictureRoles.Image"), System.Drawing.Image)
        Me.pictureRoles.Location = New System.Drawing.Point(267, 3)
        Me.pictureRoles.Name = "pictureRoles"
        Me.pictureRoles.Size = New System.Drawing.Size(33, 33)
        Me.pictureRoles.TabIndex = 59
        Me.pictureRoles.TabStop = False
        '
        'rdbCosts
        '
        Me.rdbCosts.AutoSize = True
        Me.rdbCosts.Location = New System.Drawing.Point(319, 16)
        Me.rdbCosts.Name = "rdbCosts"
        Me.rdbCosts.Size = New System.Drawing.Size(14, 13)
        Me.rdbCosts.TabIndex = 60
        Me.rdbCosts.UseVisualStyleBackColor = True
        '
        'pictureCosts
        '
        Me.pictureCosts.Image = CType(resources.GetObject("pictureCosts.Image"), System.Drawing.Image)
        Me.pictureCosts.Location = New System.Drawing.Point(339, 3)
        Me.pictureCosts.Name = "pictureCosts"
        Me.pictureCosts.Size = New System.Drawing.Size(33, 33)
        Me.pictureCosts.TabIndex = 61
        Me.pictureCosts.TabStop = False
        '
        'rdbBU
        '
        Me.rdbBU.AutoSize = True
        Me.rdbBU.Location = New System.Drawing.Point(397, 16)
        Me.rdbBU.Name = "rdbBU"
        Me.rdbBU.Size = New System.Drawing.Size(14, 13)
        Me.rdbBU.TabIndex = 62
        Me.rdbBU.UseVisualStyleBackColor = True
        Me.rdbBU.Visible = False
        '
        'pictureBU
        '
        Me.pictureBU.Image = Global.ProjectBoardBasic.My.Resources.Resources.branch
        Me.pictureBU.Location = New System.Drawing.Point(416, 3)
        Me.pictureBU.Name = "pictureBU"
        Me.pictureBU.Size = New System.Drawing.Size(33, 33)
        Me.pictureBU.TabIndex = 63
        Me.pictureBU.TabStop = False
        Me.pictureBU.Visible = False
        '
        'rdbTyp
        '
        Me.rdbTyp.AutoSize = True
        Me.rdbTyp.Location = New System.Drawing.Point(481, 16)
        Me.rdbTyp.Name = "rdbTyp"
        Me.rdbTyp.Size = New System.Drawing.Size(14, 13)
        Me.rdbTyp.TabIndex = 64
        Me.rdbTyp.UseVisualStyleBackColor = True
        Me.rdbTyp.Visible = False
        '
        'pictureTyp
        '
        Me.pictureTyp.Image = CType(resources.GetObject("pictureTyp.Image"), System.Drawing.Image)
        Me.pictureTyp.Location = New System.Drawing.Point(501, 3)
        Me.pictureTyp.Name = "pictureTyp"
        Me.pictureTyp.Size = New System.Drawing.Size(33, 33)
        Me.pictureTyp.TabIndex = 65
        Me.pictureTyp.TabStop = False
        Me.pictureTyp.Visible = False
        '
        'rdbNameList
        '
        Me.rdbNameList.AutoSize = True
        Me.rdbNameList.Checked = True
        Me.rdbNameList.Location = New System.Drawing.Point(6, 5)
        Me.rdbNameList.Name = "rdbNameList"
        Me.rdbNameList.Size = New System.Drawing.Size(47, 17)
        Me.rdbNameList.TabIndex = 66
        Me.rdbNameList.TabStop = True
        Me.rdbNameList.Text = "Liste"
        Me.rdbNameList.UseVisualStyleBackColor = True
        '
        'rdbProjStruktProj
        '
        Me.rdbProjStruktProj.AutoSize = True
        Me.rdbProjStruktProj.Location = New System.Drawing.Point(394, 5)
        Me.rdbProjStruktProj.Name = "rdbProjStruktProj"
        Me.rdbProjStruktProj.Size = New System.Drawing.Size(140, 17)
        Me.rdbProjStruktProj.TabIndex = 67
        Me.rdbProjStruktProj.Text = "Projekt-Struktur (Projekt)"
        Me.rdbProjStruktProj.UseVisualStyleBackColor = True
        '
        'rdbProjStruktTyp
        '
        Me.rdbProjStruktTyp.AutoSize = True
        Me.rdbProjStruktTyp.Location = New System.Drawing.Point(166, 5)
        Me.rdbProjStruktTyp.Name = "rdbProjStruktTyp"
        Me.rdbProjStruktTyp.Size = New System.Drawing.Size(125, 17)
        Me.rdbProjStruktTyp.TabIndex = 68
        Me.rdbProjStruktTyp.Text = "Projekt-Struktur (Typ)"
        Me.rdbProjStruktTyp.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.rdbPhaseMilest)
        Me.Panel1.Controls.Add(Me.picturePhaseMilest)
        Me.Panel1.Controls.Add(Me.rdbPhases)
        Me.Panel1.Controls.Add(Me.picturePhasen)
        Me.Panel1.Controls.Add(Me.pictureTyp)
        Me.Panel1.Controls.Add(Me.rdbMilestones)
        Me.Panel1.Controls.Add(Me.rdbTyp)
        Me.Panel1.Controls.Add(Me.pictureMilestones)
        Me.Panel1.Controls.Add(Me.pictureBU)
        Me.Panel1.Controls.Add(Me.rdbRoles)
        Me.Panel1.Controls.Add(Me.rdbBU)
        Me.Panel1.Controls.Add(Me.pictureRoles)
        Me.Panel1.Controls.Add(Me.pictureCosts)
        Me.Panel1.Controls.Add(Me.rdbCosts)
        Me.Panel1.Location = New System.Drawing.Point(12, 69)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(540, 41)
        Me.Panel1.TabIndex = 70
        '
        'rdbPhaseMilest
        '
        Me.rdbPhaseMilest.AutoSize = True
        Me.rdbPhaseMilest.Location = New System.Drawing.Point(173, 16)
        Me.rdbPhaseMilest.Name = "rdbPhaseMilest"
        Me.rdbPhaseMilest.Size = New System.Drawing.Size(14, 13)
        Me.rdbPhaseMilest.TabIndex = 67
        Me.rdbPhaseMilest.UseVisualStyleBackColor = True
        Me.rdbPhaseMilest.Visible = False
        '
        'picturePhaseMilest
        '
        Me.picturePhaseMilest.Image = Global.ProjectBoardBasic.My.Resources.Resources.gear
        Me.picturePhaseMilest.Location = New System.Drawing.Point(193, 5)
        Me.picturePhaseMilest.Name = "picturePhaseMilest"
        Me.picturePhaseMilest.Size = New System.Drawing.Size(33, 33)
        Me.picturePhaseMilest.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picturePhaseMilest.TabIndex = 66
        Me.picturePhaseMilest.TabStop = False
        Me.picturePhaseMilest.Visible = False
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.rdbProjStruktTyp)
        Me.Panel2.Controls.Add(Me.rdbNameList)
        Me.Panel2.Controls.Add(Me.rdbProjStruktProj)
        Me.Panel2.Location = New System.Drawing.Point(12, 12)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(542, 25)
        Me.Panel2.TabIndex = 71
        '
        'auswLaden
        '
        Me.auswLaden.Location = New System.Drawing.Point(439, 519)
        Me.auswLaden.Name = "auswLaden"
        Me.auswLaden.Size = New System.Drawing.Size(113, 21)
        Me.auswLaden.TabIndex = 72
        Me.auswLaden.Text = "Laden"
        Me.auswLaden.UseVisualStyleBackColor = True
        '
        'frmHierarchySelection
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(558, 595)
        Me.Controls.Add(Me.auswLaden)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.filterBox)
        Me.Controls.Add(Me.headerLine)
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
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmHierarchySelection"
        Me.Text = "Auswahl von Plan-Objekten"
        Me.TopMost = True
        CType(Me.hryStufen, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SelectionSet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.collapseCompletely, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.expandCompletely, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SelectionReset, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picturePhasen, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pictureMilestones, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pictureRoles, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pictureCosts, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pictureBU, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pictureTyp, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.picturePhaseMilest, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
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
    Friend WithEvents rdbPhases As System.Windows.Forms.RadioButton
    Friend WithEvents headerLine As System.Windows.Forms.Label
    Friend WithEvents filterBox As System.Windows.Forms.TextBox
    Friend WithEvents picturePhasen As System.Windows.Forms.PictureBox
    Friend WithEvents rdbMilestones As System.Windows.Forms.RadioButton
    Friend WithEvents pictureMilestones As System.Windows.Forms.PictureBox
    Friend WithEvents rdbRoles As System.Windows.Forms.RadioButton
    Friend WithEvents pictureRoles As System.Windows.Forms.PictureBox
    Friend WithEvents rdbCosts As System.Windows.Forms.RadioButton
    Friend WithEvents pictureCosts As System.Windows.Forms.PictureBox
    Friend WithEvents rdbBU As System.Windows.Forms.RadioButton
    Friend WithEvents pictureBU As System.Windows.Forms.PictureBox
    Friend WithEvents rdbTyp As System.Windows.Forms.RadioButton
    Friend WithEvents pictureTyp As System.Windows.Forms.PictureBox
    Friend WithEvents rdbNameList As System.Windows.Forms.RadioButton
    Friend WithEvents rdbProjStruktProj As System.Windows.Forms.RadioButton
    Friend WithEvents rdbProjStruktTyp As System.Windows.Forms.RadioButton
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents auswLaden As System.Windows.Forms.Button
    Friend WithEvents rdbPhaseMilest As System.Windows.Forms.RadioButton
    Friend WithEvents picturePhaseMilest As System.Windows.Forms.PictureBox
End Class
