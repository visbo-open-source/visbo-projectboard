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
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.BackgroundWorker3 = New System.ComponentModel.BackgroundWorker()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.auswLaden = New System.Windows.Forms.Button()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.rdbPhaseMilest = New System.Windows.Forms.RadioButton()
        Me.picturePhaseMilest = New System.Windows.Forms.PictureBox()
        Me.rdbPhases = New System.Windows.Forms.RadioButton()
        Me.picturePhasen = New System.Windows.Forms.PictureBox()
        Me.pictureTyp = New System.Windows.Forms.PictureBox()
        Me.rdbMilestones = New System.Windows.Forms.RadioButton()
        Me.rdbTyp = New System.Windows.Forms.RadioButton()
        Me.pictureMilestones = New System.Windows.Forms.PictureBox()
        Me.pictureBU = New System.Windows.Forms.PictureBox()
        Me.rdbRoles = New System.Windows.Forms.RadioButton()
        Me.rdbBU = New System.Windows.Forms.RadioButton()
        Me.pictureRoles = New System.Windows.Forms.PictureBox()
        Me.pictureCosts = New System.Windows.Forms.PictureBox()
        Me.rdbCosts = New System.Windows.Forms.RadioButton()
        Me.filterBox = New System.Windows.Forms.TextBox()
        Me.headerLine = New System.Windows.Forms.Label()
        Me.SelectionSet = New System.Windows.Forms.PictureBox()
        Me.auswSpeichern = New System.Windows.Forms.Button()
        Me.filterLabel = New System.Windows.Forms.Label()
        Me.filterDropbox = New System.Windows.Forms.ComboBox()
        Me.collapseCompletely = New System.Windows.Forms.PictureBox()
        Me.expandCompletely = New System.Windows.Forms.PictureBox()
        Me.SelectionReset = New System.Windows.Forms.PictureBox()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.rdbProjStruktTyp = New System.Windows.Forms.RadioButton()
        Me.rdbNameList = New System.Windows.Forms.RadioButton()
        Me.rdbProjStruktProj = New System.Windows.Forms.RadioButton()
        Me.AbbrButton = New System.Windows.Forms.Button()
        Me.chkbxOneChart = New System.Windows.Forms.CheckBox()
        Me.statusLabel = New System.Windows.Forms.Label()
        Me.repVorlagenDropbox = New System.Windows.Forms.ComboBox()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.labelPPTVorlage = New System.Windows.Forms.Label()
        Me.einstellungen = New System.Windows.Forms.Label()
        Me.hryTreeView = New System.Windows.Forms.TreeView()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.picturePhaseMilest, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picturePhasen, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pictureTyp, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pictureMilestones, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pictureBU, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pictureRoles, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pictureCosts, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SelectionSet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.collapseCompletely, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.expandCompletely, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SelectionReset, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'BackgroundWorker1
        '
        Me.BackgroundWorker1.WorkerReportsProgress = True
        Me.BackgroundWorker1.WorkerSupportsCancellation = True
        '
        'BackgroundWorker3
        '
        Me.BackgroundWorker3.WorkerReportsProgress = True
        Me.BackgroundWorker3.WorkerSupportsCancellation = True
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.Controls.Add(Me.auswLaden)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Controls.Add(Me.filterBox)
        Me.Panel1.Controls.Add(Me.headerLine)
        Me.Panel1.Controls.Add(Me.SelectionSet)
        Me.Panel1.Controls.Add(Me.auswSpeichern)
        Me.Panel1.Controls.Add(Me.filterLabel)
        Me.Panel1.Controls.Add(Me.filterDropbox)
        Me.Panel1.Controls.Add(Me.collapseCompletely)
        Me.Panel1.Controls.Add(Me.expandCompletely)
        Me.Panel1.Controls.Add(Me.SelectionReset)
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Controls.Add(Me.AbbrButton)
        Me.Panel1.Controls.Add(Me.chkbxOneChart)
        Me.Panel1.Controls.Add(Me.statusLabel)
        Me.Panel1.Controls.Add(Me.repVorlagenDropbox)
        Me.Panel1.Controls.Add(Me.OKButton)
        Me.Panel1.Controls.Add(Me.labelPPTVorlage)
        Me.Panel1.Controls.Add(Me.einstellungen)
        Me.Panel1.Controls.Add(Me.hryTreeView)
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(457, 463)
        Me.Panel1.TabIndex = 0
        '
        'auswLaden
        '
        Me.auswLaden.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.auswLaden.Location = New System.Drawing.Point(325, 378)
        Me.auswLaden.Name = "auswLaden"
        Me.auswLaden.Size = New System.Drawing.Size(113, 21)
        Me.auswLaden.TabIndex = 94
        Me.auswLaden.Text = "Laden"
        Me.auswLaden.UseVisualStyleBackColor = True
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.Controls.Add(Me.rdbPhaseMilest)
        Me.Panel2.Controls.Add(Me.picturePhaseMilest)
        Me.Panel2.Controls.Add(Me.rdbPhases)
        Me.Panel2.Controls.Add(Me.picturePhasen)
        Me.Panel2.Controls.Add(Me.pictureTyp)
        Me.Panel2.Controls.Add(Me.rdbMilestones)
        Me.Panel2.Controls.Add(Me.rdbTyp)
        Me.Panel2.Controls.Add(Me.pictureMilestones)
        Me.Panel2.Controls.Add(Me.pictureBU)
        Me.Panel2.Controls.Add(Me.rdbRoles)
        Me.Panel2.Controls.Add(Me.rdbBU)
        Me.Panel2.Controls.Add(Me.pictureRoles)
        Me.Panel2.Controls.Add(Me.pictureCosts)
        Me.Panel2.Controls.Add(Me.rdbCosts)
        Me.Panel2.Location = New System.Drawing.Point(9, 10)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(440, 41)
        Me.Panel2.TabIndex = 92
        '
        'rdbPhaseMilest
        '
        Me.rdbPhaseMilest.AutoSize = True
        Me.rdbPhaseMilest.Location = New System.Drawing.Point(6, 16)
        Me.rdbPhaseMilest.Name = "rdbPhaseMilest"
        Me.rdbPhaseMilest.Size = New System.Drawing.Size(14, 13)
        Me.rdbPhaseMilest.TabIndex = 67
        Me.rdbPhaseMilest.UseVisualStyleBackColor = True
        Me.rdbPhaseMilest.Visible = False
        '
        'picturePhaseMilest
        '
        Me.picturePhaseMilest.Image = Global.ProjectBoardBasic.My.Resources.Resources.phases_und_milestones_248x248
        Me.picturePhaseMilest.Location = New System.Drawing.Point(26, 3)
        Me.picturePhaseMilest.Name = "picturePhaseMilest"
        Me.picturePhaseMilest.Size = New System.Drawing.Size(33, 33)
        Me.picturePhaseMilest.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picturePhaseMilest.TabIndex = 66
        Me.picturePhaseMilest.TabStop = False
        Me.picturePhaseMilest.Visible = False
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
        'picturePhasen
        '
        Me.picturePhasen.Image = CType(resources.GetObject("picturePhasen.Image"), System.Drawing.Image)
        Me.picturePhasen.Location = New System.Drawing.Point(26, 3)
        Me.picturePhasen.Name = "picturePhasen"
        Me.picturePhasen.Size = New System.Drawing.Size(33, 33)
        Me.picturePhasen.TabIndex = 55
        Me.picturePhasen.TabStop = False
        '
        'pictureTyp
        '
        Me.pictureTyp.Image = CType(resources.GetObject("pictureTyp.Image"), System.Drawing.Image)
        Me.pictureTyp.Location = New System.Drawing.Point(365, 3)
        Me.pictureTyp.Name = "pictureTyp"
        Me.pictureTyp.Size = New System.Drawing.Size(33, 33)
        Me.pictureTyp.TabIndex = 65
        Me.pictureTyp.TabStop = False
        Me.pictureTyp.Visible = False
        '
        'rdbMilestones
        '
        Me.rdbMilestones.AutoSize = True
        Me.rdbMilestones.Location = New System.Drawing.Point(74, 16)
        Me.rdbMilestones.Name = "rdbMilestones"
        Me.rdbMilestones.Size = New System.Drawing.Size(14, 13)
        Me.rdbMilestones.TabIndex = 56
        Me.rdbMilestones.UseVisualStyleBackColor = True
        '
        'rdbTyp
        '
        Me.rdbTyp.AutoSize = True
        Me.rdbTyp.Location = New System.Drawing.Point(346, 16)
        Me.rdbTyp.Name = "rdbTyp"
        Me.rdbTyp.Size = New System.Drawing.Size(14, 13)
        Me.rdbTyp.TabIndex = 64
        Me.rdbTyp.UseVisualStyleBackColor = True
        Me.rdbTyp.Visible = False
        '
        'pictureMilestones
        '
        Me.pictureMilestones.Image = Global.ProjectBoardBasic.My.Resources.Resources.milestones_icon1
        Me.pictureMilestones.Location = New System.Drawing.Point(94, 3)
        Me.pictureMilestones.Name = "pictureMilestones"
        Me.pictureMilestones.Size = New System.Drawing.Size(33, 33)
        Me.pictureMilestones.TabIndex = 57
        Me.pictureMilestones.TabStop = False
        '
        'pictureBU
        '
        Me.pictureBU.Image = Global.ProjectBoardBasic.My.Resources.Resources.branch
        Me.pictureBU.Location = New System.Drawing.Point(293, 3)
        Me.pictureBU.Name = "pictureBU"
        Me.pictureBU.Size = New System.Drawing.Size(33, 33)
        Me.pictureBU.TabIndex = 63
        Me.pictureBU.TabStop = False
        Me.pictureBU.Visible = False
        '
        'rdbRoles
        '
        Me.rdbRoles.AutoSize = True
        Me.rdbRoles.Location = New System.Drawing.Point(140, 16)
        Me.rdbRoles.Name = "rdbRoles"
        Me.rdbRoles.Size = New System.Drawing.Size(14, 13)
        Me.rdbRoles.TabIndex = 58
        Me.rdbRoles.UseVisualStyleBackColor = True
        '
        'rdbBU
        '
        Me.rdbBU.AutoSize = True
        Me.rdbBU.Location = New System.Drawing.Point(274, 16)
        Me.rdbBU.Name = "rdbBU"
        Me.rdbBU.Size = New System.Drawing.Size(14, 13)
        Me.rdbBU.TabIndex = 62
        Me.rdbBU.UseVisualStyleBackColor = True
        Me.rdbBU.Visible = False
        '
        'pictureRoles
        '
        Me.pictureRoles.Image = CType(resources.GetObject("pictureRoles.Image"), System.Drawing.Image)
        Me.pictureRoles.Location = New System.Drawing.Point(159, 3)
        Me.pictureRoles.Name = "pictureRoles"
        Me.pictureRoles.Size = New System.Drawing.Size(33, 33)
        Me.pictureRoles.TabIndex = 59
        Me.pictureRoles.TabStop = False
        '
        'pictureCosts
        '
        Me.pictureCosts.Image = CType(resources.GetObject("pictureCosts.Image"), System.Drawing.Image)
        Me.pictureCosts.Location = New System.Drawing.Point(227, 3)
        Me.pictureCosts.Name = "pictureCosts"
        Me.pictureCosts.Size = New System.Drawing.Size(33, 33)
        Me.pictureCosts.TabIndex = 61
        Me.pictureCosts.TabStop = False
        '
        'rdbCosts
        '
        Me.rdbCosts.AutoSize = True
        Me.rdbCosts.Location = New System.Drawing.Point(206, 16)
        Me.rdbCosts.Name = "rdbCosts"
        Me.rdbCosts.Size = New System.Drawing.Size(14, 13)
        Me.rdbCosts.TabIndex = 60
        Me.rdbCosts.UseVisualStyleBackColor = True
        '
        'filterBox
        '
        Me.filterBox.Enabled = False
        Me.filterBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.filterBox.HideSelection = False
        Me.filterBox.Location = New System.Drawing.Point(122, 93)
        Me.filterBox.Name = "filterBox"
        Me.filterBox.Size = New System.Drawing.Size(150, 22)
        Me.filterBox.TabIndex = 91
        Me.filterBox.Visible = False
        '
        'headerLine
        '
        Me.headerLine.AutoSize = True
        Me.headerLine.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.headerLine.Location = New System.Drawing.Point(10, 95)
        Me.headerLine.Name = "headerLine"
        Me.headerLine.Size = New System.Drawing.Size(91, 16)
        Me.headerLine.TabIndex = 90
        Me.headerLine.Text = "Label1              "
        '
        'SelectionSet
        '
        Me.SelectionSet.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.SelectionSet.BackColor = System.Drawing.SystemColors.Control
        Me.SelectionSet.ErrorImage = CType(resources.GetObject("SelectionSet.ErrorImage"), System.Drawing.Image)
        Me.SelectionSet.Image = CType(resources.GetObject("SelectionSet.Image"), System.Drawing.Image)
        Me.SelectionSet.InitialImage = Nothing
        Me.SelectionSet.Location = New System.Drawing.Point(9, 322)
        Me.SelectionSet.Name = "SelectionSet"
        Me.SelectionSet.Size = New System.Drawing.Size(16, 16)
        Me.SelectionSet.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.SelectionSet.TabIndex = 89
        Me.SelectionSet.TabStop = False
        '
        'auswSpeichern
        '
        Me.auswSpeichern.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.auswSpeichern.Location = New System.Drawing.Point(325, 351)
        Me.auswSpeichern.Name = "auswSpeichern"
        Me.auswSpeichern.Size = New System.Drawing.Size(113, 21)
        Me.auswSpeichern.TabIndex = 88
        Me.auswSpeichern.Text = "Speichern"
        Me.auswSpeichern.UseVisualStyleBackColor = True
        '
        'filterLabel
        '
        Me.filterLabel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.filterLabel.AutoSize = True
        Me.filterLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.filterLabel.Location = New System.Drawing.Point(9, 353)
        Me.filterLabel.Name = "filterLabel"
        Me.filterLabel.Size = New System.Drawing.Size(37, 16)
        Me.filterLabel.TabIndex = 87
        Me.filterLabel.Text = "Filter"
        Me.filterLabel.Visible = False
        '
        'filterDropbox
        '
        Me.filterDropbox.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.filterDropbox.FormattingEnabled = True
        Me.filterDropbox.Location = New System.Drawing.Point(152, 352)
        Me.filterDropbox.MaxDropDownItems = 10
        Me.filterDropbox.Name = "filterDropbox"
        Me.filterDropbox.Size = New System.Drawing.Size(153, 21)
        Me.filterDropbox.Sorted = True
        Me.filterDropbox.TabIndex = 86
        Me.filterDropbox.Visible = False
        '
        'collapseCompletely
        '
        Me.collapseCompletely.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.collapseCompletely.BackColor = System.Drawing.SystemColors.Control
        Me.collapseCompletely.Image = CType(resources.GetObject("collapseCompletely.Image"), System.Drawing.Image)
        Me.collapseCompletely.Location = New System.Drawing.Point(64, 322)
        Me.collapseCompletely.Name = "collapseCompletely"
        Me.collapseCompletely.Size = New System.Drawing.Size(16, 16)
        Me.collapseCompletely.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.collapseCompletely.TabIndex = 85
        Me.collapseCompletely.TabStop = False
        '
        'expandCompletely
        '
        Me.expandCompletely.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.expandCompletely.BackColor = System.Drawing.SystemColors.Control
        Me.expandCompletely.Image = CType(resources.GetObject("expandCompletely.Image"), System.Drawing.Image)
        Me.expandCompletely.Location = New System.Drawing.Point(86, 322)
        Me.expandCompletely.Name = "expandCompletely"
        Me.expandCompletely.Size = New System.Drawing.Size(16, 16)
        Me.expandCompletely.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.expandCompletely.TabIndex = 84
        Me.expandCompletely.TabStop = False
        '
        'SelectionReset
        '
        Me.SelectionReset.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.SelectionReset.BackColor = System.Drawing.SystemColors.Control
        Me.SelectionReset.Image = CType(resources.GetObject("SelectionReset.Image"), System.Drawing.Image)
        Me.SelectionReset.InitialImage = Nothing
        Me.SelectionReset.Location = New System.Drawing.Point(31, 322)
        Me.SelectionReset.Name = "SelectionReset"
        Me.SelectionReset.Size = New System.Drawing.Size(16, 16)
        Me.SelectionReset.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.SelectionReset.TabIndex = 83
        Me.SelectionReset.TabStop = False
        '
        'Panel3
        '
        Me.Panel3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel3.Controls.Add(Me.rdbProjStruktTyp)
        Me.Panel3.Controls.Add(Me.rdbNameList)
        Me.Panel3.Controls.Add(Me.rdbProjStruktProj)
        Me.Panel3.Location = New System.Drawing.Point(9, 57)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(442, 25)
        Me.Panel3.TabIndex = 93
        '
        'rdbProjStruktTyp
        '
        Me.rdbProjStruktTyp.AutoSize = True
        Me.rdbProjStruktTyp.Location = New System.Drawing.Point(74, 5)
        Me.rdbProjStruktTyp.Name = "rdbProjStruktTyp"
        Me.rdbProjStruktTyp.Size = New System.Drawing.Size(125, 17)
        Me.rdbProjStruktTyp.TabIndex = 68
        Me.rdbProjStruktTyp.Text = "Projekt-Struktur (Typ)"
        Me.rdbProjStruktTyp.UseVisualStyleBackColor = True
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
        Me.rdbProjStruktProj.Location = New System.Drawing.Point(206, 3)
        Me.rdbProjStruktProj.Name = "rdbProjStruktProj"
        Me.rdbProjStruktProj.Size = New System.Drawing.Size(140, 17)
        Me.rdbProjStruktProj.TabIndex = 67
        Me.rdbProjStruktProj.Text = "Projekt-Struktur (Projekt)"
        Me.rdbProjStruktProj.UseVisualStyleBackColor = True
        '
        'AbbrButton
        '
        Me.AbbrButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.AbbrButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.AbbrButton.Location = New System.Drawing.Point(257, 404)
        Me.AbbrButton.Name = "AbbrButton"
        Me.AbbrButton.Size = New System.Drawing.Size(88, 23)
        Me.AbbrButton.TabIndex = 82
        Me.AbbrButton.Text = "Abbrechen"
        Me.AbbrButton.UseVisualStyleBackColor = False
        Me.AbbrButton.Visible = False
        '
        'chkbxOneChart
        '
        Me.chkbxOneChart.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkbxOneChart.AutoSize = True
        Me.chkbxOneChart.Location = New System.Drawing.Point(322, 321)
        Me.chkbxOneChart.Name = "chkbxOneChart"
        Me.chkbxOneChart.Size = New System.Drawing.Size(118, 17)
        Me.chkbxOneChart.TabIndex = 76
        Me.chkbxOneChart.Text = "Alles in einem Chart"
        Me.chkbxOneChart.UseVisualStyleBackColor = True
        Me.chkbxOneChart.Visible = False
        '
        'statusLabel
        '
        Me.statusLabel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.statusLabel.AutoSize = True
        Me.statusLabel.Location = New System.Drawing.Point(6, 431)
        Me.statusLabel.Name = "statusLabel"
        Me.statusLabel.Size = New System.Drawing.Size(39, 13)
        Me.statusLabel.TabIndex = 80
        Me.statusLabel.Text = "Label1"
        '
        'repVorlagenDropbox
        '
        Me.repVorlagenDropbox.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.repVorlagenDropbox.FormattingEnabled = True
        Me.repVorlagenDropbox.Location = New System.Drawing.Point(152, 379)
        Me.repVorlagenDropbox.Name = "repVorlagenDropbox"
        Me.repVorlagenDropbox.Size = New System.Drawing.Size(153, 21)
        Me.repVorlagenDropbox.TabIndex = 77
        Me.repVorlagenDropbox.Visible = False
        '
        'OKButton
        '
        Me.OKButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.OKButton.Location = New System.Drawing.Point(152, 404)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(116, 23)
        Me.OKButton.TabIndex = 78
        Me.OKButton.Text = "Anzeigen"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'labelPPTVorlage
        '
        Me.labelPPTVorlage.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.labelPPTVorlage.AutoSize = True
        Me.labelPPTVorlage.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelPPTVorlage.Location = New System.Drawing.Point(9, 384)
        Me.labelPPTVorlage.Name = "labelPPTVorlage"
        Me.labelPPTVorlage.Size = New System.Drawing.Size(136, 16)
        Me.labelPPTVorlage.TabIndex = 79
        Me.labelPPTVorlage.Text = "Powerpoint Template"
        Me.labelPPTVorlage.Visible = False
        '
        'einstellungen
        '
        Me.einstellungen.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.einstellungen.AutoSize = True
        Me.einstellungen.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, CType((System.Drawing.FontStyle.Italic Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.einstellungen.ForeColor = System.Drawing.SystemColors.MenuHighlight
        Me.einstellungen.Location = New System.Drawing.Point(346, 409)
        Me.einstellungen.Name = "einstellungen"
        Me.einstellungen.Size = New System.Drawing.Size(70, 13)
        Me.einstellungen.TabIndex = 81
        Me.einstellungen.Text = "Einstellungen"
        Me.einstellungen.Visible = False
        '
        'hryTreeView
        '
        Me.hryTreeView.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.hryTreeView.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.hryTreeView.FullRowSelect = True
        Me.hryTreeView.Location = New System.Drawing.Point(8, 120)
        Me.hryTreeView.Name = "hryTreeView"
        Me.hryTreeView.Size = New System.Drawing.Size(441, 196)
        Me.hryTreeView.TabIndex = 73
        '
        'frmHierarchySelection
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(457, 466)
        Me.Controls.Add(Me.Panel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmHierarchySelection"
        Me.Text = "Auswahl von Plan-Objekten"
        Me.TopMost = True
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        CType(Me.picturePhaseMilest, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picturePhasen, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pictureTyp, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pictureMilestones, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pictureBU, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pictureRoles, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pictureCosts, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SelectionSet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.collapseCompletely, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.expandCompletely, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SelectionReset, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker3 As System.ComponentModel.BackgroundWorker
    Friend WithEvents Panel1 As Windows.Forms.Panel
    Friend WithEvents auswLaden As Windows.Forms.Button
    Friend WithEvents Panel2 As Windows.Forms.Panel
    Friend WithEvents rdbPhaseMilest As Windows.Forms.RadioButton
    Friend WithEvents picturePhaseMilest As Windows.Forms.PictureBox
    Friend WithEvents rdbPhases As Windows.Forms.RadioButton
    Friend WithEvents picturePhasen As Windows.Forms.PictureBox
    Friend WithEvents pictureTyp As Windows.Forms.PictureBox
    Friend WithEvents rdbMilestones As Windows.Forms.RadioButton
    Friend WithEvents rdbTyp As Windows.Forms.RadioButton
    Friend WithEvents pictureMilestones As Windows.Forms.PictureBox
    Friend WithEvents pictureBU As Windows.Forms.PictureBox
    Friend WithEvents rdbRoles As Windows.Forms.RadioButton
    Friend WithEvents rdbBU As Windows.Forms.RadioButton
    Friend WithEvents pictureRoles As Windows.Forms.PictureBox
    Friend WithEvents pictureCosts As Windows.Forms.PictureBox
    Friend WithEvents rdbCosts As Windows.Forms.RadioButton
    Friend WithEvents filterBox As Windows.Forms.TextBox
    Friend WithEvents headerLine As Windows.Forms.Label
    Friend WithEvents SelectionSet As Windows.Forms.PictureBox
    Friend WithEvents auswSpeichern As Windows.Forms.Button
    Friend WithEvents filterLabel As Windows.Forms.Label
    Friend WithEvents filterDropbox As Windows.Forms.ComboBox
    Friend WithEvents collapseCompletely As Windows.Forms.PictureBox
    Friend WithEvents expandCompletely As Windows.Forms.PictureBox
    Friend WithEvents SelectionReset As Windows.Forms.PictureBox
    Friend WithEvents Panel3 As Windows.Forms.Panel
    Friend WithEvents rdbProjStruktTyp As Windows.Forms.RadioButton
    Friend WithEvents rdbNameList As Windows.Forms.RadioButton
    Friend WithEvents rdbProjStruktProj As Windows.Forms.RadioButton
    Friend WithEvents AbbrButton As Windows.Forms.Button
    Friend WithEvents chkbxOneChart As Windows.Forms.CheckBox
    Friend WithEvents statusLabel As Windows.Forms.Label
    Friend WithEvents repVorlagenDropbox As Windows.Forms.ComboBox
    Friend WithEvents OKButton As Windows.Forms.Button
    Friend WithEvents labelPPTVorlage As Windows.Forms.Label
    Friend WithEvents einstellungen As Windows.Forms.Label
    Friend WithEvents hryTreeView As Windows.Forms.TreeView
End Class
