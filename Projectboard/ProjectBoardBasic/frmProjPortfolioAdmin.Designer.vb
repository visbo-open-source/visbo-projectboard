<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmProjPortfolioAdmin
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmProjPortfolioAdmin))
        Me.TreeViewProjekte = New System.Windows.Forms.TreeView()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.lblStandvom = New System.Windows.Forms.Label()
        Me.dropboxScenarioNames = New System.Windows.Forms.ComboBox()
        Me.ToolTipStand = New System.Windows.Forms.ToolTip(Me.components)
        Me.lblVersionen1 = New System.Windows.Forms.Label()
        Me.versionsToKeep = New System.Windows.Forms.NumericUpDown()
        Me.lblVersionen2 = New System.Windows.Forms.Label()
        Me.backToInit = New System.Windows.Forms.PictureBox()
        Me.onlyInactive = New System.Windows.Forms.PictureBox()
        Me.onlyActive = New System.Windows.Forms.PictureBox()
        Me.deleteFilterIcon = New System.Windows.Forms.PictureBox()
        Me.filterIcon = New System.Windows.Forms.PictureBox()
        Me.SelectionSet = New System.Windows.Forms.PictureBox()
        Me.collapseCompletely = New System.Windows.Forms.PictureBox()
        Me.expandCompletely = New System.Windows.Forms.PictureBox()
        Me.SelectionReset = New System.Windows.Forms.PictureBox()
        Me.storeToDBasWell = New System.Windows.Forms.CheckBox()
        Me.requiredDate = New System.Windows.Forms.DateTimePicker()
        Me.chkbxPermanent = New System.Windows.Forms.CheckBox()
        Me.portfolioBrowserHelp = New System.Windows.Forms.HelpProvider()
        CType(Me.versionsToKeep, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.backToInit, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.onlyInactive, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.onlyActive, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.deleteFilterIcon, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.filterIcon, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SelectionSet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.collapseCompletely, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.expandCompletely, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SelectionReset, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TreeViewProjekte
        '
        Me.TreeViewProjekte.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TreeViewProjekte.Location = New System.Drawing.Point(29, 39)
        Me.TreeViewProjekte.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TreeViewProjekte.Name = "TreeViewProjekte"
        Me.TreeViewProjekte.Size = New System.Drawing.Size(393, 381)
        Me.TreeViewProjekte.TabIndex = 1
        '
        'OKButton
        '
        Me.OKButton.Location = New System.Drawing.Point(101, 542)
        Me.OKButton.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(251, 28)
        Me.OKButton.TabIndex = 6
        Me.OKButton.Text = "Button1"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'lblStandvom
        '
        Me.lblStandvom.AutoSize = True
        Me.lblStandvom.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStandvom.Location = New System.Drawing.Point(67, 12)
        Me.lblStandvom.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblStandvom.Name = "lblStandvom"
        Me.lblStandvom.Size = New System.Drawing.Size(79, 17)
        Me.lblStandvom.TabIndex = 36
        Me.lblStandvom.Text = "Stand vom:"
        '
        'dropboxScenarioNames
        '
        Me.dropboxScenarioNames.FormattingEnabled = True
        Me.dropboxScenarioNames.Location = New System.Drawing.Point(29, 486)
        Me.dropboxScenarioNames.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dropboxScenarioNames.Name = "dropboxScenarioNames"
        Me.dropboxScenarioNames.Size = New System.Drawing.Size(396, 24)
        Me.dropboxScenarioNames.TabIndex = 56
        '
        'ToolTipStand
        '
        '
        'lblVersionen1
        '
        Me.lblVersionen1.AutoSize = True
        Me.lblVersionen1.Location = New System.Drawing.Point(29, 465)
        Me.lblVersionen1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblVersionen1.Name = "lblVersionen1"
        Me.lblVersionen1.Size = New System.Drawing.Size(141, 17)
        Me.lblVersionen1.TabIndex = 66
        Me.lblVersionen1.Text = "alles löschen, ausser"
        Me.lblVersionen1.Visible = False
        '
        'versionsToKeep
        '
        Me.versionsToKeep.Location = New System.Drawing.Point(176, 462)
        Me.versionsToKeep.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.versionsToKeep.Minimum = New Decimal(New Integer() {2, 0, 0, 0})
        Me.versionsToKeep.Name = "versionsToKeep"
        Me.versionsToKeep.Size = New System.Drawing.Size(57, 22)
        Me.versionsToKeep.TabIndex = 67
        Me.versionsToKeep.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.versionsToKeep.Value = New Decimal(New Integer() {3, 0, 0, 0})
        Me.versionsToKeep.Visible = False
        '
        'lblVersionen2
        '
        Me.lblVersionen2.AutoSize = True
        Me.lblVersionen2.Location = New System.Drawing.Point(240, 465)
        Me.lblVersionen2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblVersionen2.Name = "lblVersionen2"
        Me.lblVersionen2.Size = New System.Drawing.Size(187, 17)
        Me.lblVersionen2.TabIndex = 68
        Me.lblVersionen2.Text = "unterschiedlichen Versionen"
        Me.lblVersionen2.Visible = False
        '
        'backToInit
        '
        Me.backToInit.BackColor = System.Drawing.SystemColors.Control
        Me.backToInit.Image = Global.ProjectBoardBasic.My.Resources.Resources.funnel_delete
        Me.backToInit.Location = New System.Drawing.Point(287, 427)
        Me.backToInit.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.backToInit.Name = "backToInit"
        Me.backToInit.Size = New System.Drawing.Size(21, 20)
        Me.backToInit.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.backToInit.TabIndex = 65
        Me.backToInit.TabStop = False
        '
        'onlyInactive
        '
        Me.onlyInactive.BackColor = System.Drawing.SystemColors.Control
        Me.onlyInactive.Image = Global.ProjectBoardBasic.My.Resources.Resources.nur_ungecheckte_Projekte
        Me.onlyInactive.Location = New System.Drawing.Point(257, 427)
        Me.onlyInactive.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.onlyInactive.Name = "onlyInactive"
        Me.onlyInactive.Size = New System.Drawing.Size(21, 20)
        Me.onlyInactive.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.onlyInactive.TabIndex = 64
        Me.onlyInactive.TabStop = False
        '
        'onlyActive
        '
        Me.onlyActive.BackColor = System.Drawing.SystemColors.Control
        Me.onlyActive.Image = Global.ProjectBoardBasic.My.Resources.Resources.nur_gecheckte_Projekte
        Me.onlyActive.Location = New System.Drawing.Point(228, 427)
        Me.onlyActive.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.onlyActive.Name = "onlyActive"
        Me.onlyActive.Size = New System.Drawing.Size(21, 20)
        Me.onlyActive.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.onlyActive.TabIndex = 63
        Me.onlyActive.TabStop = False
        '
        'deleteFilterIcon
        '
        Me.deleteFilterIcon.BackColor = System.Drawing.SystemColors.Control
        Me.deleteFilterIcon.Enabled = False
        Me.deleteFilterIcon.Location = New System.Drawing.Point(187, 427)
        Me.deleteFilterIcon.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.deleteFilterIcon.Name = "deleteFilterIcon"
        Me.deleteFilterIcon.Size = New System.Drawing.Size(21, 20)
        Me.deleteFilterIcon.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.deleteFilterIcon.TabIndex = 59
        Me.deleteFilterIcon.TabStop = False
        '
        'filterIcon
        '
        Me.filterIcon.BackColor = System.Drawing.SystemColors.Control
        Me.filterIcon.Image = Global.ProjectBoardBasic.My.Resources.Resources.funnel_add
        Me.filterIcon.Location = New System.Drawing.Point(157, 427)
        Me.filterIcon.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.filterIcon.Name = "filterIcon"
        Me.filterIcon.Size = New System.Drawing.Size(21, 20)
        Me.filterIcon.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.filterIcon.TabIndex = 57
        Me.filterIcon.TabStop = False
        '
        'SelectionSet
        '
        Me.SelectionSet.BackColor = System.Drawing.SystemColors.Control
        Me.SelectionSet.ErrorImage = CType(resources.GetObject("SelectionSet.ErrorImage"), System.Drawing.Image)
        Me.SelectionSet.Image = CType(resources.GetObject("SelectionSet.Image"), System.Drawing.Image)
        Me.SelectionSet.InitialImage = Nothing
        Me.SelectionSet.Location = New System.Drawing.Point(31, 428)
        Me.SelectionSet.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.SelectionSet.Name = "SelectionSet"
        Me.SelectionSet.Size = New System.Drawing.Size(21, 20)
        Me.SelectionSet.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.SelectionSet.TabIndex = 55
        Me.SelectionSet.TabStop = False
        '
        'collapseCompletely
        '
        Me.collapseCompletely.BackColor = System.Drawing.SystemColors.Control
        Me.collapseCompletely.Cursor = System.Windows.Forms.Cursors.Default
        Me.collapseCompletely.Image = CType(resources.GetObject("collapseCompletely.Image"), System.Drawing.Image)
        Me.collapseCompletely.Location = New System.Drawing.Point(91, 428)
        Me.collapseCompletely.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.collapseCompletely.Name = "collapseCompletely"
        Me.collapseCompletely.Size = New System.Drawing.Size(21, 20)
        Me.collapseCompletely.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.collapseCompletely.TabIndex = 54
        Me.collapseCompletely.TabStop = False
        '
        'expandCompletely
        '
        Me.expandCompletely.BackColor = System.Drawing.SystemColors.Control
        Me.expandCompletely.Image = CType(resources.GetObject("expandCompletely.Image"), System.Drawing.Image)
        Me.expandCompletely.Location = New System.Drawing.Point(120, 428)
        Me.expandCompletely.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.expandCompletely.Name = "expandCompletely"
        Me.expandCompletely.Size = New System.Drawing.Size(21, 20)
        Me.expandCompletely.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.expandCompletely.TabIndex = 53
        Me.expandCompletely.TabStop = False
        '
        'SelectionReset
        '
        Me.SelectionReset.BackColor = System.Drawing.SystemColors.Control
        Me.SelectionReset.Image = CType(resources.GetObject("SelectionReset.Image"), System.Drawing.Image)
        Me.SelectionReset.InitialImage = Nothing
        Me.SelectionReset.Location = New System.Drawing.Point(56, 428)
        Me.SelectionReset.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.SelectionReset.Name = "SelectionReset"
        Me.SelectionReset.Size = New System.Drawing.Size(21, 20)
        Me.SelectionReset.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.SelectionReset.TabIndex = 52
        Me.SelectionReset.TabStop = False
        '
        'storeToDBasWell
        '
        Me.storeToDBasWell.AutoSize = True
        Me.storeToDBasWell.Location = New System.Drawing.Point(321, 519)
        Me.storeToDBasWell.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.storeToDBasWell.Name = "storeToDBasWell"
        Me.storeToDBasWell.Size = New System.Drawing.Size(101, 21)
        Me.storeToDBasWell.TabIndex = 70
        Me.storeToDBasWell.Text = "store to DB"
        Me.storeToDBasWell.UseVisualStyleBackColor = True
        '
        'requiredDate
        '
        Me.requiredDate.Location = New System.Drawing.Point(156, 9)
        Me.requiredDate.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.requiredDate.Name = "requiredDate"
        Me.requiredDate.ShowCheckBox = True
        Me.requiredDate.Size = New System.Drawing.Size(265, 22)
        Me.requiredDate.TabIndex = 71
        '
        'chkbxPermanent
        '
        Me.chkbxPermanent.AutoSize = True
        Me.chkbxPermanent.Location = New System.Drawing.Point(333, 428)
        Me.chkbxPermanent.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.chkbxPermanent.Name = "chkbxPermanent"
        Me.chkbxPermanent.Size = New System.Drawing.Size(98, 21)
        Me.chkbxPermanent.TabIndex = 72
        Me.chkbxPermanent.Text = "permanent"
        Me.chkbxPermanent.UseVisualStyleBackColor = True
        '
        'frmProjPortfolioAdmin
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(457, 582)
        Me.Controls.Add(Me.chkbxPermanent)
        Me.Controls.Add(Me.requiredDate)
        Me.Controls.Add(Me.storeToDBasWell)
        Me.Controls.Add(Me.lblVersionen2)
        Me.Controls.Add(Me.versionsToKeep)
        Me.Controls.Add(Me.lblVersionen1)
        Me.Controls.Add(Me.backToInit)
        Me.Controls.Add(Me.onlyInactive)
        Me.Controls.Add(Me.onlyActive)
        Me.Controls.Add(Me.deleteFilterIcon)
        Me.Controls.Add(Me.filterIcon)
        Me.Controls.Add(Me.dropboxScenarioNames)
        Me.Controls.Add(Me.SelectionSet)
        Me.Controls.Add(Me.collapseCompletely)
        Me.Controls.Add(Me.expandCompletely)
        Me.Controls.Add(Me.SelectionReset)
        Me.Controls.Add(Me.lblStandvom)
        Me.Controls.Add(Me.OKButton)
        Me.Controls.Add(Me.TreeViewProjekte)
        Me.portfolioBrowserHelp.SetHelpString(Me, """das ist die Hilfe""")
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Name = "frmProjPortfolioAdmin"
        Me.portfolioBrowserHelp.SetShowHelp(Me, True)
        Me.Text = "Portfolio"
        Me.TopMost = True
        CType(Me.versionsToKeep, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.backToInit, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.onlyInactive, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.onlyActive, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.deleteFilterIcon, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.filterIcon, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SelectionSet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.collapseCompletely, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.expandCompletely, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SelectionReset, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents TreeViewProjekte As System.Windows.Forms.TreeView
    Friend WithEvents OKButton As System.Windows.Forms.Button
    Public WithEvents lblStandvom As System.Windows.Forms.Label
    Friend WithEvents SelectionSet As System.Windows.Forms.PictureBox
    Friend WithEvents collapseCompletely As System.Windows.Forms.PictureBox
    Friend WithEvents expandCompletely As System.Windows.Forms.PictureBox
    Friend WithEvents SelectionReset As System.Windows.Forms.PictureBox
    Friend WithEvents dropboxScenarioNames As System.Windows.Forms.ComboBox
    Friend WithEvents filterIcon As System.Windows.Forms.PictureBox
    Friend WithEvents ToolTipStand As System.Windows.Forms.ToolTip
    Friend WithEvents deleteFilterIcon As System.Windows.Forms.PictureBox
    Friend WithEvents onlyActive As System.Windows.Forms.PictureBox
    Friend WithEvents onlyInactive As System.Windows.Forms.PictureBox
    Friend WithEvents backToInit As System.Windows.Forms.PictureBox
    Friend WithEvents lblVersionen1 As System.Windows.Forms.Label
    Friend WithEvents versionsToKeep As System.Windows.Forms.NumericUpDown
    Friend WithEvents lblVersionen2 As System.Windows.Forms.Label
    Friend WithEvents storeToDBasWell As System.Windows.Forms.CheckBox
    Friend WithEvents requiredDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents chkbxPermanent As System.Windows.Forms.CheckBox
    Friend WithEvents portfolioBrowserHelp As System.Windows.Forms.HelpProvider
End Class
