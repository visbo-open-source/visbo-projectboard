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
        Me.dropBoxTimeStamps = New System.Windows.Forms.ComboBox()
        Me.dropboxScenarioNames = New System.Windows.Forms.ComboBox()
        Me.ToolTipStand = New System.Windows.Forms.ToolTip(Me.components)
        Me.rdbTTDescription = New System.Windows.Forms.RadioButton()
        Me.rdbTTDependencies = New System.Windows.Forms.RadioButton()
        Me.deleteFilterIcon = New System.Windows.Forms.PictureBox()
        Me.filterIcon = New System.Windows.Forms.PictureBox()
        Me.SelectionSet = New System.Windows.Forms.PictureBox()
        Me.collapseCompletely = New System.Windows.Forms.PictureBox()
        Me.expandCompletely = New System.Windows.Forms.PictureBox()
        Me.SelectionReset = New System.Windows.Forms.PictureBox()
        Me.LblToolTipps = New System.Windows.Forms.Label()
        Me.onlyActive = New System.Windows.Forms.PictureBox()
        Me.onlyInactive = New System.Windows.Forms.PictureBox()
        Me.backToInit = New System.Windows.Forms.PictureBox()
        CType(Me.deleteFilterIcon, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.filterIcon, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SelectionSet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.collapseCompletely, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.expandCompletely, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SelectionReset, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.onlyActive, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.onlyInactive, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.backToInit, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TreeViewProjekte
        '
        Me.TreeViewProjekte.Location = New System.Drawing.Point(22, 52)
        Me.TreeViewProjekte.Margin = New System.Windows.Forms.Padding(2)
        Me.TreeViewProjekte.Name = "TreeViewProjekte"
        Me.TreeViewProjekte.Size = New System.Drawing.Size(421, 290)
        Me.TreeViewProjekte.TabIndex = 1
        '
        'OKButton
        '
        Me.OKButton.Location = New System.Drawing.Point(166, 433)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(118, 23)
        Me.OKButton.TabIndex = 6
        Me.OKButton.Text = "Button1"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'lblStandvom
        '
        Me.lblStandvom.AutoSize = True
        Me.lblStandvom.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStandvom.Location = New System.Drawing.Point(209, 29)
        Me.lblStandvom.Name = "lblStandvom"
        Me.lblStandvom.Size = New System.Drawing.Size(61, 13)
        Me.lblStandvom.TabIndex = 36
        Me.lblStandvom.Text = "Stand vom:"
        '
        'dropBoxTimeStamps
        '
        Me.dropBoxTimeStamps.FormattingEnabled = True
        Me.dropBoxTimeStamps.Location = New System.Drawing.Point(270, 25)
        Me.dropBoxTimeStamps.Name = "dropBoxTimeStamps"
        Me.dropBoxTimeStamps.Size = New System.Drawing.Size(172, 21)
        Me.dropBoxTimeStamps.TabIndex = 37
        '
        'dropboxScenarioNames
        '
        Me.dropboxScenarioNames.FormattingEnabled = True
        Me.dropboxScenarioNames.Location = New System.Drawing.Point(34, 395)
        Me.dropboxScenarioNames.Name = "dropboxScenarioNames"
        Me.dropboxScenarioNames.Size = New System.Drawing.Size(400, 21)
        Me.dropboxScenarioNames.TabIndex = 56
        '
        'ToolTipStand
        '
        '
        'rdbTTDescription
        '
        Me.rdbTTDescription.AutoSize = True
        Me.rdbTTDescription.Location = New System.Drawing.Point(71, 9)
        Me.rdbTTDescription.Name = "rdbTTDescription"
        Me.rdbTTDescription.Size = New System.Drawing.Size(126, 17)
        Me.rdbTTDescription.TabIndex = 61
        Me.rdbTTDescription.TabStop = True
        Me.rdbTTDescription.Text = "Projekt-Beschreibung"
        Me.ToolTipStand.SetToolTip(Me.rdbTTDescription, """Test für rdbTTShwDesc""")
        Me.rdbTTDescription.UseVisualStyleBackColor = True
        '
        'rdbTTDependencies
        '
        Me.rdbTTDependencies.AutoSize = True
        Me.rdbTTDependencies.Location = New System.Drawing.Point(71, 27)
        Me.rdbTTDependencies.Name = "rdbTTDependencies"
        Me.rdbTTDependencies.Size = New System.Drawing.Size(135, 17)
        Me.rdbTTDependencies.TabIndex = 62
        Me.rdbTTDependencies.TabStop = True
        Me.rdbTTDependencies.Text = "Projekt-Abhängigkeiten"
        Me.ToolTipStand.SetToolTip(Me.rdbTTDependencies, """Test für rdbTTShwDesc""")
        Me.rdbTTDependencies.UseVisualStyleBackColor = True
        '
        'deleteFilterIcon
        '
        Me.deleteFilterIcon.BackColor = System.Drawing.SystemColors.Control
        Me.deleteFilterIcon.Enabled = False
        Me.deleteFilterIcon.Location = New System.Drawing.Point(152, 347)
        Me.deleteFilterIcon.Name = "deleteFilterIcon"
        Me.deleteFilterIcon.Size = New System.Drawing.Size(16, 16)
        Me.deleteFilterIcon.TabIndex = 59
        Me.deleteFilterIcon.TabStop = False
        '
        'filterIcon
        '
        Me.filterIcon.BackColor = System.Drawing.SystemColors.Control
        Me.filterIcon.Image = Global.ProjectBoardBasic.My.Resources.Resources.funnel_add
        Me.filterIcon.Location = New System.Drawing.Point(130, 347)
        Me.filterIcon.Name = "filterIcon"
        Me.filterIcon.Size = New System.Drawing.Size(16, 16)
        Me.filterIcon.TabIndex = 57
        Me.filterIcon.TabStop = False
        '
        'SelectionSet
        '
        Me.SelectionSet.BackColor = System.Drawing.SystemColors.Control
        Me.SelectionSet.ErrorImage = CType(resources.GetObject("SelectionSet.ErrorImage"), System.Drawing.Image)
        Me.SelectionSet.Image = CType(resources.GetObject("SelectionSet.Image"), System.Drawing.Image)
        Me.SelectionSet.InitialImage = Nothing
        Me.SelectionSet.Location = New System.Drawing.Point(23, 348)
        Me.SelectionSet.Name = "SelectionSet"
        Me.SelectionSet.Size = New System.Drawing.Size(16, 16)
        Me.SelectionSet.TabIndex = 55
        Me.SelectionSet.TabStop = False
        '
        'collapseCompletely
        '
        Me.collapseCompletely.BackColor = System.Drawing.SystemColors.Control
        Me.collapseCompletely.Image = CType(resources.GetObject("collapseCompletely.Image"), System.Drawing.Image)
        Me.collapseCompletely.Location = New System.Drawing.Point(68, 348)
        Me.collapseCompletely.Name = "collapseCompletely"
        Me.collapseCompletely.Size = New System.Drawing.Size(16, 16)
        Me.collapseCompletely.TabIndex = 54
        Me.collapseCompletely.TabStop = False
        '
        'expandCompletely
        '
        Me.expandCompletely.BackColor = System.Drawing.SystemColors.Control
        Me.expandCompletely.Image = CType(resources.GetObject("expandCompletely.Image"), System.Drawing.Image)
        Me.expandCompletely.Location = New System.Drawing.Point(90, 348)
        Me.expandCompletely.Name = "expandCompletely"
        Me.expandCompletely.Size = New System.Drawing.Size(16, 16)
        Me.expandCompletely.TabIndex = 53
        Me.expandCompletely.TabStop = False
        '
        'SelectionReset
        '
        Me.SelectionReset.BackColor = System.Drawing.SystemColors.Control
        Me.SelectionReset.Image = CType(resources.GetObject("SelectionReset.Image"), System.Drawing.Image)
        Me.SelectionReset.InitialImage = Nothing
        Me.SelectionReset.Location = New System.Drawing.Point(42, 348)
        Me.SelectionReset.Name = "SelectionReset"
        Me.SelectionReset.Size = New System.Drawing.Size(16, 16)
        Me.SelectionReset.TabIndex = 52
        Me.SelectionReset.TabStop = False
        '
        'LblToolTipps
        '
        Me.LblToolTipps.AutoSize = True
        Me.LblToolTipps.Location = New System.Drawing.Point(19, 11)
        Me.LblToolTipps.Name = "LblToolTipps"
        Me.LblToolTipps.Size = New System.Drawing.Size(50, 13)
        Me.LblToolTipps.TabIndex = 60
        Me.LblToolTipps.Text = "Tooltipps"
        '
        'onlyActive
        '
        Me.onlyActive.BackColor = System.Drawing.SystemColors.Control
        Me.onlyActive.Enabled = False
        Me.onlyActive.Location = New System.Drawing.Point(230, 347)
        Me.onlyActive.Name = "onlyActive"
        Me.onlyActive.Size = New System.Drawing.Size(16, 16)
        Me.onlyActive.TabIndex = 63
        Me.onlyActive.TabStop = False
        Me.onlyActive.Visible = False
        '
        'onlyInactive
        '
        Me.onlyInactive.BackColor = System.Drawing.SystemColors.Control
        Me.onlyInactive.Enabled = False
        Me.onlyInactive.Location = New System.Drawing.Point(252, 347)
        Me.onlyInactive.Name = "onlyInactive"
        Me.onlyInactive.Size = New System.Drawing.Size(16, 16)
        Me.onlyInactive.TabIndex = 64
        Me.onlyInactive.TabStop = False
        Me.onlyInactive.Visible = False
        '
        'backToInit
        '
        Me.backToInit.BackColor = System.Drawing.SystemColors.Control
        Me.backToInit.Enabled = False
        Me.backToInit.Location = New System.Drawing.Point(274, 347)
        Me.backToInit.Name = "backToInit"
        Me.backToInit.Size = New System.Drawing.Size(16, 16)
        Me.backToInit.TabIndex = 65
        Me.backToInit.TabStop = False
        Me.backToInit.Visible = False
        '
        'frmProjPortfolioAdmin
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(467, 473)
        Me.Controls.Add(Me.backToInit)
        Me.Controls.Add(Me.onlyInactive)
        Me.Controls.Add(Me.onlyActive)
        Me.Controls.Add(Me.rdbTTDependencies)
        Me.Controls.Add(Me.rdbTTDescription)
        Me.Controls.Add(Me.LblToolTipps)
        Me.Controls.Add(Me.deleteFilterIcon)
        Me.Controls.Add(Me.filterIcon)
        Me.Controls.Add(Me.dropboxScenarioNames)
        Me.Controls.Add(Me.SelectionSet)
        Me.Controls.Add(Me.collapseCompletely)
        Me.Controls.Add(Me.expandCompletely)
        Me.Controls.Add(Me.SelectionReset)
        Me.Controls.Add(Me.dropBoxTimeStamps)
        Me.Controls.Add(Me.lblStandvom)
        Me.Controls.Add(Me.OKButton)
        Me.Controls.Add(Me.TreeViewProjekte)
        Me.Name = "frmProjPortfolioAdmin"
        Me.Text = "Multiprojekt-Szenario"
        Me.TopMost = True
        CType(Me.deleteFilterIcon, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.filterIcon, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SelectionSet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.collapseCompletely, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.expandCompletely, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SelectionReset, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.onlyActive, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.onlyInactive, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.backToInit, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents TreeViewProjekte As System.Windows.Forms.TreeView
    Friend WithEvents OKButton As System.Windows.Forms.Button
    Public WithEvents lblStandvom As System.Windows.Forms.Label
    Friend WithEvents dropBoxTimeStamps As System.Windows.Forms.ComboBox
    Friend WithEvents SelectionSet As System.Windows.Forms.PictureBox
    Friend WithEvents collapseCompletely As System.Windows.Forms.PictureBox
    Friend WithEvents expandCompletely As System.Windows.Forms.PictureBox
    Friend WithEvents SelectionReset As System.Windows.Forms.PictureBox
    Friend WithEvents dropboxScenarioNames As System.Windows.Forms.ComboBox
    Friend WithEvents filterIcon As System.Windows.Forms.PictureBox
    Friend WithEvents ToolTipStand As System.Windows.Forms.ToolTip
    Friend WithEvents deleteFilterIcon As System.Windows.Forms.PictureBox
    Friend WithEvents LblToolTipps As System.Windows.Forms.Label
    Friend WithEvents rdbTTDescription As System.Windows.Forms.RadioButton
    Friend WithEvents rdbTTDependencies As System.Windows.Forms.RadioButton
    Friend WithEvents onlyActive As System.Windows.Forms.PictureBox
    Friend WithEvents onlyInactive As System.Windows.Forms.PictureBox
    Friend WithEvents backToInit As System.Windows.Forms.PictureBox
End Class
