<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmProjPortfolioAdmin
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmProjPortfolioAdmin))
        Me.ToolTipStand = New System.Windows.Forms.ToolTip(Me.components)
        Me.portfolioBrowserHelp = New System.Windows.Forms.HelpProvider()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.chkbxPermanent = New System.Windows.Forms.CheckBox()
        Me.requiredDate = New System.Windows.Forms.DateTimePicker()
        Me.storeToDBasWell = New System.Windows.Forms.CheckBox()
        Me.backToInit = New System.Windows.Forms.PictureBox()
        Me.onlyInactive = New System.Windows.Forms.PictureBox()
        Me.onlyActive = New System.Windows.Forms.PictureBox()
        Me.deleteFilterIcon = New System.Windows.Forms.PictureBox()
        Me.filterIcon = New System.Windows.Forms.PictureBox()
        Me.dropboxScenarioNames = New System.Windows.Forms.ComboBox()
        Me.SelectionSet = New System.Windows.Forms.PictureBox()
        Me.collapseCompletely = New System.Windows.Forms.PictureBox()
        Me.expandCompletely = New System.Windows.Forms.PictureBox()
        Me.SelectionReset = New System.Windows.Forms.PictureBox()
        Me.lblStandvom = New System.Windows.Forms.Label()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.TreeViewProjekte = New System.Windows.Forms.TreeView()
        Me.Panel1.SuspendLayout()
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
        'ToolTipStand
        '
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.Controls.Add(Me.chkbxPermanent)
        Me.Panel1.Controls.Add(Me.requiredDate)
        Me.Panel1.Controls.Add(Me.storeToDBasWell)
        Me.Panel1.Controls.Add(Me.backToInit)
        Me.Panel1.Controls.Add(Me.onlyInactive)
        Me.Panel1.Controls.Add(Me.onlyActive)
        Me.Panel1.Controls.Add(Me.deleteFilterIcon)
        Me.Panel1.Controls.Add(Me.filterIcon)
        Me.Panel1.Controls.Add(Me.dropboxScenarioNames)
        Me.Panel1.Controls.Add(Me.SelectionSet)
        Me.Panel1.Controls.Add(Me.collapseCompletely)
        Me.Panel1.Controls.Add(Me.expandCompletely)
        Me.Panel1.Controls.Add(Me.SelectionReset)
        Me.Panel1.Controls.Add(Me.lblStandvom)
        Me.Panel1.Controls.Add(Me.OKButton)
        Me.Panel1.Controls.Add(Me.TreeViewProjekte)
        Me.Panel1.Location = New System.Drawing.Point(-5, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(460, 667)
        Me.Panel1.TabIndex = 0
        '
        'chkbxPermanent
        '
        Me.chkbxPermanent.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.chkbxPermanent.AutoSize = True
        Me.chkbxPermanent.Location = New System.Drawing.Point(332, 545)
        Me.chkbxPermanent.Margin = New System.Windows.Forms.Padding(4)
        Me.chkbxPermanent.Name = "chkbxPermanent"
        Me.chkbxPermanent.Size = New System.Drawing.Size(98, 21)
        Me.chkbxPermanent.TabIndex = 91
        Me.chkbxPermanent.Text = "permanent"
        Me.chkbxPermanent.UseVisualStyleBackColor = True
        '
        'requiredDate
        '
        Me.requiredDate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.requiredDate.Location = New System.Drawing.Point(160, 11)
        Me.requiredDate.Margin = New System.Windows.Forms.Padding(4)
        Me.requiredDate.Name = "requiredDate"
        Me.requiredDate.ShowCheckBox = True
        Me.requiredDate.Size = New System.Drawing.Size(265, 22)
        Me.requiredDate.TabIndex = 90
        '
        'storeToDBasWell
        '
        Me.storeToDBasWell.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.storeToDBasWell.AutoSize = True
        Me.storeToDBasWell.Location = New System.Drawing.Point(320, 606)
        Me.storeToDBasWell.Margin = New System.Windows.Forms.Padding(4)
        Me.storeToDBasWell.Name = "storeToDBasWell"
        Me.storeToDBasWell.Size = New System.Drawing.Size(101, 21)
        Me.storeToDBasWell.TabIndex = 89
        Me.storeToDBasWell.Text = "store to DB"
        Me.storeToDBasWell.UseVisualStyleBackColor = True
        '
        'backToInit
        '
        Me.backToInit.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.backToInit.BackColor = System.Drawing.SystemColors.Control
        Me.backToInit.Image = Global.ProjectBoardBasic.My.Resources.Resources.funnel_delete
        Me.backToInit.Location = New System.Drawing.Point(286, 544)
        Me.backToInit.Margin = New System.Windows.Forms.Padding(4)
        Me.backToInit.Name = "backToInit"
        Me.backToInit.Size = New System.Drawing.Size(21, 20)
        Me.backToInit.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.backToInit.TabIndex = 85
        Me.backToInit.TabStop = False
        '
        'onlyInactive
        '
        Me.onlyInactive.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.onlyInactive.BackColor = System.Drawing.SystemColors.Control
        Me.onlyInactive.Image = Global.ProjectBoardBasic.My.Resources.Resources.nur_ungecheckte_Projekte
        Me.onlyInactive.Location = New System.Drawing.Point(256, 544)
        Me.onlyInactive.Margin = New System.Windows.Forms.Padding(4)
        Me.onlyInactive.Name = "onlyInactive"
        Me.onlyInactive.Size = New System.Drawing.Size(21, 20)
        Me.onlyInactive.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.onlyInactive.TabIndex = 84
        Me.onlyInactive.TabStop = False
        '
        'onlyActive
        '
        Me.onlyActive.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.onlyActive.BackColor = System.Drawing.SystemColors.Control
        Me.onlyActive.Image = Global.ProjectBoardBasic.My.Resources.Resources.nur_gecheckte_Projekte
        Me.onlyActive.Location = New System.Drawing.Point(227, 544)
        Me.onlyActive.Margin = New System.Windows.Forms.Padding(4)
        Me.onlyActive.Name = "onlyActive"
        Me.onlyActive.Size = New System.Drawing.Size(21, 20)
        Me.onlyActive.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.onlyActive.TabIndex = 83
        Me.onlyActive.TabStop = False
        '
        'deleteFilterIcon
        '
        Me.deleteFilterIcon.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.deleteFilterIcon.BackColor = System.Drawing.SystemColors.Control
        Me.deleteFilterIcon.Enabled = False
        Me.deleteFilterIcon.Location = New System.Drawing.Point(186, 544)
        Me.deleteFilterIcon.Margin = New System.Windows.Forms.Padding(4)
        Me.deleteFilterIcon.Name = "deleteFilterIcon"
        Me.deleteFilterIcon.Size = New System.Drawing.Size(21, 20)
        Me.deleteFilterIcon.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.deleteFilterIcon.TabIndex = 82
        Me.deleteFilterIcon.TabStop = False
        '
        'filterIcon
        '
        Me.filterIcon.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.filterIcon.BackColor = System.Drawing.SystemColors.Control
        Me.filterIcon.Image = Global.ProjectBoardBasic.My.Resources.Resources.funnel_add
        Me.filterIcon.Location = New System.Drawing.Point(156, 544)
        Me.filterIcon.Margin = New System.Windows.Forms.Padding(4)
        Me.filterIcon.Name = "filterIcon"
        Me.filterIcon.Size = New System.Drawing.Size(21, 20)
        Me.filterIcon.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.filterIcon.TabIndex = 81
        Me.filterIcon.TabStop = False
        '
        'dropboxScenarioNames
        '
        Me.dropboxScenarioNames.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dropboxScenarioNames.FormattingEnabled = True
        Me.dropboxScenarioNames.Location = New System.Drawing.Point(29, 574)
        Me.dropboxScenarioNames.Margin = New System.Windows.Forms.Padding(4)
        Me.dropboxScenarioNames.Name = "dropboxScenarioNames"
        Me.dropboxScenarioNames.Size = New System.Drawing.Size(396, 24)
        Me.dropboxScenarioNames.TabIndex = 80
        '
        'SelectionSet
        '
        Me.SelectionSet.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.SelectionSet.BackColor = System.Drawing.SystemColors.Control
        Me.SelectionSet.ErrorImage = CType(resources.GetObject("SelectionSet.ErrorImage"), System.Drawing.Image)
        Me.SelectionSet.Image = CType(resources.GetObject("SelectionSet.Image"), System.Drawing.Image)
        Me.SelectionSet.InitialImage = Nothing
        Me.SelectionSet.Location = New System.Drawing.Point(30, 545)
        Me.SelectionSet.Margin = New System.Windows.Forms.Padding(4)
        Me.SelectionSet.Name = "SelectionSet"
        Me.SelectionSet.Size = New System.Drawing.Size(21, 20)
        Me.SelectionSet.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.SelectionSet.TabIndex = 79
        Me.SelectionSet.TabStop = False
        '
        'collapseCompletely
        '
        Me.collapseCompletely.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.collapseCompletely.BackColor = System.Drawing.SystemColors.Control
        Me.collapseCompletely.Cursor = System.Windows.Forms.Cursors.Default
        Me.collapseCompletely.Image = CType(resources.GetObject("collapseCompletely.Image"), System.Drawing.Image)
        Me.collapseCompletely.Location = New System.Drawing.Point(90, 545)
        Me.collapseCompletely.Margin = New System.Windows.Forms.Padding(4)
        Me.collapseCompletely.Name = "collapseCompletely"
        Me.collapseCompletely.Size = New System.Drawing.Size(21, 20)
        Me.collapseCompletely.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.collapseCompletely.TabIndex = 78
        Me.collapseCompletely.TabStop = False
        '
        'expandCompletely
        '
        Me.expandCompletely.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.expandCompletely.BackColor = System.Drawing.SystemColors.Control
        Me.expandCompletely.Image = CType(resources.GetObject("expandCompletely.Image"), System.Drawing.Image)
        Me.expandCompletely.Location = New System.Drawing.Point(119, 545)
        Me.expandCompletely.Margin = New System.Windows.Forms.Padding(4)
        Me.expandCompletely.Name = "expandCompletely"
        Me.expandCompletely.Size = New System.Drawing.Size(21, 20)
        Me.expandCompletely.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.expandCompletely.TabIndex = 77
        Me.expandCompletely.TabStop = False
        '
        'SelectionReset
        '
        Me.SelectionReset.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.SelectionReset.BackColor = System.Drawing.SystemColors.Control
        Me.SelectionReset.Image = CType(resources.GetObject("SelectionReset.Image"), System.Drawing.Image)
        Me.SelectionReset.InitialImage = Nothing
        Me.SelectionReset.Location = New System.Drawing.Point(55, 545)
        Me.SelectionReset.Margin = New System.Windows.Forms.Padding(4)
        Me.SelectionReset.Name = "SelectionReset"
        Me.SelectionReset.Size = New System.Drawing.Size(21, 20)
        Me.SelectionReset.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.SelectionReset.TabIndex = 76
        Me.SelectionReset.TabStop = False
        '
        'lblStandvom
        '
        Me.lblStandvom.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblStandvom.AutoSize = True
        Me.lblStandvom.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStandvom.Location = New System.Drawing.Point(66, 14)
        Me.lblStandvom.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblStandvom.Name = "lblStandvom"
        Me.lblStandvom.Size = New System.Drawing.Size(79, 17)
        Me.lblStandvom.TabIndex = 75
        Me.lblStandvom.Text = "Stand vom:"
        '
        'OKButton
        '
        Me.OKButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.OKButton.Location = New System.Drawing.Point(29, 627)
        Me.OKButton.Margin = New System.Windows.Forms.Padding(4)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(250, 28)
        Me.OKButton.TabIndex = 74
        Me.OKButton.Text = "Button1"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'TreeViewProjekte
        '
        Me.TreeViewProjekte.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TreeViewProjekte.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TreeViewProjekte.Location = New System.Drawing.Point(32, 41)
        Me.TreeViewProjekte.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.TreeViewProjekte.Name = "TreeViewProjekte"
        Me.TreeViewProjekte.Size = New System.Drawing.Size(394, 488)
        Me.TreeViewProjekte.TabIndex = 73
        '
        'frmProjPortfolioAdmin
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(120.0!, 120.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.ClientSize = New System.Drawing.Size(457, 669)
        Me.Controls.Add(Me.Panel1)
        Me.portfolioBrowserHelp.SetHelpString(Me, """das ist die Hilfe""")
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "frmProjPortfolioAdmin"
        Me.portfolioBrowserHelp.SetShowHelp(Me, True)
        Me.Text = "Portfolio"
        Me.TopMost = True
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
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

    End Sub
    Friend WithEvents ToolTipStand As System.Windows.Forms.ToolTip
    Friend WithEvents portfolioBrowserHelp As System.Windows.Forms.HelpProvider
    Friend WithEvents Panel1 As Windows.Forms.Panel
    Friend WithEvents chkbxPermanent As Windows.Forms.CheckBox
    Friend WithEvents requiredDate As Windows.Forms.DateTimePicker
    Friend WithEvents storeToDBasWell As Windows.Forms.CheckBox
    Friend WithEvents backToInit As Windows.Forms.PictureBox
    Friend WithEvents onlyInactive As Windows.Forms.PictureBox
    Friend WithEvents onlyActive As Windows.Forms.PictureBox
    Friend WithEvents deleteFilterIcon As Windows.Forms.PictureBox
    Friend WithEvents filterIcon As Windows.Forms.PictureBox
    Friend WithEvents dropboxScenarioNames As Windows.Forms.ComboBox
    Friend WithEvents SelectionSet As Windows.Forms.PictureBox
    Friend WithEvents collapseCompletely As Windows.Forms.PictureBox
    Friend WithEvents expandCompletely As Windows.Forms.PictureBox
    Friend WithEvents SelectionReset As Windows.Forms.PictureBox
    Public WithEvents lblStandvom As Windows.Forms.Label
    Friend WithEvents OKButton As Windows.Forms.Button
    Public WithEvents TreeViewProjekte As Windows.Forms.TreeView
End Class
