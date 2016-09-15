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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmProjPortfolioAdmin))
        Me.TreeViewProjekte = New System.Windows.Forms.TreeView()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.txtboxLabel = New System.Windows.Forms.Label()
        Me.txtDropbox = New System.Windows.Forms.ComboBox()
        Me.considerDependencies = New System.Windows.Forms.CheckBox()
        Me.lblStandvom = New System.Windows.Forms.Label()
        Me.dropBoxTimeStamps = New System.Windows.Forms.ComboBox()
        Me.SelectionSet = New System.Windows.Forms.PictureBox()
        Me.collapseCompletely = New System.Windows.Forms.PictureBox()
        Me.expandCompletely = New System.Windows.Forms.PictureBox()
        Me.SelectionReset = New System.Windows.Forms.PictureBox()
        CType(Me.SelectionSet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.collapseCompletely, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.expandCompletely, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SelectionReset, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TreeViewProjekte
        '
        Me.TreeViewProjekte.Location = New System.Drawing.Point(34, 52)
        Me.TreeViewProjekte.Margin = New System.Windows.Forms.Padding(2)
        Me.TreeViewProjekte.Name = "TreeViewProjekte"
        Me.TreeViewProjekte.Size = New System.Drawing.Size(395, 290)
        Me.TreeViewProjekte.TabIndex = 1
        '
        'OKButton
        '
        Me.OKButton.Location = New System.Drawing.Point(175, 445)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(107, 23)
        Me.OKButton.TabIndex = 6
        Me.OKButton.Text = "Button1"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'txtboxLabel
        '
        Me.txtboxLabel.AutoSize = True
        Me.txtboxLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtboxLabel.Location = New System.Drawing.Point(31, 383)
        Me.txtboxLabel.Name = "txtboxLabel"
        Me.txtboxLabel.Size = New System.Drawing.Size(72, 13)
        Me.txtboxLabel.TabIndex = 33
        Me.txtboxLabel.Text = "Filter-Auswahl"
        '
        'txtDropbox
        '
        Me.txtDropbox.FormattingEnabled = True
        Me.txtDropbox.Location = New System.Drawing.Point(34, 406)
        Me.txtDropbox.Name = "txtDropbox"
        Me.txtDropbox.Size = New System.Drawing.Size(395, 21)
        Me.txtDropbox.TabIndex = 34
        '
        'considerDependencies
        '
        Me.considerDependencies.AutoSize = True
        Me.considerDependencies.Location = New System.Drawing.Point(257, 350)
        Me.considerDependencies.Name = "considerDependencies"
        Me.considerDependencies.Size = New System.Drawing.Size(178, 17)
        Me.considerDependencies.TabIndex = 35
        Me.considerDependencies.Text = "Abhängigkeiten berücksichtigen"
        Me.considerDependencies.UseVisualStyleBackColor = True
        Me.considerDependencies.Visible = False
        '
        'lblStandvom
        '
        Me.lblStandvom.AutoSize = True
        Me.lblStandvom.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStandvom.Location = New System.Drawing.Point(194, 29)
        Me.lblStandvom.Name = "lblStandvom"
        Me.lblStandvom.Size = New System.Drawing.Size(61, 13)
        Me.lblStandvom.TabIndex = 36
        Me.lblStandvom.Text = "Stand vom:"
        '
        'dropBoxTimeStamps
        '
        Me.dropBoxTimeStamps.FormattingEnabled = True
        Me.dropBoxTimeStamps.Location = New System.Drawing.Point(255, 26)
        Me.dropBoxTimeStamps.Name = "dropBoxTimeStamps"
        Me.dropBoxTimeStamps.Size = New System.Drawing.Size(172, 21)
        Me.dropBoxTimeStamps.TabIndex = 37
        '
        'SelectionSet
        '
        Me.SelectionSet.BackColor = System.Drawing.SystemColors.Control
        Me.SelectionSet.ErrorImage = CType(resources.GetObject("SelectionSet.ErrorImage"), System.Drawing.Image)
        Me.SelectionSet.Image = CType(resources.GetObject("SelectionSet.Image"), System.Drawing.Image)
        Me.SelectionSet.InitialImage = Nothing
        Me.SelectionSet.Location = New System.Drawing.Point(35, 348)
        Me.SelectionSet.Name = "SelectionSet"
        Me.SelectionSet.Size = New System.Drawing.Size(16, 16)
        Me.SelectionSet.TabIndex = 55
        Me.SelectionSet.TabStop = False
        '
        'collapseCompletely
        '
        Me.collapseCompletely.BackColor = System.Drawing.SystemColors.Control
        Me.collapseCompletely.Image = CType(resources.GetObject("collapseCompletely.Image"), System.Drawing.Image)
        Me.collapseCompletely.Location = New System.Drawing.Point(91, 348)
        Me.collapseCompletely.Name = "collapseCompletely"
        Me.collapseCompletely.Size = New System.Drawing.Size(16, 16)
        Me.collapseCompletely.TabIndex = 54
        Me.collapseCompletely.TabStop = False
        '
        'expandCompletely
        '
        Me.expandCompletely.BackColor = System.Drawing.SystemColors.Control
        Me.expandCompletely.Image = CType(resources.GetObject("expandCompletely.Image"), System.Drawing.Image)
        Me.expandCompletely.Location = New System.Drawing.Point(113, 348)
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
        Me.SelectionReset.Location = New System.Drawing.Point(55, 348)
        Me.SelectionReset.Name = "SelectionReset"
        Me.SelectionReset.Size = New System.Drawing.Size(16, 16)
        Me.SelectionReset.TabIndex = 52
        Me.SelectionReset.TabStop = False
        '
        'frmProjPortfolioAdmin
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(467, 489)
        Me.Controls.Add(Me.SelectionSet)
        Me.Controls.Add(Me.collapseCompletely)
        Me.Controls.Add(Me.expandCompletely)
        Me.Controls.Add(Me.SelectionReset)
        Me.Controls.Add(Me.dropBoxTimeStamps)
        Me.Controls.Add(Me.lblStandvom)
        Me.Controls.Add(Me.considerDependencies)
        Me.Controls.Add(Me.txtDropbox)
        Me.Controls.Add(Me.txtboxLabel)
        Me.Controls.Add(Me.OKButton)
        Me.Controls.Add(Me.TreeViewProjekte)
        Me.Name = "frmProjPortfolioAdmin"
        Me.Text = "Multiprojekt-Szenario"
        Me.TopMost = True
        CType(Me.SelectionSet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.collapseCompletely, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.expandCompletely, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SelectionReset, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents TreeViewProjekte As System.Windows.Forms.TreeView
    Friend WithEvents txtDropbox As System.Windows.Forms.ComboBox
    Friend WithEvents OKButton As System.Windows.Forms.Button
    Public WithEvents txtboxLabel As System.Windows.Forms.Label
    Friend WithEvents considerDependencies As System.Windows.Forms.CheckBox
    Public WithEvents lblStandvom As System.Windows.Forms.Label
    Friend WithEvents dropBoxTimeStamps As System.Windows.Forms.ComboBox
    Friend WithEvents SelectionSet As System.Windows.Forms.PictureBox
    Friend WithEvents collapseCompletely As System.Windows.Forms.PictureBox
    Friend WithEvents expandCompletely As System.Windows.Forms.PictureBox
    Friend WithEvents SelectionReset As System.Windows.Forms.PictureBox
End Class
