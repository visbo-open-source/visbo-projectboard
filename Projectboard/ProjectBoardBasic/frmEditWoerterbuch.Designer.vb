<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEditWoerterbuch
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmEditWoerterbuch))
        Me.rdbListShowsPhases = New System.Windows.Forms.RadioButton()
        Me.rdbListShowsMilestones = New System.Windows.Forms.RadioButton()
        Me.filterUnknown = New System.Windows.Forms.TextBox()
        Me.unknownList = New System.Windows.Forms.ListBox()
        Me.standardList = New System.Windows.Forms.ListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.filterStandard = New System.Windows.Forms.TextBox()
        Me.editUnknownItem = New System.Windows.Forms.TextBox()
        Me.ignoreButton = New System.Windows.Forms.Button()
        Me.addRulesToDictionary = New System.Windows.Forms.Button()
        Me.replaceButton = New System.Windows.Forms.Button()
        Me.visElements = New System.Windows.Forms.Button()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.clearUnknownList = New System.Windows.Forms.PictureBox()
        Me.setItemToBeUnknown = New System.Windows.Forms.PictureBox()
        Me.setItemToBeKnown = New System.Windows.Forms.PictureBox()
        Me.clearStandardList = New System.Windows.Forms.PictureBox()
        Me.showOnlySummaryTasks = New System.Windows.Forms.CheckBox()
        Me.storeButton = New System.Windows.Forms.Button()
        Me.filtersAreCoupled = New System.Windows.Forms.CheckBox()
        Me.StatusStrip1.SuspendLayout()
        CType(Me.clearUnknownList, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.setItemToBeUnknown, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.setItemToBeKnown, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.clearStandardList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'rdbListShowsPhases
        '
        Me.rdbListShowsPhases.AutoSize = True
        Me.rdbListShowsPhases.Location = New System.Drawing.Point(22, 17)
        Me.rdbListShowsPhases.Name = "rdbListShowsPhases"
        Me.rdbListShowsPhases.Size = New System.Drawing.Size(112, 17)
        Me.rdbListShowsPhases.TabIndex = 0
        Me.rdbListShowsPhases.TabStop = True
        Me.rdbListShowsPhases.Text = "Phasen/Vorgänge"
        Me.rdbListShowsPhases.UseVisualStyleBackColor = True
        '
        'rdbListShowsMilestones
        '
        Me.rdbListShowsMilestones.AutoSize = True
        Me.rdbListShowsMilestones.Location = New System.Drawing.Point(183, 17)
        Me.rdbListShowsMilestones.Name = "rdbListShowsMilestones"
        Me.rdbListShowsMilestones.Size = New System.Drawing.Size(84, 17)
        Me.rdbListShowsMilestones.TabIndex = 1
        Me.rdbListShowsMilestones.TabStop = True
        Me.rdbListShowsMilestones.Text = "Meilensteine"
        Me.rdbListShowsMilestones.UseVisualStyleBackColor = True
        '
        'filterUnknown
        '
        Me.filterUnknown.Location = New System.Drawing.Point(22, 79)
        Me.filterUnknown.Name = "filterUnknown"
        Me.filterUnknown.Size = New System.Drawing.Size(341, 20)
        Me.filterUnknown.TabIndex = 2
        '
        'unknownList
        '
        Me.unknownList.FormattingEnabled = True
        Me.unknownList.HorizontalScrollbar = True
        Me.unknownList.Location = New System.Drawing.Point(22, 122)
        Me.unknownList.Name = "unknownList"
        Me.unknownList.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.unknownList.Size = New System.Drawing.Size(341, 264)
        Me.unknownList.Sorted = True
        Me.unknownList.TabIndex = 3
        '
        'standardList
        '
        Me.standardList.FormattingEnabled = True
        Me.standardList.HorizontalScrollbar = True
        Me.standardList.Location = New System.Drawing.Point(410, 122)
        Me.standardList.Name = "standardList"
        Me.standardList.Size = New System.Drawing.Size(341, 264)
        Me.standardList.Sorted = True
        Me.standardList.TabIndex = 4
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(18, 48)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(236, 20)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Unbekannte Bezeichnungen"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(406, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(212, 20)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Standard Bezeichnungen"
        '
        'filterStandard
        '
        Me.filterStandard.Location = New System.Drawing.Point(410, 79)
        Me.filterStandard.Name = "filterStandard"
        Me.filterStandard.Size = New System.Drawing.Size(341, 20)
        Me.filterStandard.TabIndex = 7
        '
        'editUnknownItem
        '
        Me.editUnknownItem.Enabled = False
        Me.editUnknownItem.Location = New System.Drawing.Point(22, 405)
        Me.editUnknownItem.Name = "editUnknownItem"
        Me.editUnknownItem.Size = New System.Drawing.Size(729, 20)
        Me.editUnknownItem.TabIndex = 8
        '
        'ignoreButton
        '
        Me.ignoreButton.BackColor = System.Drawing.Color.Firebrick
        Me.ignoreButton.Enabled = False
        Me.ignoreButton.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.ignoreButton.Location = New System.Drawing.Point(138, 468)
        Me.ignoreButton.Name = "ignoreButton"
        Me.ignoreButton.Size = New System.Drawing.Size(82, 57)
        Me.ignoreButton.TabIndex = 9
        Me.ignoreButton.Text = "Element immer ignorieren"
        Me.ignoreButton.UseVisualStyleBackColor = False
        '
        'addRulesToDictionary
        '
        Me.addRulesToDictionary.BackColor = System.Drawing.Color.SteelBlue
        Me.addRulesToDictionary.Enabled = False
        Me.addRulesToDictionary.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.addRulesToDictionary.Location = New System.Drawing.Point(288, 468)
        Me.addRulesToDictionary.Name = "addRulesToDictionary"
        Me.addRulesToDictionary.Size = New System.Drawing.Size(189, 57)
        Me.addRulesToDictionary.TabIndex = 10
        Me.addRulesToDictionary.Text = "Paar zum Wörterbuch hinzufügen"
        Me.addRulesToDictionary.UseVisualStyleBackColor = False
        '
        'replaceButton
        '
        Me.replaceButton.BackColor = System.Drawing.Color.DarkOrange
        Me.replaceButton.Enabled = False
        Me.replaceButton.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.replaceButton.Location = New System.Drawing.Point(552, 468)
        Me.replaceButton.Name = "replaceButton"
        Me.replaceButton.Size = New System.Drawing.Size(82, 57)
        Me.replaceButton.TabIndex = 11
        Me.replaceButton.Text = "Standard Bezeichnung ändern"
        Me.replaceButton.UseVisualStyleBackColor = False
        '
        'visElements
        '
        Me.visElements.BackColor = System.Drawing.Color.Green
        Me.visElements.Enabled = False
        Me.visElements.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.visElements.Location = New System.Drawing.Point(288, 431)
        Me.visElements.Name = "visElements"
        Me.visElements.Size = New System.Drawing.Size(189, 31)
        Me.visElements.TabIndex = 12
        Me.visElements.Text = "Elemente visualisieren"
        Me.visElements.UseVisualStyleBackColor = False
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 540)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(774, 22)
        Me.StatusStrip1.TabIndex = 13
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(119, 17)
        Me.ToolStripStatusLabel1.Text = "ToolStripStatusLabel1"
        '
        'clearUnknownList
        '
        Me.clearUnknownList.Image = Global.ProjectBoardBasic.My.Resources.Resources.selection_delete
        Me.clearUnknownList.Location = New System.Drawing.Point(344, 104)
        Me.clearUnknownList.Name = "clearUnknownList"
        Me.clearUnknownList.Size = New System.Drawing.Size(18, 16)
        Me.clearUnknownList.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.clearUnknownList.TabIndex = 19
        Me.clearUnknownList.TabStop = False
        '
        'setItemToBeUnknown
        '
        Me.setItemToBeUnknown.Enabled = False
        Me.setItemToBeUnknown.Image = Global.ProjectBoardBasic.My.Resources.Resources.Pfeil_links_32x321
        Me.setItemToBeUnknown.Location = New System.Drawing.Point(370, 268)
        Me.setItemToBeUnknown.Name = "setItemToBeUnknown"
        Me.setItemToBeUnknown.Size = New System.Drawing.Size(34, 34)
        Me.setItemToBeUnknown.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.setItemToBeUnknown.TabIndex = 17
        Me.setItemToBeUnknown.TabStop = False
        '
        'setItemToBeKnown
        '
        Me.setItemToBeKnown.Enabled = False
        Me.setItemToBeKnown.Image = Global.ProjectBoardBasic.My.Resources.Resources.Pfeil_rechts_32x321
        Me.setItemToBeKnown.Location = New System.Drawing.Point(370, 176)
        Me.setItemToBeKnown.Name = "setItemToBeKnown"
        Me.setItemToBeKnown.Size = New System.Drawing.Size(34, 34)
        Me.setItemToBeKnown.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.setItemToBeKnown.TabIndex = 16
        Me.setItemToBeKnown.TabStop = False
        '
        'clearStandardList
        '
        Me.clearStandardList.Image = Global.ProjectBoardBasic.My.Resources.Resources.selection_delete
        Me.clearStandardList.Location = New System.Drawing.Point(733, 104)
        Me.clearStandardList.Name = "clearStandardList"
        Me.clearStandardList.Size = New System.Drawing.Size(18, 16)
        Me.clearStandardList.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.clearStandardList.TabIndex = 20
        Me.clearStandardList.TabStop = False
        '
        'showOnlySummaryTasks
        '
        Me.showOnlySummaryTasks.AutoSize = True
        Me.showOnlySummaryTasks.Location = New System.Drawing.Point(23, 102)
        Me.showOnlySummaryTasks.Name = "showOnlySummaryTasks"
        Me.showOnlySummaryTasks.Size = New System.Drawing.Size(172, 17)
        Me.showOnlySummaryTasks.TabIndex = 21
        Me.showOnlySummaryTasks.Text = "nur Sammelvorgänge anzeigen"
        Me.showOnlySummaryTasks.UseVisualStyleBackColor = True
        Me.showOnlySummaryTasks.Visible = False
        '
        'storeButton
        '
        Me.storeButton.Enabled = False
        Me.storeButton.Location = New System.Drawing.Point(676, 502)
        Me.storeButton.Name = "storeButton"
        Me.storeButton.Size = New System.Drawing.Size(75, 23)
        Me.storeButton.TabIndex = 22
        Me.storeButton.Text = "Speichern"
        Me.storeButton.UseVisualStyleBackColor = True
        '
        'filtersAreCoupled
        '
        Me.filtersAreCoupled.AutoSize = True
        Me.filtersAreCoupled.Checked = True
        Me.filtersAreCoupled.CheckState = System.Windows.Forms.CheckState.Checked
        Me.filtersAreCoupled.Location = New System.Drawing.Point(372, 82)
        Me.filtersAreCoupled.Name = "filtersAreCoupled"
        Me.filtersAreCoupled.Size = New System.Drawing.Size(32, 17)
        Me.filtersAreCoupled.TabIndex = 23
        Me.filtersAreCoupled.Text = "="
        Me.filtersAreCoupled.UseVisualStyleBackColor = True
        '
        'frmEditWoerterbuch
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(774, 562)
        Me.Controls.Add(Me.filtersAreCoupled)
        Me.Controls.Add(Me.storeButton)
        Me.Controls.Add(Me.showOnlySummaryTasks)
        Me.Controls.Add(Me.clearStandardList)
        Me.Controls.Add(Me.clearUnknownList)
        Me.Controls.Add(Me.setItemToBeUnknown)
        Me.Controls.Add(Me.setItemToBeKnown)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.visElements)
        Me.Controls.Add(Me.replaceButton)
        Me.Controls.Add(Me.addRulesToDictionary)
        Me.Controls.Add(Me.ignoreButton)
        Me.Controls.Add(Me.editUnknownItem)
        Me.Controls.Add(Me.filterStandard)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.standardList)
        Me.Controls.Add(Me.unknownList)
        Me.Controls.Add(Me.filterUnknown)
        Me.Controls.Add(Me.rdbListShowsMilestones)
        Me.Controls.Add(Me.rdbListShowsPhases)
        Me.HelpButton = True
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmEditWoerterbuch"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Editieren des Wörterbuchs"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        CType(Me.clearUnknownList, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.setItemToBeUnknown, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.setItemToBeKnown, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.clearStandardList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents rdbListShowsPhases As System.Windows.Forms.RadioButton
    Friend WithEvents rdbListShowsMilestones As System.Windows.Forms.RadioButton
    Friend WithEvents filterUnknown As System.Windows.Forms.TextBox
    Friend WithEvents unknownList As System.Windows.Forms.ListBox
    Friend WithEvents standardList As System.Windows.Forms.ListBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents filterStandard As System.Windows.Forms.TextBox
    Friend WithEvents editUnknownItem As System.Windows.Forms.TextBox
    Friend WithEvents ignoreButton As System.Windows.Forms.Button
    Friend WithEvents addRulesToDictionary As System.Windows.Forms.Button
    Friend WithEvents replaceButton As System.Windows.Forms.Button
    Friend WithEvents visElements As System.Windows.Forms.Button
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents setItemToBeKnown As System.Windows.Forms.PictureBox
    Friend WithEvents setItemToBeUnknown As System.Windows.Forms.PictureBox
    Friend WithEvents clearUnknownList As System.Windows.Forms.PictureBox
    Friend WithEvents clearStandardList As System.Windows.Forms.PictureBox
    Friend WithEvents showOnlySummaryTasks As System.Windows.Forms.CheckBox
    Friend WithEvents storeButton As System.Windows.Forms.Button
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents filtersAreCoupled As System.Windows.Forms.CheckBox
End Class
