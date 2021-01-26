<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMppSettings
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMppSettings))
        Me.shwPhaseText = New System.Windows.Forms.CheckBox()
        Me.shwPhaseDate = New System.Windows.Forms.CheckBox()
        Me.shwProjectLine = New System.Windows.Forms.CheckBox()
        Me.ShwMilestoneDate = New System.Windows.Forms.CheckBox()
        Me.ShwMilestoneText = New System.Windows.Forms.CheckBox()
        Me.shwAmpeln = New System.Windows.Forms.CheckBox()
        Me.shwLegend = New System.Windows.Forms.CheckBox()
        Me.shwVerticals = New System.Windows.Forms.CheckBox()
        Me.notStrictly = New System.Windows.Forms.CheckBox()
        Me.okButton = New System.Windows.Forms.Button()
        Me.allOnOnePage = New System.Windows.Forms.CheckBox()
        Me.sortiertNachDauer = New System.Windows.Forms.CheckBox()
        Me.shwExtendedMode = New System.Windows.Forms.CheckBox()
        Me.useAbbrev = New System.Windows.Forms.CheckBox()
        Me.shwHorizontals = New System.Windows.Forms.CheckBox()
        Me.KwInMilestone = New System.Windows.Forms.CheckBox()
        Me.useOriginalNames = New System.Windows.Forms.CheckBox()
        Me.filterEmptyProjects = New System.Windows.Forms.CheckBox()
        Me.shwInvoices = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'shwPhaseText
        '
        Me.shwPhaseText.AutoSize = True
        Me.shwPhaseText.Location = New System.Drawing.Point(14, 105)
        Me.shwPhaseText.Name = "shwPhaseText"
        Me.shwPhaseText.Size = New System.Drawing.Size(124, 17)
        Me.shwPhaseText.TabIndex = 0
        Me.shwPhaseText.Text = "Phasen Beschriftung"
        Me.shwPhaseText.UseVisualStyleBackColor = True
        '
        'shwPhaseDate
        '
        Me.shwPhaseDate.AutoSize = True
        Me.shwPhaseDate.Location = New System.Drawing.Point(14, 129)
        Me.shwPhaseDate.Name = "shwPhaseDate"
        Me.shwPhaseDate.Size = New System.Drawing.Size(96, 17)
        Me.shwPhaseDate.TabIndex = 1
        Me.shwPhaseDate.Text = "Phasen Datum"
        Me.shwPhaseDate.UseVisualStyleBackColor = True
        '
        'shwProjectLine
        '
        Me.shwProjectLine.AutoSize = True
        Me.shwProjectLine.Location = New System.Drawing.Point(14, 21)
        Me.shwProjectLine.Name = "shwProjectLine"
        Me.shwProjectLine.Size = New System.Drawing.Size(77, 17)
        Me.shwProjectLine.TabIndex = 3
        Me.shwProjectLine.Text = "Projektlinie"
        Me.shwProjectLine.UseVisualStyleBackColor = True
        '
        'ShwMilestoneDate
        '
        Me.ShwMilestoneDate.AutoSize = True
        Me.ShwMilestoneDate.Location = New System.Drawing.Point(181, 129)
        Me.ShwMilestoneDate.Name = "ShwMilestoneDate"
        Me.ShwMilestoneDate.Size = New System.Drawing.Size(113, 17)
        Me.ShwMilestoneDate.TabIndex = 5
        Me.ShwMilestoneDate.Text = "Meilenstein Datum"
        Me.ShwMilestoneDate.UseVisualStyleBackColor = True
        '
        'ShwMilestoneText
        '
        Me.ShwMilestoneText.AutoSize = True
        Me.ShwMilestoneText.Location = New System.Drawing.Point(181, 105)
        Me.ShwMilestoneText.Name = "ShwMilestoneText"
        Me.ShwMilestoneText.Size = New System.Drawing.Size(141, 17)
        Me.ShwMilestoneText.TabIndex = 4
        Me.ShwMilestoneText.Text = "Meilenstein Beschriftung"
        Me.ShwMilestoneText.UseVisualStyleBackColor = True
        '
        'shwAmpeln
        '
        Me.shwAmpeln.AutoSize = True
        Me.shwAmpeln.Location = New System.Drawing.Point(181, 21)
        Me.shwAmpeln.Name = "shwAmpeln"
        Me.shwAmpeln.Size = New System.Drawing.Size(61, 17)
        Me.shwAmpeln.TabIndex = 6
        Me.shwAmpeln.Text = "Ampeln"
        Me.shwAmpeln.UseVisualStyleBackColor = True
        '
        'shwLegend
        '
        Me.shwLegend.AutoSize = True
        Me.shwLegend.Location = New System.Drawing.Point(14, 219)
        Me.shwLegend.Name = "shwLegend"
        Me.shwLegend.Size = New System.Drawing.Size(114, 17)
        Me.shwLegend.TabIndex = 8
        Me.shwLegend.Text = "Legende anzeigen"
        Me.shwLegend.UseVisualStyleBackColor = True
        '
        'shwVerticals
        '
        Me.shwVerticals.AutoSize = True
        Me.shwVerticals.Location = New System.Drawing.Point(14, 196)
        Me.shwVerticals.Name = "shwVerticals"
        Me.shwVerticals.Size = New System.Drawing.Size(98, 17)
        Me.shwVerticals.TabIndex = 7
        Me.shwVerticals.Text = "Vertikale Linien"
        Me.shwVerticals.UseVisualStyleBackColor = True
        '
        'notStrictly
        '
        Me.notStrictly.AutoSize = True
        Me.notStrictly.Location = New System.Drawing.Point(14, 44)
        Me.notStrictly.Name = "notStrictly"
        Me.notStrictly.Size = New System.Drawing.Size(295, 17)
        Me.notStrictly.TabIndex = 9
        Me.notStrictly.Text = "ein Planelement im Zeitraum: alle anderen auch anzeigen"
        Me.notStrictly.UseVisualStyleBackColor = True
        '
        'okButton
        '
        Me.okButton.Location = New System.Drawing.Point(132, 281)
        Me.okButton.Name = "okButton"
        Me.okButton.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.okButton.Size = New System.Drawing.Size(75, 23)
        Me.okButton.TabIndex = 10
        Me.okButton.Text = "OK"
        Me.okButton.UseVisualStyleBackColor = True
        '
        'allOnOnePage
        '
        Me.allOnOnePage.AutoSize = True
        Me.allOnOnePage.Location = New System.Drawing.Point(181, 219)
        Me.allOnOnePage.Name = "allOnOnePage"
        Me.allOnOnePage.Size = New System.Drawing.Size(91, 17)
        Me.allOnOnePage.TabIndex = 11
        Me.allOnOnePage.Text = "auf eine Seite"
        Me.allOnOnePage.UseVisualStyleBackColor = True
        '
        'sortiertNachDauer
        '
        Me.sortiertNachDauer.AutoSize = True
        Me.sortiertNachDauer.Location = New System.Drawing.Point(14, 242)
        Me.sortiertNachDauer.Name = "sortiertNachDauer"
        Me.sortiertNachDauer.Size = New System.Drawing.Size(116, 17)
        Me.sortiertNachDauer.TabIndex = 12
        Me.sortiertNachDauer.Text = "sortiert nach Dauer"
        Me.sortiertNachDauer.UseVisualStyleBackColor = True
        '
        'shwExtendedMode
        '
        Me.shwExtendedMode.AutoSize = True
        Me.shwExtendedMode.Location = New System.Drawing.Point(181, 242)
        Me.shwExtendedMode.Name = "shwExtendedMode"
        Me.shwExtendedMode.Size = New System.Drawing.Size(101, 17)
        Me.shwExtendedMode.TabIndex = 13
        Me.shwExtendedMode.Text = "Extended Mode"
        Me.shwExtendedMode.UseVisualStyleBackColor = True
        '
        'useAbbrev
        '
        Me.useAbbrev.AutoSize = True
        Me.useAbbrev.Location = New System.Drawing.Point(14, 152)
        Me.useAbbrev.Name = "useAbbrev"
        Me.useAbbrev.Size = New System.Drawing.Size(133, 17)
        Me.useAbbrev.TabIndex = 14
        Me.useAbbrev.Text = "Abkürzung verwenden"
        Me.useAbbrev.UseVisualStyleBackColor = True
        '
        'shwHorizontals
        '
        Me.shwHorizontals.AutoSize = True
        Me.shwHorizontals.Location = New System.Drawing.Point(181, 196)
        Me.shwHorizontals.Name = "shwHorizontals"
        Me.shwHorizontals.Size = New System.Drawing.Size(110, 17)
        Me.shwHorizontals.TabIndex = 15
        Me.shwHorizontals.Text = "Horizontale Linien"
        Me.shwHorizontals.UseVisualStyleBackColor = True
        '
        'KwInMilestone
        '
        Me.KwInMilestone.AutoSize = True
        Me.KwInMilestone.Location = New System.Drawing.Point(181, 152)
        Me.KwInMilestone.Name = "KwInMilestone"
        Me.KwInMilestone.Size = New System.Drawing.Size(113, 17)
        Me.KwInMilestone.TabIndex = 16
        Me.KwInMilestone.Text = "KW im Meilenstein"
        Me.KwInMilestone.UseVisualStyleBackColor = True
        '
        'useOriginalNames
        '
        Me.useOriginalNames.AutoSize = True
        Me.useOriginalNames.Location = New System.Drawing.Point(14, 67)
        Me.useOriginalNames.Name = "useOriginalNames"
        Me.useOriginalNames.Size = New System.Drawing.Size(154, 17)
        Me.useOriginalNames.TabIndex = 17
        Me.useOriginalNames.Text = "Original-Namen verwenden"
        Me.useOriginalNames.UseVisualStyleBackColor = True
        '
        'filterEmptyProjects
        '
        Me.filterEmptyProjects.AutoSize = True
        Me.filterEmptyProjects.Location = New System.Drawing.Point(181, 67)
        Me.filterEmptyProjects.Name = "filterEmptyProjects"
        Me.filterEmptyProjects.Size = New System.Drawing.Size(133, 17)
        Me.filterEmptyProjects.TabIndex = 18
        Me.filterEmptyProjects.Text = """Leere"" Projekte filtern"
        Me.filterEmptyProjects.UseVisualStyleBackColor = True
        '
        'shwInvoices
        '
        Me.shwInvoices.AutoSize = True
        Me.shwInvoices.Location = New System.Drawing.Point(181, 173)
        Me.shwInvoices.Name = "shwInvoices"
        Me.shwInvoices.Size = New System.Drawing.Size(122, 17)
        Me.shwInvoices.TabIndex = 19
        Me.shwInvoices.Text = "Rechnung / Penalty"
        Me.shwInvoices.UseVisualStyleBackColor = True
        '
        'frmMppSettings
        '
        Me.AcceptButton = Me.okButton
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(328, 329)
        Me.Controls.Add(Me.shwInvoices)
        Me.Controls.Add(Me.filterEmptyProjects)
        Me.Controls.Add(Me.useOriginalNames)
        Me.Controls.Add(Me.KwInMilestone)
        Me.Controls.Add(Me.shwHorizontals)
        Me.Controls.Add(Me.useAbbrev)
        Me.Controls.Add(Me.shwExtendedMode)
        Me.Controls.Add(Me.sortiertNachDauer)
        Me.Controls.Add(Me.allOnOnePage)
        Me.Controls.Add(Me.okButton)
        Me.Controls.Add(Me.notStrictly)
        Me.Controls.Add(Me.shwLegend)
        Me.Controls.Add(Me.shwVerticals)
        Me.Controls.Add(Me.shwAmpeln)
        Me.Controls.Add(Me.ShwMilestoneDate)
        Me.Controls.Add(Me.ShwMilestoneText)
        Me.Controls.Add(Me.shwProjectLine)
        Me.Controls.Add(Me.shwPhaseDate)
        Me.Controls.Add(Me.shwPhaseText)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmMppSettings"
        Me.Text = "Einstellungen"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents shwPhaseText As System.Windows.Forms.CheckBox
    Friend WithEvents shwPhaseDate As System.Windows.Forms.CheckBox
    Friend WithEvents shwProjectLine As System.Windows.Forms.CheckBox
    Friend WithEvents ShwMilestoneDate As System.Windows.Forms.CheckBox
    Friend WithEvents ShwMilestoneText As System.Windows.Forms.CheckBox
    Friend WithEvents shwAmpeln As System.Windows.Forms.CheckBox
    Friend WithEvents shwLegend As System.Windows.Forms.CheckBox
    Friend WithEvents shwVerticals As System.Windows.Forms.CheckBox
    Friend WithEvents notStrictly As System.Windows.Forms.CheckBox
    Friend WithEvents okButton As System.Windows.Forms.Button
    Friend WithEvents allOnOnePage As System.Windows.Forms.CheckBox
    Friend WithEvents sortiertNachDauer As System.Windows.Forms.CheckBox
    Friend WithEvents shwExtendedMode As System.Windows.Forms.CheckBox
    Friend WithEvents useAbbrev As System.Windows.Forms.CheckBox
    Friend WithEvents shwHorizontals As System.Windows.Forms.CheckBox
    Friend WithEvents KwInMilestone As System.Windows.Forms.CheckBox
    Friend WithEvents useOriginalNames As System.Windows.Forms.CheckBox
    Friend WithEvents filterEmptyProjects As System.Windows.Forms.CheckBox
    Friend WithEvents shwInvoices As Windows.Forms.CheckBox
End Class
