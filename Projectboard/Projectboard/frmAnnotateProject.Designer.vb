<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAnnotateProject
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
        Me.annotatePhases = New System.Windows.Forms.CheckBox()
        Me.annotateMilestones = New System.Windows.Forms.CheckBox()
        Me.showStdNames = New System.Windows.Forms.RadioButton()
        Me.showOrigNames = New System.Windows.Forms.RadioButton()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.showAbbrev = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'annotatePhases
        '
        Me.annotatePhases.AutoSize = True
        Me.annotatePhases.Checked = True
        Me.annotatePhases.CheckState = System.Windows.Forms.CheckState.Checked
        Me.annotatePhases.Location = New System.Drawing.Point(23, 23)
        Me.annotatePhases.Name = "annotatePhases"
        Me.annotatePhases.Size = New System.Drawing.Size(117, 17)
        Me.annotatePhases.TabIndex = 0
        Me.annotatePhases.Text = "Phasen beschriften"
        Me.annotatePhases.UseVisualStyleBackColor = True
        '
        'annotateMilestones
        '
        Me.annotateMilestones.AutoSize = True
        Me.annotateMilestones.Checked = True
        Me.annotateMilestones.CheckState = System.Windows.Forms.CheckState.Checked
        Me.annotateMilestones.Location = New System.Drawing.Point(23, 46)
        Me.annotateMilestones.Name = "annotateMilestones"
        Me.annotateMilestones.Size = New System.Drawing.Size(140, 17)
        Me.annotateMilestones.TabIndex = 1
        Me.annotateMilestones.Text = "Meilensteine beschriften"
        Me.annotateMilestones.UseVisualStyleBackColor = True
        '
        'showStdNames
        '
        Me.showStdNames.AutoSize = True
        Me.showStdNames.Checked = True
        Me.showStdNames.Location = New System.Drawing.Point(23, 80)
        Me.showStdNames.Name = "showStdNames"
        Me.showStdNames.Size = New System.Drawing.Size(105, 17)
        Me.showStdNames.TabIndex = 2
        Me.showStdNames.TabStop = True
        Me.showStdNames.Text = "Standard Namen"
        Me.showStdNames.UseVisualStyleBackColor = True
        '
        'showOrigNames
        '
        Me.showOrigNames.AutoSize = True
        Me.showOrigNames.Location = New System.Drawing.Point(131, 80)
        Me.showOrigNames.Name = "showOrigNames"
        Me.showOrigNames.Size = New System.Drawing.Size(97, 17)
        Me.showOrigNames.TabIndex = 3
        Me.showOrigNames.Text = "Original Namen"
        Me.showOrigNames.UseVisualStyleBackColor = True
        '
        'OKButton
        '
        Me.OKButton.Location = New System.Drawing.Point(52, 138)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(131, 23)
        Me.OKButton.TabIndex = 4
        Me.OKButton.Text = "Beschriften"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'showAbbrev
        '
        Me.showAbbrev.AutoSize = True
        Me.showAbbrev.Location = New System.Drawing.Point(23, 104)
        Me.showAbbrev.Name = "showAbbrev"
        Me.showAbbrev.Size = New System.Drawing.Size(129, 17)
        Me.showAbbrev.TabIndex = 5
        Me.showAbbrev.Text = "nur Abkürzung zeigen"
        Me.showAbbrev.UseVisualStyleBackColor = True
        '
        'frmAnnotateProject
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(233, 182)
        Me.Controls.Add(Me.showAbbrev)
        Me.Controls.Add(Me.OKButton)
        Me.Controls.Add(Me.showOrigNames)
        Me.Controls.Add(Me.showStdNames)
        Me.Controls.Add(Me.annotateMilestones)
        Me.Controls.Add(Me.annotatePhases)
        Me.Name = "frmAnnotateProject"
        Me.Text = "Projekt beschriften"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents annotatePhases As System.Windows.Forms.CheckBox
    Friend WithEvents annotateMilestones As System.Windows.Forms.CheckBox
    Friend WithEvents showStdNames As System.Windows.Forms.RadioButton
    Friend WithEvents showOrigNames As System.Windows.Forms.RadioButton
    Friend WithEvents OKButton As System.Windows.Forms.Button
    Friend WithEvents showAbbrev As System.Windows.Forms.CheckBox
End Class
