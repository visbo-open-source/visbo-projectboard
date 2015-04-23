<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMilestoneInformation
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
        Me.resultDate = New System.Windows.Forms.TextBox()
        Me.bewertungsText = New System.Windows.Forms.TextBox()
        Me.projectName = New System.Windows.Forms.TextBox()
        Me.resultName = New System.Windows.Forms.TextBox()
        Me.breadCrumb = New System.Windows.Forms.TextBox()
        Me.lfdNr = New System.Windows.Forms.TextBox()
        Me.showOrigItem = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'resultDate
        '
        Me.resultDate.Enabled = False
        Me.resultDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.resultDate.Location = New System.Drawing.Point(362, 59)
        Me.resultDate.Name = "resultDate"
        Me.resultDate.ReadOnly = True
        Me.resultDate.Size = New System.Drawing.Size(112, 26)
        Me.resultDate.TabIndex = 10
        '
        'bewertungsText
        '
        Me.bewertungsText.BackColor = System.Drawing.SystemColors.Control
        Me.bewertungsText.Enabled = False
        Me.bewertungsText.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bewertungsText.Location = New System.Drawing.Point(26, 117)
        Me.bewertungsText.MaximumSize = New System.Drawing.Size(448, 138)
        Me.bewertungsText.MinimumSize = New System.Drawing.Size(448, 138)
        Me.bewertungsText.Multiline = True
        Me.bewertungsText.Name = "bewertungsText"
        Me.bewertungsText.ReadOnly = True
        Me.bewertungsText.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.bewertungsText.Size = New System.Drawing.Size(448, 138)
        Me.bewertungsText.TabIndex = 19
        '
        'projectName
        '
        Me.projectName.BackColor = System.Drawing.SystemColors.Control
        Me.projectName.Enabled = False
        Me.projectName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.projectName.Location = New System.Drawing.Point(26, 22)
        Me.projectName.Name = "projectName"
        Me.projectName.ReadOnly = True
        Me.projectName.Size = New System.Drawing.Size(141, 20)
        Me.projectName.TabIndex = 20
        '
        'resultName
        '
        Me.resultName.BackColor = System.Drawing.SystemColors.Control
        Me.resultName.Enabled = False
        Me.resultName.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.resultName.Location = New System.Drawing.Point(26, 59)
        Me.resultName.Name = "resultName"
        Me.resultName.ReadOnly = True
        Me.resultName.Size = New System.Drawing.Size(285, 26)
        Me.resultName.TabIndex = 24
        '
        'breadCrumb
        '
        Me.breadCrumb.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.breadCrumb.Location = New System.Drawing.Point(182, 22)
        Me.breadCrumb.Name = "breadCrumb"
        Me.breadCrumb.ReadOnly = True
        Me.breadCrumb.Size = New System.Drawing.Size(292, 20)
        Me.breadCrumb.TabIndex = 25
        '
        'lfdNr
        '
        Me.lfdNr.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lfdNr.Location = New System.Drawing.Point(315, 59)
        Me.lfdNr.Name = "lfdNr"
        Me.lfdNr.ReadOnly = True
        Me.lfdNr.Size = New System.Drawing.Size(42, 26)
        Me.lfdNr.TabIndex = 26
        '
        'showOrigItem
        '
        Me.showOrigItem.AutoSize = True
        Me.showOrigItem.Location = New System.Drawing.Point(26, 91)
        Me.showOrigItem.Name = "showOrigItem"
        Me.showOrigItem.Size = New System.Drawing.Size(92, 17)
        Me.showOrigItem.TabIndex = 27
        Me.showOrigItem.Text = "Original Name"
        Me.showOrigItem.UseVisualStyleBackColor = True
        '
        'frmMilestoneInformation
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoSize = True
        Me.BackColor = System.Drawing.SystemColors.Window
        Me.ClientSize = New System.Drawing.Size(494, 277)
        Me.Controls.Add(Me.showOrigItem)
        Me.Controls.Add(Me.lfdNr)
        Me.Controls.Add(Me.breadCrumb)
        Me.Controls.Add(Me.resultName)
        Me.Controls.Add(Me.projectName)
        Me.Controls.Add(Me.bewertungsText)
        Me.Controls.Add(Me.resultDate)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmMilestoneInformation"
        Me.Text = "Meilenstein Informationen"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents projectName As System.Windows.Forms.TextBox
    Public WithEvents resultName As System.Windows.Forms.TextBox
    Public WithEvents resultDate As System.Windows.Forms.TextBox
    Public WithEvents bewertungsText As System.Windows.Forms.TextBox
    Public WithEvents breadCrumb As System.Windows.Forms.TextBox
    Public WithEvents lfdNr As System.Windows.Forms.TextBox
    Public WithEvents showOrigItem As System.Windows.Forms.CheckBox
End Class
