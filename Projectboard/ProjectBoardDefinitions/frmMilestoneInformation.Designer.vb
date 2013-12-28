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
        Me.phaseName = New System.Windows.Forms.TextBox()
        Me.resultName = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'resultDate
        '
        Me.resultDate.Enabled = False
        Me.resultDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.resultDate.Location = New System.Drawing.Point(310, 82)
        Me.resultDate.Name = "resultDate"
        Me.resultDate.Size = New System.Drawing.Size(119, 29)
        Me.resultDate.TabIndex = 10
        '
        'bewertungsText
        '
        Me.bewertungsText.BackColor = System.Drawing.SystemColors.Window
        Me.bewertungsText.Enabled = False
        Me.bewertungsText.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bewertungsText.Location = New System.Drawing.Point(25, 128)
        Me.bewertungsText.MaximumSize = New System.Drawing.Size(420, 140)
        Me.bewertungsText.MinimumSize = New System.Drawing.Size(420, 140)
        Me.bewertungsText.Multiline = True
        Me.bewertungsText.Name = "bewertungsText"
        Me.bewertungsText.ReadOnly = True
        Me.bewertungsText.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.bewertungsText.Size = New System.Drawing.Size(420, 140)
        Me.bewertungsText.TabIndex = 19
        '
        'projectName
        '
        Me.projectName.BackColor = System.Drawing.SystemColors.Window
        Me.projectName.Enabled = False
        Me.projectName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.projectName.Location = New System.Drawing.Point(25, 23)
        Me.projectName.Name = "projectName"
        Me.projectName.ReadOnly = True
        Me.projectName.Size = New System.Drawing.Size(404, 22)
        Me.projectName.TabIndex = 20
        '
        'phaseName
        '
        Me.phaseName.BackColor = System.Drawing.SystemColors.Window
        Me.phaseName.Enabled = False
        Me.phaseName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.phaseName.Location = New System.Drawing.Point(25, 53)
        Me.phaseName.Name = "phaseName"
        Me.phaseName.ReadOnly = True
        Me.phaseName.Size = New System.Drawing.Size(279, 22)
        Me.phaseName.TabIndex = 22
        '
        'resultName
        '
        Me.resultName.BackColor = System.Drawing.SystemColors.Window
        Me.resultName.Enabled = False
        Me.resultName.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.resultName.Location = New System.Drawing.Point(25, 82)
        Me.resultName.Name = "resultName"
        Me.resultName.ReadOnly = True
        Me.resultName.Size = New System.Drawing.Size(279, 29)
        Me.resultName.TabIndex = 24
        '
        'frmMilestoneInformation
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Window
        Me.ClientSize = New System.Drawing.Size(463, 299)
        Me.Controls.Add(Me.resultName)
        Me.Controls.Add(Me.phaseName)
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
    Public WithEvents phaseName As System.Windows.Forms.TextBox
    Public WithEvents resultName As System.Windows.Forms.TextBox
    Public WithEvents resultDate As System.Windows.Forms.TextBox
    Public WithEvents bewertungsText As System.Windows.Forms.TextBox
End Class
