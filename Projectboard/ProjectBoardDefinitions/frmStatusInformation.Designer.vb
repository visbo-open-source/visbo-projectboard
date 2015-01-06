<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmStatusInformation
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
        Me.bewertungsText = New System.Windows.Forms.TextBox()
        Me.projectName = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'bewertungsText
        '
        Me.bewertungsText.BackColor = System.Drawing.SystemColors.Window
        Me.bewertungsText.Enabled = False
        Me.bewertungsText.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bewertungsText.Location = New System.Drawing.Point(23, 67)
        Me.bewertungsText.MaximumSize = New System.Drawing.Size(448, 138)
        Me.bewertungsText.MinimumSize = New System.Drawing.Size(320, 79)
        Me.bewertungsText.Multiline = True
        Me.bewertungsText.Name = "bewertungsText"
        Me.bewertungsText.ReadOnly = True
        Me.bewertungsText.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.bewertungsText.Size = New System.Drawing.Size(448, 138)
        Me.bewertungsText.TabIndex = 30
        '
        'projectName
        '
        Me.projectName.BackColor = System.Drawing.SystemColors.Window
        Me.projectName.Enabled = False
        Me.projectName.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.projectName.Location = New System.Drawing.Point(25, 22)
        Me.projectName.Name = "projectName"
        Me.projectName.ReadOnly = True
        Me.projectName.ShortcutsEnabled = False
        Me.projectName.Size = New System.Drawing.Size(446, 29)
        Me.projectName.TabIndex = 34
        '
        'frmStatusInformation
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoSize = True
        Me.BackColor = System.Drawing.SystemColors.Window
        Me.ClientSize = New System.Drawing.Size(494, 234)
        Me.Controls.Add(Me.projectName)
        Me.Controls.Add(Me.bewertungsText)
        Me.Name = "frmStatusInformation"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Status Information"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents bewertungsText As System.Windows.Forms.TextBox
    Public WithEvents projectName As System.Windows.Forms.TextBox
End Class
