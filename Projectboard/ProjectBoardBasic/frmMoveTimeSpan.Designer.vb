<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMoveTimeSpan
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
        Me.moveToLeft = New System.Windows.Forms.Button()
        Me.moveToRight = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'moveToLeft
        '
        Me.moveToLeft.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.moveToLeft.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.moveToLeft.Location = New System.Drawing.Point(73, 30)
        Me.moveToLeft.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.moveToLeft.Name = "moveToLeft"
        Me.moveToLeft.Size = New System.Drawing.Size(100, 58)
        Me.moveToLeft.TabIndex = 0
        Me.moveToLeft.Text = "<"
        Me.moveToLeft.UseVisualStyleBackColor = True
        '
        'moveToRight
        '
        Me.moveToRight.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.moveToRight.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.moveToRight.Location = New System.Drawing.Point(199, 30)
        Me.moveToRight.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.moveToRight.Name = "moveToRight"
        Me.moveToRight.Size = New System.Drawing.Size(100, 58)
        Me.moveToRight.TabIndex = 1
        Me.moveToRight.Text = ">"
        Me.moveToRight.UseVisualStyleBackColor = True
        '
        'frmMoveTimeSpan
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(120.0!, 120.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.ClientSize = New System.Drawing.Size(379, 116)
        Me.Controls.Add(Me.moveToRight)
        Me.Controls.Add(Me.moveToLeft)
        Me.Location = New System.Drawing.Point(100, 8)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmMoveTimeSpan"
        Me.Text = "Zeitraum verschieben"
        Me.TopMost = True
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents moveToLeft As System.Windows.Forms.Button
    Friend WithEvents moveToRight As System.Windows.Forms.Button
End Class
