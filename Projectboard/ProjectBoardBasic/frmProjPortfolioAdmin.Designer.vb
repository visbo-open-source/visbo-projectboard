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
        Me.TreeViewProjekte = New System.Windows.Forms.TreeView()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.txtboxLabel = New System.Windows.Forms.Label()
        Me.txtDropbox = New System.Windows.Forms.ComboBox()
        Me.considerDependencies = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'TreeViewProjekte
        '
        Me.TreeViewProjekte.Location = New System.Drawing.Point(34, 25)
        Me.TreeViewProjekte.Margin = New System.Windows.Forms.Padding(2)
        Me.TreeViewProjekte.Name = "TreeViewProjekte"
        Me.TreeViewProjekte.Size = New System.Drawing.Size(395, 290)
        Me.TreeViewProjekte.TabIndex = 1
        '
        'OKButton
        '
        Me.OKButton.Location = New System.Drawing.Point(175, 389)
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
        Me.txtboxLabel.Location = New System.Drawing.Point(31, 327)
        Me.txtboxLabel.Name = "txtboxLabel"
        Me.txtboxLabel.Size = New System.Drawing.Size(72, 13)
        Me.txtboxLabel.TabIndex = 33
        Me.txtboxLabel.Text = "Filter-Auswahl"
        '
        'txtDropbox
        '
        Me.txtDropbox.FormattingEnabled = True
        Me.txtDropbox.Location = New System.Drawing.Point(34, 350)
        Me.txtDropbox.Name = "txtDropbox"
        Me.txtDropbox.Size = New System.Drawing.Size(395, 21)
        Me.txtDropbox.TabIndex = 34
        '
        'considerDependencies
        '
        Me.considerDependencies.AutoSize = True
        Me.considerDependencies.Location = New System.Drawing.Point(251, 326)
        Me.considerDependencies.Name = "considerDependencies"
        Me.considerDependencies.Size = New System.Drawing.Size(178, 17)
        Me.considerDependencies.TabIndex = 35
        Me.considerDependencies.Text = "Abhängigkeiten berücksichtigen"
        Me.considerDependencies.UseVisualStyleBackColor = True
        Me.considerDependencies.Visible = False
        '
        'frmProjPortfolioAdmin
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(467, 430)
        Me.Controls.Add(Me.considerDependencies)
        Me.Controls.Add(Me.txtDropbox)
        Me.Controls.Add(Me.txtboxLabel)
        Me.Controls.Add(Me.OKButton)
        Me.Controls.Add(Me.TreeViewProjekte)
        Me.Name = "frmProjPortfolioAdmin"
        Me.Text = "Multiprojekt-Szenario"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents TreeViewProjekte As System.Windows.Forms.TreeView
    Friend WithEvents txtDropbox As System.Windows.Forms.ComboBox
    Friend WithEvents OKButton As System.Windows.Forms.Button
    Public WithEvents txtboxLabel As System.Windows.Forms.Label
    Friend WithEvents considerDependencies As System.Windows.Forms.CheckBox
End Class
