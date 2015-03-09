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
        Me.portfolioName = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.applyFilter = New System.Windows.Forms.CheckBox()
        Me.defineFilter = New System.Windows.Forms.LinkLabel()
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
        'portfolioName
        '
        Me.portfolioName.FormattingEnabled = True
        Me.portfolioName.Location = New System.Drawing.Point(127, 344)
        Me.portfolioName.Name = "portfolioName"
        Me.portfolioName.Size = New System.Drawing.Size(302, 21)
        Me.portfolioName.TabIndex = 4
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(34, 348)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(76, 13)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Portfolio Name"
        '
        'OKButton
        '
        Me.OKButton.Location = New System.Drawing.Point(179, 385)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(107, 23)
        Me.OKButton.TabIndex = 6
        Me.OKButton.Text = "Button1"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'applyFilter
        '
        Me.applyFilter.AutoSize = True
        Me.applyFilter.Location = New System.Drawing.Point(186, 321)
        Me.applyFilter.Name = "applyFilter"
        Me.applyFilter.Size = New System.Drawing.Size(101, 17)
        Me.applyFilter.TabIndex = 7
        Me.applyFilter.Text = "Filter anwenden"
        Me.applyFilter.UseVisualStyleBackColor = True
        '
        'defineFilter
        '
        Me.defineFilter.AutoSize = True
        Me.defineFilter.Location = New System.Drawing.Point(347, 321)
        Me.defineFilter.Name = "defineFilter"
        Me.defineFilter.Size = New System.Drawing.Size(78, 13)
        Me.defineFilter.TabIndex = 8
        Me.defineFilter.TabStop = True
        Me.defineFilter.Text = "Filter definieren"
        '
        'frmProjPortfolioAdmin
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(467, 430)
        Me.Controls.Add(Me.defineFilter)
        Me.Controls.Add(Me.applyFilter)
        Me.Controls.Add(Me.OKButton)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.portfolioName)
        Me.Controls.Add(Me.TreeViewProjekte)
        Me.Name = "frmProjPortfolioAdmin"
        Me.Text = "Portfolio erstellen"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents TreeViewProjekte As System.Windows.Forms.TreeView
    Friend WithEvents portfolioName As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents OKButton As System.Windows.Forms.Button
    Friend WithEvents applyFilter As System.Windows.Forms.CheckBox
    Friend WithEvents defineFilter As System.Windows.Forms.LinkLabel
End Class
