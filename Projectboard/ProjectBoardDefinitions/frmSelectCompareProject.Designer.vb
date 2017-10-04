<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSelectCompareProject
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSelectCompareProject))
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.compPhases = New System.Windows.Forms.RadioButton()
        Me.compResources = New System.Windows.Forms.RadioButton()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.AbbrButton = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'ListBox1
        '
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Location = New System.Drawing.Point(22, 22)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(394, 212)
        Me.ListBox1.TabIndex = 0
        '
        'compPhases
        '
        Me.compPhases.AutoSize = True
        Me.compPhases.Checked = True
        Me.compPhases.Location = New System.Drawing.Point(22, 253)
        Me.compPhases.Name = "compPhases"
        Me.compPhases.Size = New System.Drawing.Size(128, 17)
        Me.compPhases.TabIndex = 1
        Me.compPhases.TabStop = True
        Me.compPhases.Text = "Phasen Charakteristik"
        Me.compPhases.UseVisualStyleBackColor = True
        '
        'compResources
        '
        Me.compResources.AutoSize = True
        Me.compResources.Location = New System.Drawing.Point(22, 274)
        Me.compResources.Name = "compResources"
        Me.compResources.Size = New System.Drawing.Size(122, 17)
        Me.compResources.TabIndex = 2
        Me.compResources.Text = "Ressourcen Bedarfe"
        Me.compResources.UseVisualStyleBackColor = True
        '
        'OKButton
        '
        Me.OKButton.Location = New System.Drawing.Point(113, 302)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(80, 22)
        Me.OKButton.TabIndex = 3
        Me.OKButton.Text = "OK"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'AbbrButton
        '
        Me.AbbrButton.Location = New System.Drawing.Point(222, 302)
        Me.AbbrButton.Name = "AbbrButton"
        Me.AbbrButton.Size = New System.Drawing.Size(80, 22)
        Me.AbbrButton.TabIndex = 4
        Me.AbbrButton.Text = "Abbrechen"
        Me.AbbrButton.UseVisualStyleBackColor = True
        '
        'frmSelectCompareProject
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.ClientSize = New System.Drawing.Size(429, 342)
        Me.Controls.Add(Me.AbbrButton)
        Me.Controls.Add(Me.OKButton)
        Me.Controls.Add(Me.compResources)
        Me.Controls.Add(Me.compPhases)
        Me.Controls.Add(Me.ListBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmSelectCompareProject"
        Me.Text = "zu vergleichendes Projekt wählen"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents ListBox1 As System.Windows.Forms.ListBox
    Public WithEvents compPhases As System.Windows.Forms.RadioButton
    Public WithEvents compResources As System.Windows.Forms.RadioButton
    Public WithEvents OKButton As System.Windows.Forms.Button
    Public WithEvents AbbrButton As System.Windows.Forms.Button
End Class
