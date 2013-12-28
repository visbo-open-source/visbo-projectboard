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
        Me.ListBox1.Location = New System.Drawing.Point(20, 22)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(370, 225)
        Me.ListBox1.TabIndex = 0
        '
        'compPhases
        '
        Me.compPhases.AutoSize = True
        Me.compPhases.Checked = True
        Me.compPhases.Location = New System.Drawing.Point(20, 257)
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
        Me.compResources.Location = New System.Drawing.Point(20, 278)
        Me.compResources.Name = "compResources"
        Me.compResources.Size = New System.Drawing.Size(122, 17)
        Me.compResources.TabIndex = 2
        Me.compResources.Text = "Ressourcen Bedarfe"
        Me.compResources.UseVisualStyleBackColor = True
        '
        'OKButton
        '
        Me.OKButton.Location = New System.Drawing.Point(106, 306)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(75, 23)
        Me.OKButton.TabIndex = 3
        Me.OKButton.Text = "OK"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'AbbrButton
        '
        Me.AbbrButton.Location = New System.Drawing.Point(208, 306)
        Me.AbbrButton.Name = "AbbrButton"
        Me.AbbrButton.Size = New System.Drawing.Size(75, 23)
        Me.AbbrButton.TabIndex = 4
        Me.AbbrButton.Text = "Abbrechen"
        Me.AbbrButton.UseVisualStyleBackColor = True
        '
        'frmSelectCompareProject
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(402, 347)
        Me.Controls.Add(Me.AbbrButton)
        Me.Controls.Add(Me.OKButton)
        Me.Controls.Add(Me.compResources)
        Me.Controls.Add(Me.compPhases)
        Me.Controls.Add(Me.ListBox1)
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
