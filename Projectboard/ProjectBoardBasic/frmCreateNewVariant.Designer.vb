<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCreateNewVariant
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.infoText = New System.Windows.Forms.Label()
        Me.newVariant = New System.Windows.Forms.TextBox()
        Me.projektName = New System.Windows.Forms.TextBox()
        Me.variantenName = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 25)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(78, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Neue Variante:"
        '
        'infoText
        '
        Me.infoText.AutoSize = True
        Me.infoText.Location = New System.Drawing.Point(12, 81)
        Me.infoText.Name = "infoText"
        Me.infoText.Size = New System.Drawing.Size(319, 13)
        Me.infoText.TabIndex = 1
        Me.infoText.Text = "Die neue Variante wird auf Basis dieser Projekt-Variante angelegt: "
        '
        'newVariant
        '
        Me.newVariant.Location = New System.Drawing.Point(110, 22)
        Me.newVariant.Name = "newVariant"
        Me.newVariant.Size = New System.Drawing.Size(211, 20)
        Me.newVariant.TabIndex = 2
        '
        'projektName
        '
        Me.projektName.Enabled = False
        Me.projektName.Location = New System.Drawing.Point(110, 112)
        Me.projektName.Name = "projektName"
        Me.projektName.Size = New System.Drawing.Size(211, 20)
        Me.projektName.TabIndex = 3
        '
        'variantenName
        '
        Me.variantenName.Enabled = False
        Me.variantenName.Location = New System.Drawing.Point(110, 142)
        Me.variantenName.Name = "variantenName"
        Me.variantenName.Size = New System.Drawing.Size(211, 20)
        Me.variantenName.TabIndex = 4
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 116)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(43, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Projekt:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 146)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(49, 13)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Variante:"
        '
        'OKButton
        '
        Me.OKButton.Location = New System.Drawing.Point(143, 185)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(75, 23)
        Me.OKButton.TabIndex = 7
        Me.OKButton.Text = "OK"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'frmCreateNewVariant
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(350, 229)
        Me.Controls.Add(Me.OKButton)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.variantenName)
        Me.Controls.Add(Me.projektName)
        Me.Controls.Add(Me.newVariant)
        Me.Controls.Add(Me.infoText)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmCreateNewVariant"
        Me.Text = "Neue Variante anlegen"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents infoText As System.Windows.Forms.Label
    Friend WithEvents projektName As System.Windows.Forms.TextBox
    Friend WithEvents variantenName As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents OKButton As System.Windows.Forms.Button
    Public WithEvents newVariant As System.Windows.Forms.TextBox
End Class
