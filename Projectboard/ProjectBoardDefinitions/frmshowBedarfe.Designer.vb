<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmshowBedarfe
    Inherits System.Windows.Forms.Form

    'Das Formular überschreibt den Löschvorgang, um die Komponentenliste zu bereinigen.
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

    'Wird vom Windows Form-Designer benötigt.
    Private components As System.ComponentModel.IContainer

    'Hinweis: Die folgende Prozedur ist für den Windows Form-Designer erforderlich.
    'Das Bearbeiten ist mit dem Windows Form-Designer möglich.  
    'Das Bearbeiten mit dem Code-Editor ist nicht möglich.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.lblTyp = New System.Windows.Forms.Label()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.lblName = New System.Windows.Forms.Label()
        Me.ComboBox2 = New System.Windows.Forms.ComboBox()
        Me.AnzeigenButton = New System.Windows.Forms.Button()
        Me.AbbrButton = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lblTyp
        '
        Me.lblTyp.AutoSize = True
        Me.lblTyp.Location = New System.Drawing.Point(49, 38)
        Me.lblTyp.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblTyp.Name = "lblTyp"
        Me.lblTyp.Size = New System.Drawing.Size(32, 17)
        Me.lblTyp.TabIndex = 0
        Me.lblTyp.Text = "Typ"
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(167, 34)
        Me.ComboBox1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(279, 24)
        Me.ComboBox1.TabIndex = 1
        '
        'lblName
        '
        Me.lblName.AutoSize = True
        Me.lblName.Location = New System.Drawing.Point(53, 90)
        Me.lblName.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblName.Name = "lblName"
        Me.lblName.Size = New System.Drawing.Size(45, 17)
        Me.lblName.TabIndex = 2
        Me.lblName.Text = "Name"
        '
        'ComboBox2
        '
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Location = New System.Drawing.Point(167, 86)
        Me.ComboBox2.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(279, 24)
        Me.ComboBox2.TabIndex = 3
        '
        'AnzeigenButton
        '
        Me.AnzeigenButton.Location = New System.Drawing.Point(115, 159)
        Me.AnzeigenButton.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.AnzeigenButton.Name = "AnzeigenButton"
        Me.AnzeigenButton.Size = New System.Drawing.Size(100, 28)
        Me.AnzeigenButton.TabIndex = 4
        Me.AnzeigenButton.Text = "Anzeigen"
        Me.AnzeigenButton.UseVisualStyleBackColor = True
        '
        'AbbrButton
        '
        Me.AbbrButton.Location = New System.Drawing.Point(287, 159)
        Me.AbbrButton.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.AbbrButton.Name = "AbbrButton"
        Me.AbbrButton.Size = New System.Drawing.Size(100, 28)
        Me.AbbrButton.TabIndex = 5
        Me.AbbrButton.Text = "Schließen"
        Me.AbbrButton.UseVisualStyleBackColor = True
        '
        'frmshowBedarfe
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(120.0!, 120.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.ClientSize = New System.Drawing.Size(505, 212)
        Me.Controls.Add(Me.AbbrButton)
        Me.Controls.Add(Me.AnzeigenButton)
        Me.Controls.Add(Me.ComboBox2)
        Me.Controls.Add(Me.lblName)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.lblTyp)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Name = "frmshowBedarfe"
        Me.Text = "detaillierte Bedarfe"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents lblTyp As System.Windows.Forms.Label
    Public WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Public WithEvents lblName As System.Windows.Forms.Label
    Public WithEvents ComboBox2 As System.Windows.Forms.ComboBox
    Public WithEvents AnzeigenButton As System.Windows.Forms.Button
    Public WithEvents AbbrButton As System.Windows.Forms.Button
End Class
