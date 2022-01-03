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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmshowBedarfe))
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
        Me.lblTyp.Location = New System.Drawing.Point(39, 30)
        Me.lblTyp.Name = "lblTyp"
        Me.lblTyp.Size = New System.Drawing.Size(25, 13)
        Me.lblTyp.TabIndex = 0
        Me.lblTyp.Text = "Typ"
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(134, 27)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(224, 21)
        Me.ComboBox1.TabIndex = 1
        '
        'lblName
        '
        Me.lblName.AutoSize = True
        Me.lblName.Location = New System.Drawing.Point(42, 72)
        Me.lblName.Name = "lblName"
        Me.lblName.Size = New System.Drawing.Size(35, 13)
        Me.lblName.TabIndex = 2
        Me.lblName.Text = "Name"
        '
        'ComboBox2
        '
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Location = New System.Drawing.Point(134, 69)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(224, 21)
        Me.ComboBox2.TabIndex = 3
        '
        'AnzeigenButton
        '
        Me.AnzeigenButton.Location = New System.Drawing.Point(92, 127)
        Me.AnzeigenButton.Name = "AnzeigenButton"
        Me.AnzeigenButton.Size = New System.Drawing.Size(80, 22)
        Me.AnzeigenButton.TabIndex = 4
        Me.AnzeigenButton.Text = "Anzeigen"
        Me.AnzeigenButton.UseVisualStyleBackColor = True
        '
        'AbbrButton
        '
        Me.AbbrButton.Location = New System.Drawing.Point(230, 127)
        Me.AbbrButton.Name = "AbbrButton"
        Me.AbbrButton.Size = New System.Drawing.Size(80, 22)
        Me.AbbrButton.TabIndex = 5
        Me.AbbrButton.Text = "Schließen"
        Me.AbbrButton.UseVisualStyleBackColor = True
        '
        'frmshowBedarfe
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.ClientSize = New System.Drawing.Size(404, 170)
        Me.Controls.Add(Me.AbbrButton)
        Me.Controls.Add(Me.AnzeigenButton)
        Me.Controls.Add(Me.ComboBox2)
        Me.Controls.Add(Me.lblName)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.lblTyp)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmshowBedarfe"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
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
