<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmconfirmOptimierung
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
        Me.ButtonJA = New System.Windows.Forms.Button()
        Me.ButtonNEIN = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'ButtonJA
        '
        Me.ButtonJA.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.ButtonJA.Location = New System.Drawing.Point(37, 66)
        Me.ButtonJA.Name = "ButtonJA"
        Me.ButtonJA.Size = New System.Drawing.Size(75, 23)
        Me.ButtonJA.TabIndex = 0
        Me.ButtonJA.Text = "Ja"
        Me.ButtonJA.UseVisualStyleBackColor = True
        '
        'ButtonNEIN
        '
        Me.ButtonNEIN.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.ButtonNEIN.Location = New System.Drawing.Point(222, 66)
        Me.ButtonNEIN.Name = "ButtonNEIN"
        Me.ButtonNEIN.Size = New System.Drawing.Size(75, 23)
        Me.ButtonNEIN.TabIndex = 1
        Me.ButtonNEIN.Text = "Nein"
        Me.ButtonNEIN.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(37, 30)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(263, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Möchten Sie das Optimierungs-Ergebnis übernehmen?"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(34, 108)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(0, 13)
        Me.Label2.TabIndex = 3
        '
        'frmconfirmOptimierung
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(347, 118)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ButtonNEIN)
        Me.Controls.Add(Me.ButtonJA)
        Me.Name = "frmconfirmOptimierung"
        Me.Text = "Optimierungs-Ergebnis"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents ButtonJA As System.Windows.Forms.Button
    Public WithEvents ButtonNEIN As System.Windows.Forms.Button
    Public WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents Label2 As System.Windows.Forms.Label
End Class
