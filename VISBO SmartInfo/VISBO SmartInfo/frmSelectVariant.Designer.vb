<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSelectVariant
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
        Me.variantNamesListBox = New System.Windows.Forms.ListBox()
        Me.showButton = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'variantNamesListBox
        '
        Me.variantNamesListBox.FormattingEnabled = True
        Me.variantNamesListBox.Location = New System.Drawing.Point(12, 25)
        Me.variantNamesListBox.Name = "variantNamesListBox"
        Me.variantNamesListBox.Size = New System.Drawing.Size(324, 225)
        Me.variantNamesListBox.TabIndex = 0
        '
        'showButton
        '
        Me.showButton.Location = New System.Drawing.Point(120, 267)
        Me.showButton.Name = "showButton"
        Me.showButton.Size = New System.Drawing.Size(121, 23)
        Me.showButton.TabIndex = 1
        Me.showButton.Text = "Variante anzeigen"
        Me.showButton.UseVisualStyleBackColor = True
        '
        'frmSelectVariant
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(348, 302)
        Me.Controls.Add(Me.showButton)
        Me.Controls.Add(Me.variantNamesListBox)
        Me.Name = "frmSelectVariant"
        Me.Text = "Variante anzeigen"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents variantNamesListBox As System.Windows.Forms.ListBox
    Friend WithEvents showButton As System.Windows.Forms.Button
End Class
