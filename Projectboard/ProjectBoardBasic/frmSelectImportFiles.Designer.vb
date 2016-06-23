<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSelectImportFiles
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
        Me.ListImportFiles = New System.Windows.Forms.ListBox()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.SelectAbbruch = New System.Windows.Forms.Button()
        Me.alleButton = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'ListImportFiles
        '
        Me.ListImportFiles.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ListImportFiles.FormattingEnabled = True
        Me.ListImportFiles.HorizontalScrollbar = True
        Me.ListImportFiles.ItemHeight = 15
        Me.ListImportFiles.Location = New System.Drawing.Point(12, 12)
        Me.ListImportFiles.Name = "ListImportFiles"
        Me.ListImportFiles.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple
        Me.ListImportFiles.Size = New System.Drawing.Size(473, 319)
        Me.ListImportFiles.Sorted = True
        Me.ListImportFiles.TabIndex = 0
        '
        'OKButton
        '
        Me.OKButton.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.OKButton.Location = New System.Drawing.Point(254, 353)
        Me.OKButton.Margin = New System.Windows.Forms.Padding(2)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(82, 22)
        Me.OKButton.TabIndex = 4
        Me.OKButton.Text = "OK" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'SelectAbbruch
        '
        Me.SelectAbbruch.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.SelectAbbruch.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SelectAbbruch.Location = New System.Drawing.Point(383, 352)
        Me.SelectAbbruch.Margin = New System.Windows.Forms.Padding(2)
        Me.SelectAbbruch.Name = "SelectAbbruch"
        Me.SelectAbbruch.Size = New System.Drawing.Size(102, 23)
        Me.SelectAbbruch.TabIndex = 5
        Me.SelectAbbruch.Text = "Abbrechen" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.SelectAbbruch.UseVisualStyleBackColor = True
        '
        'alleButton
        '
        Me.alleButton.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.alleButton.Location = New System.Drawing.Point(12, 353)
        Me.alleButton.Margin = New System.Windows.Forms.Padding(2)
        Me.alleButton.Name = "alleButton"
        Me.alleButton.Size = New System.Drawing.Size(82, 22)
        Me.alleButton.TabIndex = 6
        Me.alleButton.Text = "Alle"
        Me.alleButton.UseVisualStyleBackColor = True
        '
        'frmSelectImportFiles
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(497, 397)
        Me.Controls.Add(Me.alleButton)
        Me.Controls.Add(Me.SelectAbbruch)
        Me.Controls.Add(Me.OKButton)
        Me.Controls.Add(Me.ListImportFiles)
        Me.Name = "frmSelectImportFiles"
        Me.Text = "Auswahl der Import-Dateien"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ListImportFiles As System.Windows.Forms.ListBox
    Friend WithEvents OKButton As System.Windows.Forms.Button
    Friend WithEvents SelectAbbruch As System.Windows.Forms.Button
    Friend WithEvents alleButton As System.Windows.Forms.Button
End Class
