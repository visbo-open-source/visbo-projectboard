<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRemoveConstellation
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
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.Abbrechen = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'ListBox1
        '
        Me.ListBox1.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.HorizontalScrollbar = True
        Me.ListBox1.ItemHeight = 19
        Me.ListBox1.Location = New System.Drawing.Point(34, 23)
        Me.ListBox1.Margin = New System.Windows.Forms.Padding(4)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(424, 232)
        Me.ListBox1.Sorted = True
        Me.ListBox1.TabIndex = 1
        '
        'OKButton
        '
        Me.OKButton.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.OKButton.Location = New System.Drawing.Point(88, 277)
        Me.OKButton.Margin = New System.Windows.Forms.Padding(4)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(100, 28)
        Me.OKButton.TabIndex = 2
        Me.OKButton.Text = "OK"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'Abbrechen
        '
        Me.Abbrechen.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Abbrechen.Location = New System.Drawing.Point(273, 277)
        Me.Abbrechen.Margin = New System.Windows.Forms.Padding(4)
        Me.Abbrechen.Name = "Abbrechen"
        Me.Abbrechen.Size = New System.Drawing.Size(100, 28)
        Me.Abbrechen.TabIndex = 3
        Me.Abbrechen.Text = "Abbrechen"
        Me.Abbrechen.UseVisualStyleBackColor = True
        '
        'frmRemoveConstellation
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(492, 328)
        Me.Controls.Add(Me.Abbrechen)
        Me.Controls.Add(Me.OKButton)
        Me.Controls.Add(Me.ListBox1)
        Me.Name = "frmRemoveConstellation"
        Me.Text = "Portfolio löschen"
        Me.ResumeLayout(False)

    End Sub
    Public WithEvents ListBox1 As System.Windows.Forms.ListBox
    Public WithEvents OKButton As System.Windows.Forms.Button
    Public WithEvents Abbrechen As System.Windows.Forms.Button
End Class
