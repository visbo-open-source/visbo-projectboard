<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSelectRPlanImport
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
        Me.RPLANImportDropbox = New System.Windows.Forms.ComboBox()
        Me.importRPLAN = New System.Windows.Forms.Button()
        Me.SelectAbbruch = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'RPLANImportDropbox
        '
        Me.RPLANImportDropbox.DropDownHeight = 200
        Me.RPLANImportDropbox.FormattingEnabled = True
        Me.RPLANImportDropbox.IntegralHeight = False
        Me.RPLANImportDropbox.Location = New System.Drawing.Point(25, 27)
        Me.RPLANImportDropbox.MaxDropDownItems = 10
        Me.RPLANImportDropbox.Name = "RPLANImportDropbox"
        Me.RPLANImportDropbox.Size = New System.Drawing.Size(378, 24)
        Me.RPLANImportDropbox.TabIndex = 0
        '
        'importRPLAN
        '
        Me.importRPLAN.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.importRPLAN.Location = New System.Drawing.Point(25, 90)
        Me.importRPLAN.Name = "importRPLAN"
        Me.importRPLAN.Size = New System.Drawing.Size(103, 25)
        Me.importRPLAN.TabIndex = 3
        Me.importRPLAN.Text = "OK" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.importRPLAN.UseVisualStyleBackColor = True
        '
        'SelectAbbruch
        '
        Me.SelectAbbruch.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.SelectAbbruch.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SelectAbbruch.Location = New System.Drawing.Point(274, 89)
        Me.SelectAbbruch.Name = "SelectAbbruch"
        Me.SelectAbbruch.Size = New System.Drawing.Size(128, 25)
        Me.SelectAbbruch.TabIndex = 4
        Me.SelectAbbruch.Text = "Abbrechen" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.SelectAbbruch.UseVisualStyleBackColor = True
        '
        'frmSelectRPlanImport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Window
        Me.ClientSize = New System.Drawing.Size(423, 135)
        Me.Controls.Add(Me.SelectAbbruch)
        Me.Controls.Add(Me.importRPLAN)
        Me.Controls.Add(Me.RPLANImportDropbox)
        Me.Name = "frmSelectRPlanImport"
        Me.Text = "RPLAN Dateien für Import auswählen"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents RPLANImportDropbox As System.Windows.Forms.ComboBox
    Friend WithEvents importRPLAN As System.Windows.Forms.Button
    Friend WithEvents SelectAbbruch As System.Windows.Forms.Button
End Class
