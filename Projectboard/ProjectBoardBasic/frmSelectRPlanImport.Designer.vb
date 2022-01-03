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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSelectRPlanImport))
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
        Me.RPLANImportDropbox.Location = New System.Drawing.Point(20, 22)
        Me.RPLANImportDropbox.Margin = New System.Windows.Forms.Padding(2)
        Me.RPLANImportDropbox.MaxDropDownItems = 10
        Me.RPLANImportDropbox.Name = "RPLANImportDropbox"
        Me.RPLANImportDropbox.Size = New System.Drawing.Size(303, 21)
        Me.RPLANImportDropbox.TabIndex = 0
        '
        'importRPLAN
        '
        Me.importRPLAN.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.importRPLAN.Location = New System.Drawing.Point(20, 72)
        Me.importRPLAN.Margin = New System.Windows.Forms.Padding(2)
        Me.importRPLAN.Name = "importRPLAN"
        Me.importRPLAN.Size = New System.Drawing.Size(82, 20)
        Me.importRPLAN.TabIndex = 3
        Me.importRPLAN.Text = "OK" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.importRPLAN.UseVisualStyleBackColor = True
        '
        'SelectAbbruch
        '
        Me.SelectAbbruch.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.SelectAbbruch.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SelectAbbruch.Location = New System.Drawing.Point(219, 71)
        Me.SelectAbbruch.Margin = New System.Windows.Forms.Padding(2)
        Me.SelectAbbruch.Name = "SelectAbbruch"
        Me.SelectAbbruch.Size = New System.Drawing.Size(102, 20)
        Me.SelectAbbruch.TabIndex = 4
        Me.SelectAbbruch.Text = "Abbrechen" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.SelectAbbruch.UseVisualStyleBackColor = True
        '
        'frmSelectRPlanImport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.BackColor = System.Drawing.SystemColors.Window
        Me.ClientSize = New System.Drawing.Size(338, 108)
        Me.Controls.Add(Me.SelectAbbruch)
        Me.Controls.Add(Me.importRPLAN)
        Me.Controls.Add(Me.RPLANImportDropbox)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "frmSelectRPlanImport"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "RPLAN Dateien für Import auswählen"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents RPLANImportDropbox As System.Windows.Forms.ComboBox
    Friend WithEvents importRPLAN As System.Windows.Forms.Button
    Friend WithEvents SelectAbbruch As System.Windows.Forms.Button
End Class
