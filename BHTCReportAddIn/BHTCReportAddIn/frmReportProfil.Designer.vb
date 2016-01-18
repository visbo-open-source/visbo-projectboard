<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmReportProfil
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
        Me.RepProfilListbox = New System.Windows.Forms.ListBox()
        Me.zeitLabel = New System.Windows.Forms.Label()
        Me.vonDate = New System.Windows.Forms.DateTimePicker()
        Me.bisDate = New System.Windows.Forms.DateTimePicker()
        Me.ReportErstellen = New System.Windows.Forms.Button()
        Me.changeProfil = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'RepProfilListbox
        '
        Me.RepProfilListbox.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RepProfilListbox.FormattingEnabled = True
        Me.RepProfilListbox.HorizontalScrollbar = True
        Me.RepProfilListbox.ItemHeight = 16
        Me.RepProfilListbox.Location = New System.Drawing.Point(34, 31)
        Me.RepProfilListbox.Margin = New System.Windows.Forms.Padding(5)
        Me.RepProfilListbox.Name = "RepProfilListbox"
        Me.RepProfilListbox.ScrollAlwaysVisible = True
        Me.RepProfilListbox.Size = New System.Drawing.Size(579, 372)
        Me.RepProfilListbox.TabIndex = 0
        '
        'zeitLabel
        '
        Me.zeitLabel.AutoSize = True
        Me.zeitLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte), True)
        Me.zeitLabel.Location = New System.Drawing.Point(34, 454)
        Me.zeitLabel.Name = "zeitLabel"
        Me.zeitLabel.Size = New System.Drawing.Size(76, 20)
        Me.zeitLabel.TabIndex = 1
        Me.zeitLabel.Text = "Zeitraum:"
        Me.zeitLabel.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'vonDate
        '
        Me.vonDate.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.vonDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.vonDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.vonDate.Location = New System.Drawing.Point(218, 451)
        Me.vonDate.Name = "vonDate"
        Me.vonDate.Size = New System.Drawing.Size(137, 26)
        Me.vonDate.TabIndex = 4
        '
        'bisDate
        '
        Me.bisDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bisDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.bisDate.Location = New System.Drawing.Point(483, 448)
        Me.bisDate.Name = "bisDate"
        Me.bisDate.Size = New System.Drawing.Size(130, 26)
        Me.bisDate.TabIndex = 5
        '
        'ReportErstellen
        '
        Me.ReportErstellen.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ReportErstellen.Location = New System.Drawing.Point(38, 496)
        Me.ReportErstellen.Name = "ReportErstellen"
        Me.ReportErstellen.Size = New System.Drawing.Size(225, 27)
        Me.ReportErstellen.TabIndex = 6
        Me.ReportErstellen.Text = "Bericht erstellen"
        Me.ReportErstellen.UseVisualStyleBackColor = True
        '
        'changeProfil
        '
        Me.changeProfil.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.changeProfil.Location = New System.Drawing.Point(388, 496)
        Me.changeProfil.Name = "changeProfil"
        Me.changeProfil.Size = New System.Drawing.Size(225, 27)
        Me.changeProfil.TabIndex = 7
        Me.changeProfil.Text = "Profil bearbeiten"
        Me.changeProfil.UseVisualStyleBackColor = True
        '
        'frmReportProfil
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(662, 546)
        Me.Controls.Add(Me.changeProfil)
        Me.Controls.Add(Me.ReportErstellen)
        Me.Controls.Add(Me.bisDate)
        Me.Controls.Add(Me.vonDate)
        Me.Controls.Add(Me.zeitLabel)
        Me.Controls.Add(Me.RepProfilListbox)
        Me.Name = "frmReportProfil"
        Me.Text = "Auswahl Report Profil"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents RepProfilListbox As System.Windows.Forms.ListBox
    Friend WithEvents zeitLabel As System.Windows.Forms.Label
    Friend WithEvents vonDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents bisDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents ReportErstellen As System.Windows.Forms.Button
    Friend WithEvents changeProfil As System.Windows.Forms.Button
End Class
