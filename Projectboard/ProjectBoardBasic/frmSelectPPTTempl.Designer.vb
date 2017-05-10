<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSelectPPTTempl
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSelectPPTTempl))
        Me.createReport = New System.Windows.Forms.Button()
        Me.SelectAbbruch = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.RepVorlagenDropbox = New System.Windows.Forms.ComboBox()
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.statusNotification = New System.Windows.Forms.TextBox()
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker()
        Me.einstellungen = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'createReport
        '
        Me.createReport.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.createReport.Location = New System.Drawing.Point(23, 104)
        Me.createReport.Margin = New System.Windows.Forms.Padding(2)
        Me.createReport.Name = "createReport"
        Me.createReport.Size = New System.Drawing.Size(88, 25)
        Me.createReport.TabIndex = 1
        Me.createReport.Text = "OK"
        Me.createReport.UseVisualStyleBackColor = True
        '
        'SelectAbbruch
        '
        Me.SelectAbbruch.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SelectAbbruch.Location = New System.Drawing.Point(189, 104)
        Me.SelectAbbruch.Margin = New System.Windows.Forms.Padding(2)
        Me.SelectAbbruch.Name = "SelectAbbruch"
        Me.SelectAbbruch.Size = New System.Drawing.Size(121, 25)
        Me.SelectAbbruch.TabIndex = 2
        Me.SelectAbbruch.Text = "Abbrechen"
        Me.SelectAbbruch.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(20, 35)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(93, 15)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Report-Vorlage:"
        '
        'RepVorlagenDropbox
        '
        Me.RepVorlagenDropbox.DropDownHeight = 200
        Me.RepVorlagenDropbox.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RepVorlagenDropbox.FormattingEnabled = True
        Me.RepVorlagenDropbox.IntegralHeight = False
        Me.RepVorlagenDropbox.ItemHeight = 15
        Me.RepVorlagenDropbox.Location = New System.Drawing.Point(115, 35)
        Me.RepVorlagenDropbox.Margin = New System.Windows.Forms.Padding(2)
        Me.RepVorlagenDropbox.Name = "RepVorlagenDropbox"
        Me.RepVorlagenDropbox.Size = New System.Drawing.Size(192, 23)
        Me.RepVorlagenDropbox.TabIndex = 20
        '
        'BackgroundWorker1
        '
        Me.BackgroundWorker1.WorkerReportsProgress = True
        Me.BackgroundWorker1.WorkerSupportsCancellation = True
        '
        'statusNotification
        '
        Me.statusNotification.Location = New System.Drawing.Point(23, 147)
        Me.statusNotification.Margin = New System.Windows.Forms.Padding(2)
        Me.statusNotification.Name = "statusNotification"
        Me.statusNotification.Size = New System.Drawing.Size(287, 20)
        Me.statusNotification.TabIndex = 21
        Me.statusNotification.Text = "Status-Meldungen"
        '
        'BackgroundWorker2
        '
        Me.BackgroundWorker2.WorkerReportsProgress = True
        Me.BackgroundWorker2.WorkerSupportsCancellation = True
        '
        'einstellungen
        '
        Me.einstellungen.AutoSize = True
        Me.einstellungen.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, CType((System.Drawing.FontStyle.Italic Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.einstellungen.ForeColor = System.Drawing.SystemColors.MenuHighlight
        Me.einstellungen.Location = New System.Drawing.Point(237, 75)
        Me.einstellungen.Name = "einstellungen"
        Me.einstellungen.Size = New System.Drawing.Size(70, 13)
        Me.einstellungen.TabIndex = 23
        Me.einstellungen.Text = "Einstellungen"
        '
        'frmSelectPPTTempl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.BackColor = System.Drawing.SystemColors.Window
        Me.ClientSize = New System.Drawing.Size(345, 193)
        Me.Controls.Add(Me.einstellungen)
        Me.Controls.Add(Me.statusNotification)
        Me.Controls.Add(Me.RepVorlagenDropbox)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.SelectAbbruch)
        Me.Controls.Add(Me.createReport)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "frmSelectPPTTempl"
        Me.Text = "Auswählen Report-Vorlage"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents createReport As System.Windows.Forms.Button
    Friend WithEvents SelectAbbruch As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents RepVorlagenDropbox As System.Windows.Forms.ComboBox
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents statusNotification As System.Windows.Forms.TextBox
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents einstellungen As System.Windows.Forms.Label
End Class
