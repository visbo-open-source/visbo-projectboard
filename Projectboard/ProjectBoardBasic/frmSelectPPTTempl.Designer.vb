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
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.BackgroundWorker2 = New System.ComponentModel.BackgroundWorker()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.einstellungen = New System.Windows.Forms.Label()
        Me.statusNotification = New System.Windows.Forms.TextBox()
        Me.RepVorlagenDropbox = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SelectAbbruch = New System.Windows.Forms.Button()
        Me.createReport = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'BackgroundWorker1
        '
        Me.BackgroundWorker1.WorkerReportsProgress = True
        Me.BackgroundWorker1.WorkerSupportsCancellation = True
        '
        'BackgroundWorker2
        '
        Me.BackgroundWorker2.WorkerReportsProgress = True
        Me.BackgroundWorker2.WorkerSupportsCancellation = True
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.Controls.Add(Me.einstellungen)
        Me.Panel1.Controls.Add(Me.statusNotification)
        Me.Panel1.Controls.Add(Me.RepVorlagenDropbox)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.SelectAbbruch)
        Me.Panel1.Controls.Add(Me.createReport)
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(431, 242)
        Me.Panel1.TabIndex = 0
        '
        'einstellungen
        '
        Me.einstellungen.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.einstellungen.AutoSize = True
        Me.einstellungen.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, CType((System.Drawing.FontStyle.Italic Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.einstellungen.ForeColor = System.Drawing.SystemColors.MenuHighlight
        Me.einstellungen.Location = New System.Drawing.Point(304, 90)
        Me.einstellungen.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.einstellungen.Name = "einstellungen"
        Me.einstellungen.Size = New System.Drawing.Size(93, 17)
        Me.einstellungen.TabIndex = 29
        Me.einstellungen.Text = "Einstellungen"
        '
        'statusNotification
        '
        Me.statusNotification.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.statusNotification.Location = New System.Drawing.Point(37, 180)
        Me.statusNotification.Margin = New System.Windows.Forms.Padding(2)
        Me.statusNotification.Name = "statusNotification"
        Me.statusNotification.Size = New System.Drawing.Size(358, 22)
        Me.statusNotification.TabIndex = 28
        Me.statusNotification.Text = "Status-Meldungen"
        '
        'RepVorlagenDropbox
        '
        Me.RepVorlagenDropbox.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RepVorlagenDropbox.DropDownHeight = 200
        Me.RepVorlagenDropbox.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RepVorlagenDropbox.FormattingEnabled = True
        Me.RepVorlagenDropbox.IntegralHeight = False
        Me.RepVorlagenDropbox.ItemHeight = 18
        Me.RepVorlagenDropbox.Location = New System.Drawing.Point(152, 40)
        Me.RepVorlagenDropbox.Margin = New System.Windows.Forms.Padding(2)
        Me.RepVorlagenDropbox.Name = "RepVorlagenDropbox"
        Me.RepVorlagenDropbox.Size = New System.Drawing.Size(239, 26)
        Me.RepVorlagenDropbox.TabIndex = 27
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(33, 40)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(112, 18)
        Me.Label1.TabIndex = 26
        Me.Label1.Text = "Report-Vorlage:"
        '
        'SelectAbbruch
        '
        Me.SelectAbbruch.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SelectAbbruch.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SelectAbbruch.Location = New System.Drawing.Point(244, 126)
        Me.SelectAbbruch.Margin = New System.Windows.Forms.Padding(2)
        Me.SelectAbbruch.Name = "SelectAbbruch"
        Me.SelectAbbruch.Size = New System.Drawing.Size(151, 31)
        Me.SelectAbbruch.TabIndex = 25
        Me.SelectAbbruch.Text = "Abbrechen"
        Me.SelectAbbruch.UseVisualStyleBackColor = True
        '
        'createReport
        '
        Me.createReport.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.createReport.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.createReport.Location = New System.Drawing.Point(37, 126)
        Me.createReport.Margin = New System.Windows.Forms.Padding(2)
        Me.createReport.Name = "createReport"
        Me.createReport.Size = New System.Drawing.Size(110, 31)
        Me.createReport.TabIndex = 24
        Me.createReport.Text = "OK"
        Me.createReport.UseVisualStyleBackColor = True
        '
        'frmSelectPPTTempl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(120.0!, 120.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.BackColor = System.Drawing.SystemColors.Window
        Me.ClientSize = New System.Drawing.Size(431, 241)
        Me.Controls.Add(Me.Panel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "frmSelectPPTTempl"
        Me.Text = "Auswählen Report-Vorlage"
        Me.TopMost = True
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents BackgroundWorker2 As System.ComponentModel.BackgroundWorker
    Friend WithEvents Panel1 As Windows.Forms.Panel
    Friend WithEvents einstellungen As Windows.Forms.Label
    Friend WithEvents statusNotification As Windows.Forms.TextBox
    Friend WithEvents RepVorlagenDropbox As Windows.Forms.ComboBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents SelectAbbruch As Windows.Forms.Button
    Friend WithEvents createReport As Windows.Forms.Button
End Class
