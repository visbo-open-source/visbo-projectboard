<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMeRcEinstellungen
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
        Me.chkbx_showHeader = New System.Windows.Forms.CheckBox()
        Me.chkbx_compareWithVersion = New System.Windows.Forms.CheckBox()
        Me.chkbx_allowOvertime = New System.Windows.Forms.CheckBox()
        Me.chkbx_noAutoDistribution = New System.Windows.Forms.CheckBox()
        Me.VersionDatePicker = New System.Windows.Forms.DateTimePicker()
        Me.ok_Btn = New System.Windows.Forms.Button()
        Me.cancel_btn = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'chkbx_showHeader
        '
        Me.chkbx_showHeader.AutoSize = True
        Me.chkbx_showHeader.Location = New System.Drawing.Point(22, 20)
        Me.chkbx_showHeader.Name = "chkbx_showHeader"
        Me.chkbx_showHeader.Size = New System.Drawing.Size(107, 17)
        Me.chkbx_showHeader.TabIndex = 0
        Me.chkbx_showHeader.Text = "Header anzeigen"
        Me.chkbx_showHeader.UseMnemonic = False
        Me.chkbx_showHeader.UseVisualStyleBackColor = True
        '
        'chkbx_compareWithVersion
        '
        Me.chkbx_compareWithVersion.AutoSize = True
        Me.chkbx_compareWithVersion.Location = New System.Drawing.Point(22, 56)
        Me.chkbx_compareWithVersion.Name = "chkbx_compareWithVersion"
        Me.chkbx_compareWithVersion.Size = New System.Drawing.Size(132, 17)
        Me.chkbx_compareWithVersion.TabIndex = 1
        Me.chkbx_compareWithVersion.Text = "vergleiche mit Version "
        Me.chkbx_compareWithVersion.UseVisualStyleBackColor = True
        '
        'chkbx_allowOvertime
        '
        Me.chkbx_allowOvertime.AutoSize = True
        Me.chkbx_allowOvertime.Location = New System.Drawing.Point(22, 97)
        Me.chkbx_allowOvertime.Name = "chkbx_allowOvertime"
        Me.chkbx_allowOvertime.Size = New System.Drawing.Size(135, 17)
        Me.chkbx_allowOvertime.TabIndex = 2
        Me.chkbx_allowOvertime.Text = "Überbuchung erlauben"
        Me.chkbx_allowOvertime.UseVisualStyleBackColor = True
        '
        'chkbx_noAutoDistribution
        '
        Me.chkbx_noAutoDistribution.AutoSize = True
        Me.chkbx_noAutoDistribution.Location = New System.Drawing.Point(22, 138)
        Me.chkbx_noAutoDistribution.Name = "chkbx_noAutoDistribution"
        Me.chkbx_noAutoDistribution.Size = New System.Drawing.Size(194, 17)
        Me.chkbx_noAutoDistribution.TabIndex = 3
        Me.chkbx_noAutoDistribution.Text = "keine autom Ressourcen-Verteilung"
        Me.chkbx_noAutoDistribution.UseVisualStyleBackColor = True
        '
        'VersionDatePicker
        '
        Me.VersionDatePicker.Location = New System.Drawing.Point(178, 55)
        Me.VersionDatePicker.Name = "VersionDatePicker"
        Me.VersionDatePicker.Size = New System.Drawing.Size(200, 20)
        Me.VersionDatePicker.TabIndex = 4
        '
        'ok_Btn
        '
        Me.ok_Btn.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.ok_Btn.Location = New System.Drawing.Point(80, 181)
        Me.ok_Btn.Name = "ok_Btn"
        Me.ok_Btn.Size = New System.Drawing.Size(75, 23)
        Me.ok_Btn.TabIndex = 5
        Me.ok_Btn.Text = "OK"
        Me.ok_Btn.UseVisualStyleBackColor = True
        '
        'cancel_btn
        '
        Me.cancel_btn.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cancel_btn.Location = New System.Drawing.Point(223, 181)
        Me.cancel_btn.Name = "cancel_btn"
        Me.cancel_btn.Size = New System.Drawing.Size(75, 23)
        Me.cancel_btn.TabIndex = 6
        Me.cancel_btn.Text = "Abbruch"
        Me.cancel_btn.UseVisualStyleBackColor = True
        '
        'frmMeRcEinstellungen
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(406, 223)
        Me.Controls.Add(Me.cancel_btn)
        Me.Controls.Add(Me.ok_Btn)
        Me.Controls.Add(Me.VersionDatePicker)
        Me.Controls.Add(Me.chkbx_noAutoDistribution)
        Me.Controls.Add(Me.chkbx_allowOvertime)
        Me.Controls.Add(Me.chkbx_compareWithVersion)
        Me.Controls.Add(Me.chkbx_showHeader)
        Me.Name = "frmMeRcEinstellungen"
        Me.Text = "Einstellungen"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents chkbx_compareWithVersion As Windows.Forms.CheckBox
    Friend WithEvents chkbx_allowOvertime As Windows.Forms.CheckBox
    Friend WithEvents chkbx_noAutoDistribution As Windows.Forms.CheckBox
    Friend WithEvents VersionDatePicker As Windows.Forms.DateTimePicker
    Private WithEvents chkbx_showHeader As Windows.Forms.CheckBox
    Friend WithEvents ok_Btn As Windows.Forms.Button
    Friend WithEvents cancel_btn As Windows.Forms.Button
End Class
