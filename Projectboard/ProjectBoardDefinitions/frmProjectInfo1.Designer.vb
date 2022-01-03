<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmProjectInfo1
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmProjectInfo1))
        Me.lblProjectName = New System.Windows.Forms.Label()
        Me.lblDBVersion = New System.Windows.Forms.Label()
        Me.dbForecast = New System.Windows.Forms.TextBox()
        Me.lblCurrentVersion = New System.Windows.Forms.Label()
        Me.currentForecast = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'lblProjectName
        '
        Me.lblProjectName.AutoSize = True
        Me.lblProjectName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProjectName.Location = New System.Drawing.Point(8, 9)
        Me.lblProjectName.Name = "lblProjectName"
        Me.lblProjectName.Size = New System.Drawing.Size(71, 13)
        Me.lblProjectName.TabIndex = 4
        Me.lblProjectName.Text = "Project-Name"
        '
        'lblDBVersion
        '
        Me.lblDBVersion.AutoSize = True
        Me.lblDBVersion.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDBVersion.Location = New System.Drawing.Point(165, 33)
        Me.lblDBVersion.Name = "lblDBVersion"
        Me.lblDBVersion.Size = New System.Drawing.Size(73, 13)
        Me.lblDBVersion.TabIndex = 6
        Me.lblDBVersion.Text = "DB (18.10.16)"
        '
        'dbForecast
        '
        Me.dbForecast.Enabled = False
        Me.dbForecast.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dbForecast.Location = New System.Drawing.Point(165, 57)
        Me.dbForecast.Name = "dbForecast"
        Me.dbForecast.Size = New System.Drawing.Size(86, 22)
        Me.dbForecast.TabIndex = 8
        Me.dbForecast.Text = "-30000,00 T€"
        '
        'lblCurrentVersion
        '
        Me.lblCurrentVersion.AutoSize = True
        Me.lblCurrentVersion.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCurrentVersion.Location = New System.Drawing.Point(8, 33)
        Me.lblCurrentVersion.Name = "lblCurrentVersion"
        Me.lblCurrentVersion.Size = New System.Drawing.Size(82, 13)
        Me.lblCurrentVersion.TabIndex = 9
        Me.lblCurrentVersion.Text = "aktuelle Version"
        '
        'currentForecast
        '
        Me.currentForecast.Enabled = False
        Me.currentForecast.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.currentForecast.Location = New System.Drawing.Point(8, 57)
        Me.currentForecast.Name = "currentForecast"
        Me.currentForecast.Size = New System.Drawing.Size(86, 22)
        Me.currentForecast.TabIndex = 10
        Me.currentForecast.Text = "-30000,00 T€"
        '
        'frmProjectInfo1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(268, 89)
        Me.Controls.Add(Me.currentForecast)
        Me.Controls.Add(Me.lblCurrentVersion)
        Me.Controls.Add(Me.dbForecast)
        Me.Controls.Add(Me.lblDBVersion)
        Me.Controls.Add(Me.lblProjectName)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmProjectInfo1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Projekt-Info"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblCurrentVersion As System.Windows.Forms.Label
    Public WithEvents lblProjectName As System.Windows.Forms.Label
    Public WithEvents lblDBVersion As System.Windows.Forms.Label
    Public WithEvents dbForecast As System.Windows.Forms.TextBox
    Public WithEvents currentForecast As System.Windows.Forms.TextBox
End Class
