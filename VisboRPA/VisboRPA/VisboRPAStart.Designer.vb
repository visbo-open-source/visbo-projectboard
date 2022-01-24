<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class VisboRPAStart
    Inherits System.Windows.Forms.Form

    'Das Formular überschreibt den Löschvorgang, um die Komponentenliste zu bereinigen.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(VisboRPAStart))
        Me.btn_start = New System.Windows.Forms.Button()
        Me.btn_stop = New System.Windows.Forms.Button()
        Me.watchFolder = New System.IO.FileSystemWatcher()
        CType(Me.watchFolder, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btn_start
        '
        Me.btn_start.Location = New System.Drawing.Point(12, 32)
        Me.btn_start.Name = "btn_start"
        Me.btn_start.Size = New System.Drawing.Size(75, 23)
        Me.btn_start.TabIndex = 0
        Me.btn_start.Text = "Start"
        Me.btn_start.UseVisualStyleBackColor = True
        '
        'btn_stop
        '
        Me.btn_stop.Location = New System.Drawing.Point(93, 32)
        Me.btn_stop.Name = "btn_stop"
        Me.btn_stop.Size = New System.Drawing.Size(75, 23)
        Me.btn_stop.TabIndex = 1
        Me.btn_stop.Text = "Stop"
        Me.btn_stop.UseVisualStyleBackColor = True
        '
        'watchFolder
        '
        Me.watchFolder.EnableRaisingEvents = True
        Me.watchFolder.NotifyFilter = System.IO.NotifyFilters.FileName
        Me.watchFolder.Path = "C:\VISBO\VISBO Config Data"
        Me.watchFolder.SynchronizingObject = Me
        '
        'VisboRPAStart
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(187, 62)
        Me.Controls.Add(Me.btn_stop)
        Me.Controls.Add(Me.btn_start)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "VisboRPAStart"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "VISBO RPA"
        Me.WindowState = System.Windows.Forms.FormWindowState.Minimized
        CType(Me.watchFolder, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents btn_start As Windows.Forms.Button
    Friend WithEvents btn_stop As Windows.Forms.Button
    Friend WithEvents watchFolder As IO.FileSystemWatcher
End Class
