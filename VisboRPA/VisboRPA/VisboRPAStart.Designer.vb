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
        Me.watchFolder = New System.IO.FileSystemWatcher()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.activePortfolioSel = New System.Windows.Forms.ComboBox()
        Me.activePortfolioLbl = New System.Windows.Forms.Label()
        Me.VCselectionLbl = New System.Windows.Forms.Label()
        Me.VCSelection = New System.Windows.Forms.ComboBox()
        Me.statusMessage = New System.Windows.Forms.Label()
        Me.rpaDir = New System.Windows.Forms.TextBox()
        Me.durchsuchen = New System.Windows.Forms.Button()
        Me.ueberschrift = New System.Windows.Forms.Label()
        Me.btn_stop = New System.Windows.Forms.Button()
        Me.btn_start = New System.Windows.Forms.Button()
        CType(Me.watchFolder, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'watchFolder
        '
        Me.watchFolder.EnableRaisingEvents = True
        Me.watchFolder.NotifyFilter = System.IO.NotifyFilters.FileName
        Me.watchFolder.Path = "C:\VISBO\VISBO Config Data"
        Me.watchFolder.SynchronizingObject = Me
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BackColor = System.Drawing.SystemColors.Control
        Me.Panel1.Controls.Add(Me.activePortfolioSel)
        Me.Panel1.Controls.Add(Me.activePortfolioLbl)
        Me.Panel1.Controls.Add(Me.VCselectionLbl)
        Me.Panel1.Controls.Add(Me.VCSelection)
        Me.Panel1.Controls.Add(Me.statusMessage)
        Me.Panel1.Controls.Add(Me.rpaDir)
        Me.Panel1.Controls.Add(Me.durchsuchen)
        Me.Panel1.Controls.Add(Me.ueberschrift)
        Me.Panel1.Controls.Add(Me.btn_stop)
        Me.Panel1.Controls.Add(Me.btn_start)
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(669, 309)
        Me.Panel1.TabIndex = 0
        '
        'activePortfolioSel
        '
        Me.activePortfolioSel.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.activePortfolioSel.FormattingEnabled = True
        Me.activePortfolioSel.Location = New System.Drawing.Point(31, 207)
        Me.activePortfolioSel.Name = "activePortfolioSel"
        Me.activePortfolioSel.Size = New System.Drawing.Size(468, 21)
        Me.activePortfolioSel.TabIndex = 14
        '
        'activePortfolioLbl
        '
        Me.activePortfolioLbl.AutoSize = True
        Me.activePortfolioLbl.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.activePortfolioLbl.Location = New System.Drawing.Point(28, 176)
        Me.activePortfolioLbl.Name = "activePortfolioLbl"
        Me.activePortfolioLbl.Size = New System.Drawing.Size(97, 16)
        Me.activePortfolioLbl.TabIndex = 13
        Me.activePortfolioLbl.Text = "Active Portfolio"
        '
        'VCselectionLbl
        '
        Me.VCselectionLbl.AutoSize = True
        Me.VCselectionLbl.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.VCselectionLbl.Location = New System.Drawing.Point(28, 100)
        Me.VCselectionLbl.Name = "VCselectionLbl"
        Me.VCselectionLbl.Size = New System.Drawing.Size(89, 17)
        Me.VCselectionLbl.TabIndex = 12
        Me.VCselectionLbl.Text = "Visbo Center"
        '
        'VCSelection
        '
        Me.VCSelection.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.VCSelection.FormattingEnabled = True
        Me.VCSelection.Location = New System.Drawing.Point(31, 129)
        Me.VCSelection.Name = "VCSelection"
        Me.VCSelection.Size = New System.Drawing.Size(468, 21)
        Me.VCSelection.TabIndex = 11
        '
        'statusMessage
        '
        Me.statusMessage.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.statusMessage.AutoSize = True
        Me.statusMessage.Location = New System.Drawing.Point(28, 292)
        Me.statusMessage.Name = "statusMessage"
        Me.statusMessage.Size = New System.Drawing.Size(0, 13)
        Me.statusMessage.TabIndex = 10
        '
        'rpaDir
        '
        Me.rpaDir.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.rpaDir.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.rpaDir.Location = New System.Drawing.Point(31, 62)
        Me.rpaDir.Name = "rpaDir"
        Me.rpaDir.ReadOnly = True
        Me.rpaDir.Size = New System.Drawing.Size(468, 20)
        Me.rpaDir.TabIndex = 9
        '
        'durchsuchen
        '
        Me.durchsuchen.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.durchsuchen.Location = New System.Drawing.Point(505, 61)
        Me.durchsuchen.Name = "durchsuchen"
        Me.durchsuchen.Size = New System.Drawing.Size(137, 23)
        Me.durchsuchen.TabIndex = 8
        Me.durchsuchen.Text = "Durchsuchen"
        Me.durchsuchen.UseVisualStyleBackColor = True
        '
        'ueberschrift
        '
        Me.ueberschrift.AutoSize = True
        Me.ueberschrift.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ueberschrift.Location = New System.Drawing.Point(27, 29)
        Me.ueberschrift.Name = "ueberschrift"
        Me.ueberschrift.Size = New System.Drawing.Size(315, 17)
        Me.ueberschrift.TabIndex = 7
        Me.ueberschrift.Text = "Folder der zu Importierenden Dateien auswählen"
        '
        'btn_stop
        '
        Me.btn_stop.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_stop.Location = New System.Drawing.Point(567, 259)
        Me.btn_stop.Name = "btn_stop"
        Me.btn_stop.Size = New System.Drawing.Size(75, 22)
        Me.btn_stop.TabIndex = 5
        Me.btn_stop.Text = "Stop"
        Me.btn_stop.UseVisualStyleBackColor = True
        '
        'btn_start
        '
        Me.btn_start.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btn_start.Location = New System.Drawing.Point(31, 259)
        Me.btn_start.Name = "btn_start"
        Me.btn_start.Size = New System.Drawing.Size(75, 22)
        Me.btn_start.TabIndex = 4
        Me.btn_start.Text = "Start"
        Me.btn_start.UseVisualStyleBackColor = True
        '
        'VisboRPAStart
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.ClientSize = New System.Drawing.Size(662, 311)
        Me.Controls.Add(Me.Panel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "VisboRPAStart"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "VISBO RPA"
        CType(Me.watchFolder, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents watchFolder As IO.FileSystemWatcher
    Friend WithEvents FolderBrowserDialog1 As Windows.Forms.FolderBrowserDialog
    Friend WithEvents Panel1 As Windows.Forms.Panel
    Friend WithEvents ueberschrift As Windows.Forms.Label
    Friend WithEvents btn_stop As Windows.Forms.Button
    Friend WithEvents btn_start As Windows.Forms.Button
    Friend WithEvents durchsuchen As Windows.Forms.Button
    Friend WithEvents rpaDir As Windows.Forms.TextBox
    Friend WithEvents statusMessage As Windows.Forms.Label
    Friend WithEvents VCSelection As Windows.Forms.ComboBox
    Friend WithEvents activePortfolioSel As Windows.Forms.ComboBox
    Friend WithEvents activePortfolioLbl As Windows.Forms.Label
    Friend WithEvents VCselectionLbl As Windows.Forms.Label
End Class
