<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEarliestLatestStart
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
        Me.labellatestStart = New System.Windows.Forms.Label()
        Me.labelearliestStart = New System.Windows.Forms.Label()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.AbbruchButton = New System.Windows.Forms.Button()
        Me.EarliestStart = New System.Windows.Forms.TrackBar()
        Me.minEarliestStart = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.LatestStart = New System.Windows.Forms.TrackBar()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.maxLatestStart = New System.Windows.Forms.Label()
        Me.aktearliestStart = New System.Windows.Forms.Label()
        Me.aktlatestStart = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.EarliestStart, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.LatestStart, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'labellatestStart
        '
        Me.labellatestStart.AutoSize = True
        Me.labellatestStart.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labellatestStart.Location = New System.Drawing.Point(225, 49)
        Me.labellatestStart.Name = "labellatestStart"
        Me.labellatestStart.Size = New System.Drawing.Size(149, 18)
        Me.labellatestStart.TabIndex = 3
        Me.labellatestStart.Text = "spätester Start      "
        '
        'labelearliestStart
        '
        Me.labelearliestStart.AutoSize = True
        Me.labelearliestStart.BackColor = System.Drawing.SystemColors.Menu
        Me.labelearliestStart.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelearliestStart.ForeColor = System.Drawing.SystemColors.ControlText
        Me.labelearliestStart.Location = New System.Drawing.Point(34, 49)
        Me.labelearliestStart.Name = "labelearliestStart"
        Me.labelearliestStart.Size = New System.Drawing.Size(121, 18)
        Me.labelearliestStart.TabIndex = 4
        Me.labelearliestStart.Text = "frühester Start "
        '
        'OKButton
        '
        Me.OKButton.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.OKButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OKButton.Location = New System.Drawing.Point(37, 170)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(117, 39)
        Me.OKButton.TabIndex = 5
        Me.OKButton.Text = "OK"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'AbbruchButton
        '
        Me.AbbruchButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.AbbruchButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AbbruchButton.Location = New System.Drawing.Point(284, 171)
        Me.AbbruchButton.Name = "AbbruchButton"
        Me.AbbruchButton.Size = New System.Drawing.Size(123, 39)
        Me.AbbruchButton.TabIndex = 6
        Me.AbbruchButton.Text = "Abbrechen"
        Me.AbbruchButton.UseVisualStyleBackColor = True
        '
        'EarliestStart
        '
        Me.EarliestStart.LargeChange = 1
        Me.EarliestStart.Location = New System.Drawing.Point(26, 92)
        Me.EarliestStart.Maximum = 0
        Me.EarliestStart.Minimum = -10
        Me.EarliestStart.Name = "EarliestStart"
        Me.EarliestStart.Size = New System.Drawing.Size(193, 56)
        Me.EarliestStart.TabIndex = 7
        '
        'minEarliestStart
        '
        Me.minEarliestStart.AutoSize = True
        Me.minEarliestStart.Location = New System.Drawing.Point(23, 124)
        Me.minEarliestStart.Name = "minEarliestStart"
        Me.minEarliestStart.Size = New System.Drawing.Size(29, 17)
        Me.minEarliestStart.TabIndex = 8
        Me.minEarliestStart.Text = "-10"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(192, 124)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(16, 17)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "0"
        '
        'LatestStart
        '
        Me.LatestStart.LargeChange = 1
        Me.LatestStart.Location = New System.Drawing.Point(214, 92)
        Me.LatestStart.Name = "LatestStart"
        Me.LatestStart.Size = New System.Drawing.Size(193, 56)
        Me.LatestStart.TabIndex = 11
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(225, 124)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(16, 17)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "0"
        '
        'maxLatestStart
        '
        Me.maxLatestStart.AutoSize = True
        Me.maxLatestStart.Location = New System.Drawing.Point(383, 124)
        Me.maxLatestStart.Name = "maxLatestStart"
        Me.maxLatestStart.Size = New System.Drawing.Size(24, 17)
        Me.maxLatestStart.TabIndex = 13
        Me.maxLatestStart.Text = "10"
        '
        'aktearliestStart
        '
        Me.aktearliestStart.AutoSize = True
        Me.aktearliestStart.Location = New System.Drawing.Point(37, 69)
        Me.aktearliestStart.Name = "aktearliestStart"
        Me.aktearliestStart.Size = New System.Drawing.Size(16, 17)
        Me.aktearliestStart.TabIndex = 14
        Me.aktearliestStart.Text = "0"
        '
        'aktlatestStart
        '
        Me.aktlatestStart.AutoSize = True
        Me.aktlatestStart.Location = New System.Drawing.Point(225, 69)
        Me.aktlatestStart.Name = "aktlatestStart"
        Me.aktlatestStart.Size = New System.Drawing.Size(16, 17)
        Me.aktlatestStart.TabIndex = 15
        Me.aktlatestStart.Text = "0"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(180, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(91, 20)
        Me.Label1.TabIndex = 16
        Me.Label1.Text = "in Monaten"
        '
        'frmEarliestLatestStart
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(120.0!, 120.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.ClientSize = New System.Drawing.Size(435, 222)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.aktlatestStart)
        Me.Controls.Add(Me.aktearliestStart)
        Me.Controls.Add(Me.maxLatestStart)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.LatestStart)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.minEarliestStart)
        Me.Controls.Add(Me.EarliestStart)
        Me.Controls.Add(Me.AbbruchButton)
        Me.Controls.Add(Me.OKButton)
        Me.Controls.Add(Me.labelearliestStart)
        Me.Controls.Add(Me.labellatestStart)
        Me.Name = "frmEarliestLatestStart"
        Me.Text = "Zeitspanne für den Projektstart"
        CType(Me.EarliestStart, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.LatestStart, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents labellatestStart As System.Windows.Forms.Label
    Friend WithEvents labelearliestStart As System.Windows.Forms.Label
    Friend WithEvents OKButton As System.Windows.Forms.Button
    Friend WithEvents AbbruchButton As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents EarliestStart As System.Windows.Forms.TrackBar
    Public WithEvents LatestStart As System.Windows.Forms.TrackBar
    Public WithEvents minEarliestStart As System.Windows.Forms.Label
    Public WithEvents maxLatestStart As System.Windows.Forms.Label
    Friend WithEvents aktearliestStart As System.Windows.Forms.Label
    Friend WithEvents aktlatestStart As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
