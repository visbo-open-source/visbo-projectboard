<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmOptimizeKPI
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.progressText = New System.Windows.Forms.Label()
        Me.abbruchButton = New System.Windows.Forms.Button()
        Me.goldMedal = New System.Windows.Forms.Button()
        Me.silverMedal = New System.Windows.Forms.Button()
        Me.bronceMedal = New System.Windows.Forms.Button()
        Me.auswahlKPI = New System.Windows.Forms.ComboBox()
        Me.startButton = New System.Windows.Forms.Button()
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.SuspendLayout()
        '
        'progressText
        '
        Me.progressText.AutoSize = True
        Me.progressText.Location = New System.Drawing.Point(25, 184)
        Me.progressText.Name = "progressText"
        Me.progressText.Size = New System.Drawing.Size(68, 13)
        Me.progressText.TabIndex = 0
        Me.progressText.Text = ".. fortschritt .."
        '
        'abbruchButton
        '
        Me.abbruchButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.abbruchButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.abbruchButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.abbruchButton.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.abbruchButton.Location = New System.Drawing.Point(55, 63)
        Me.abbruchButton.Name = "abbruchButton"
        Me.abbruchButton.Size = New System.Drawing.Size(138, 100)
        Me.abbruchButton.TabIndex = 1
        Me.abbruchButton.Text = "Abbruch"
        Me.abbruchButton.UseVisualStyleBackColor = False
        '
        'goldMedal
        '
        Me.goldMedal.Location = New System.Drawing.Point(86, 66)
        Me.goldMedal.Name = "goldMedal"
        Me.goldMedal.Size = New System.Drawing.Size(75, 23)
        Me.goldMedal.TabIndex = 2
        Me.goldMedal.Text = "Platz 1"
        Me.goldMedal.UseVisualStyleBackColor = True
        '
        'silverMedal
        '
        Me.silverMedal.Location = New System.Drawing.Point(86, 102)
        Me.silverMedal.Name = "silverMedal"
        Me.silverMedal.Size = New System.Drawing.Size(75, 23)
        Me.silverMedal.TabIndex = 3
        Me.silverMedal.Text = "Platz 2"
        Me.silverMedal.UseVisualStyleBackColor = True
        '
        'bronceMedal
        '
        Me.bronceMedal.Location = New System.Drawing.Point(86, 140)
        Me.bronceMedal.Name = "bronceMedal"
        Me.bronceMedal.Size = New System.Drawing.Size(75, 23)
        Me.bronceMedal.TabIndex = 4
        Me.bronceMedal.Text = "Platz 3"
        Me.bronceMedal.UseVisualStyleBackColor = True
        '
        'auswahlKPI
        '
        Me.auswahlKPI.FormattingEnabled = True
        Me.auswahlKPI.Location = New System.Drawing.Point(55, 25)
        Me.auswahlKPI.Name = "auswahlKPI"
        Me.auswahlKPI.Size = New System.Drawing.Size(138, 21)
        Me.auswahlKPI.TabIndex = 5
        '
        'startButton
        '
        Me.startButton.Location = New System.Drawing.Point(86, 169)
        Me.startButton.Name = "startButton"
        Me.startButton.Size = New System.Drawing.Size(75, 23)
        Me.startButton.TabIndex = 6
        Me.startButton.Text = "Optimieren"
        Me.startButton.UseVisualStyleBackColor = True
        '
        'BackgroundWorker1
        '
        Me.BackgroundWorker1.WorkerReportsProgress = True
        Me.BackgroundWorker1.WorkerSupportsCancellation = True
        '
        'frmOptimizeKPI
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(247, 216)
        Me.Controls.Add(Me.startButton)
        Me.Controls.Add(Me.auswahlKPI)
        Me.Controls.Add(Me.bronceMedal)
        Me.Controls.Add(Me.silverMedal)
        Me.Controls.Add(Me.goldMedal)
        Me.Controls.Add(Me.abbruchButton)
        Me.Controls.Add(Me.progressText)
        Me.Name = "frmOptimizeKPI"
        Me.Text = "Kennzahlen optimieren"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents progressText As System.Windows.Forms.Label
    Friend WithEvents abbruchButton As System.Windows.Forms.Button
    Friend WithEvents silverMedal As System.Windows.Forms.Button
    Friend WithEvents bronceMedal As System.Windows.Forms.Button
    Public WithEvents goldMedal As System.Windows.Forms.Button
    Friend WithEvents auswahlKPI As System.Windows.Forms.ComboBox
    Friend WithEvents startButton As System.Windows.Forms.Button
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
End Class
