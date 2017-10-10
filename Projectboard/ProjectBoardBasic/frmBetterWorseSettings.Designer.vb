<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBetterWorseSettings
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBetterWorseSettings))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.CBnextMS = New System.Windows.Forms.CheckBox()
        Me.CBendOfP = New System.Windows.Forms.CheckBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.RBvglAbs = New System.Windows.Forms.RadioButton()
        Me.RBvglRel = New System.Windows.Forms.RadioButton()
        Me.RBvglL = New System.Windows.Forms.RadioButton()
        Me.RBvglB = New System.Windows.Forms.RadioButton()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.CostToleranz = New System.Windows.Forms.Label()
        Me.timeToleranz = New System.Windows.Forms.Label()
        Me.costTolerance = New System.Windows.Forms.TextBox()
        Me.timeTolerance = New System.Windows.Forms.TextBox()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.Panel5 = New System.Windows.Forms.Panel()
        Me.OkButton = New System.Windows.Forms.Button()
        Me.AbbrButton = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.Panel5.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.CBnextMS)
        Me.Panel1.Controls.Add(Me.CBendOfP)
        Me.Panel1.Location = New System.Drawing.Point(44, 21)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(182, 88)
        Me.Panel1.TabIndex = 0
        '
        'CBnextMS
        '
        Me.CBnextMS.AutoSize = True
        Me.CBnextMS.Checked = True
        Me.CBnextMS.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CBnextMS.Location = New System.Drawing.Point(20, 48)
        Me.CBnextMS.Name = "CBnextMS"
        Me.CBnextMS.Size = New System.Drawing.Size(123, 17)
        Me.CBnextMS.TabIndex = 1
        Me.CBnextMS.Text = "nächster Meilenstein"
        Me.CBnextMS.UseVisualStyleBackColor = True
        '
        'CBendOfP
        '
        Me.CBendOfP.AutoSize = True
        Me.CBendOfP.Location = New System.Drawing.Point(20, 14)
        Me.CBendOfP.Name = "CBendOfP"
        Me.CBendOfP.Size = New System.Drawing.Size(87, 17)
        Me.CBendOfP.TabIndex = 0
        Me.CBendOfP.Text = "Projekt-Ende"
        Me.CBendOfP.UseVisualStyleBackColor = True
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.RBvglAbs)
        Me.Panel2.Controls.Add(Me.RBvglRel)
        Me.Panel2.Location = New System.Drawing.Point(44, 114)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(182, 88)
        Me.Panel2.TabIndex = 1
        '
        'RBvglAbs
        '
        Me.RBvglAbs.AutoSize = True
        Me.RBvglAbs.Checked = True
        Me.RBvglAbs.Location = New System.Drawing.Point(20, 50)
        Me.RBvglAbs.Name = "RBvglAbs"
        Me.RBvglAbs.Size = New System.Drawing.Size(127, 17)
        Me.RBvglAbs.TabIndex = 1
        Me.RBvglAbs.TabStop = True
        Me.RBvglAbs.Text = "absolute Abweichung"
        Me.RBvglAbs.UseVisualStyleBackColor = True
        '
        'RBvglRel
        '
        Me.RBvglRel.AutoSize = True
        Me.RBvglRel.Location = New System.Drawing.Point(20, 16)
        Me.RBvglRel.Name = "RBvglRel"
        Me.RBvglRel.Size = New System.Drawing.Size(121, 17)
        Me.RBvglRel.TabIndex = 0
        Me.RBvglRel.Text = "relative Abweichung"
        Me.RBvglRel.UseVisualStyleBackColor = True
        '
        'RBvglL
        '
        Me.RBvglL.AutoSize = True
        Me.RBvglL.Location = New System.Drawing.Point(16, 48)
        Me.RBvglL.Name = "RBvglL"
        Me.RBvglL.Size = New System.Drawing.Size(152, 17)
        Me.RBvglL.TabIndex = 1
        Me.RBvglL.Text = "Vergleich mit letztem Stand"
        Me.RBvglL.UseVisualStyleBackColor = True
        '
        'RBvglB
        '
        Me.RBvglB.AutoSize = True
        Me.RBvglB.Checked = True
        Me.RBvglB.Location = New System.Drawing.Point(16, 14)
        Me.RBvglB.Name = "RBvglB"
        Me.RBvglB.Size = New System.Drawing.Size(152, 17)
        Me.RBvglB.TabIndex = 0
        Me.RBvglB.TabStop = True
        Me.RBvglB.Text = "Vergleich mit Beauftragung"
        Me.RBvglB.UseVisualStyleBackColor = True
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.RBvglL)
        Me.Panel3.Controls.Add(Me.RBvglB)
        Me.Panel3.Location = New System.Drawing.Point(233, 21)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(191, 88)
        Me.Panel3.TabIndex = 2
        '
        'Panel4
        '
        Me.Panel4.Controls.Add(Me.CostToleranz)
        Me.Panel4.Controls.Add(Me.timeToleranz)
        Me.Panel4.Controls.Add(Me.costTolerance)
        Me.Panel4.Controls.Add(Me.timeTolerance)
        Me.Panel4.Location = New System.Drawing.Point(233, 114)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(191, 88)
        Me.Panel4.TabIndex = 2
        '
        'CostToleranz
        '
        Me.CostToleranz.AutoSize = True
        Me.CostToleranz.Location = New System.Drawing.Point(13, 50)
        Me.CostToleranz.Name = "CostToleranz"
        Me.CostToleranz.Size = New System.Drawing.Size(84, 13)
        Me.CostToleranz.TabIndex = 3
        Me.CostToleranz.Text = "Kosten-Toleranz"
        '
        'timeToleranz
        '
        Me.timeToleranz.AutoSize = True
        Me.timeToleranz.Location = New System.Drawing.Point(13, 18)
        Me.timeToleranz.Name = "timeToleranz"
        Me.timeToleranz.Size = New System.Drawing.Size(69, 13)
        Me.timeToleranz.TabIndex = 2
        Me.timeToleranz.Text = "Zeit-Toleranz"
        '
        'costTolerance
        '
        Me.costTolerance.Location = New System.Drawing.Point(128, 46)
        Me.costTolerance.Name = "costTolerance"
        Me.costTolerance.Size = New System.Drawing.Size(50, 20)
        Me.costTolerance.TabIndex = 1
        '
        'timeTolerance
        '
        Me.timeTolerance.Location = New System.Drawing.Point(128, 16)
        Me.timeTolerance.Name = "timeTolerance"
        Me.timeTolerance.Size = New System.Drawing.Size(50, 20)
        Me.timeTolerance.TabIndex = 0
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Checked = True
        Me.RadioButton1.Location = New System.Drawing.Point(20, 50)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(86, 17)
        Me.RadioButton1.TabIndex = 1
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "Ausstrahlung"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Location = New System.Drawing.Point(20, 16)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(137, 17)
        Me.RadioButton2.TabIndex = 0
        Me.RadioButton2.Text = "strategische Bedeutung"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'Panel5
        '
        Me.Panel5.Controls.Add(Me.RadioButton1)
        Me.Panel5.Controls.Add(Me.RadioButton2)
        Me.Panel5.Location = New System.Drawing.Point(44, 208)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(182, 88)
        Me.Panel5.TabIndex = 2
        '
        'OkButton
        '
        Me.OkButton.Location = New System.Drawing.Point(249, 223)
        Me.OkButton.Name = "OkButton"
        Me.OkButton.Size = New System.Drawing.Size(162, 22)
        Me.OkButton.TabIndex = 3
        Me.OkButton.Text = "OK"
        Me.OkButton.UseVisualStyleBackColor = True
        '
        'AbbrButton
        '
        Me.AbbrButton.Location = New System.Drawing.Point(249, 257)
        Me.AbbrButton.Name = "AbbrButton"
        Me.AbbrButton.Size = New System.Drawing.Size(162, 22)
        Me.AbbrButton.TabIndex = 4
        Me.AbbrButton.Text = "Abbrechen"
        Me.AbbrButton.UseVisualStyleBackColor = True
        '
        'frmBetterWorseSettings
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.ClientSize = New System.Drawing.Size(462, 312)
        Me.Controls.Add(Me.AbbrButton)
        Me.Controls.Add(Me.OkButton)
        Me.Controls.Add(Me.Panel5)
        Me.Controls.Add(Me.Panel4)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmBetterWorseSettings"
        Me.Text = "Einstellungen für Besser / Schlechter Diagramme"
        Me.TopMost = True
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.Panel4.ResumeLayout(False)
        Me.Panel4.PerformLayout()
        Me.Panel5.ResumeLayout(False)
        Me.Panel5.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents CBnextMS As System.Windows.Forms.CheckBox
    Friend WithEvents CBendOfP As System.Windows.Forms.CheckBox
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents RBvglAbs As System.Windows.Forms.RadioButton
    Friend WithEvents RBvglRel As System.Windows.Forms.RadioButton
    Friend WithEvents RBvglL As System.Windows.Forms.RadioButton
    Friend WithEvents RBvglB As System.Windows.Forms.RadioButton
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents timeToleranz As System.Windows.Forms.Label
    Friend WithEvents costTolerance As System.Windows.Forms.TextBox
    Friend WithEvents timeTolerance As System.Windows.Forms.TextBox
    Friend WithEvents CostToleranz As System.Windows.Forms.Label
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents OkButton As System.Windows.Forms.Button
    Friend WithEvents AbbrButton As System.Windows.Forms.Button
End Class
