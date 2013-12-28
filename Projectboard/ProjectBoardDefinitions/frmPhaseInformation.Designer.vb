<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPhaseInformation
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
        Me.phaseName = New System.Windows.Forms.TextBox()
        Me.phaseStart = New System.Windows.Forms.TextBox()
        Me.phaseEnde = New System.Windows.Forms.TextBox()
        Me.phaseDauer = New System.Windows.Forms.TextBox()
        Me.erlaeuterung = New System.Windows.Forms.TextBox()
        Me.projectName = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'phaseName
        '
        Me.phaseName.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.phaseName.Location = New System.Drawing.Point(18, 65)
        Me.phaseName.Name = "phaseName"
        Me.phaseName.Size = New System.Drawing.Size(421, 29)
        Me.phaseName.TabIndex = 1
        '
        'phaseStart
        '
        Me.phaseStart.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.phaseStart.Location = New System.Drawing.Point(18, 111)
        Me.phaseStart.Name = "phaseStart"
        Me.phaseStart.Size = New System.Drawing.Size(132, 22)
        Me.phaseStart.TabIndex = 2
        '
        'phaseEnde
        '
        Me.phaseEnde.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.phaseEnde.Location = New System.Drawing.Point(306, 111)
        Me.phaseEnde.Name = "phaseEnde"
        Me.phaseEnde.Size = New System.Drawing.Size(132, 22)
        Me.phaseEnde.TabIndex = 4
        '
        'phaseDauer
        '
        Me.phaseDauer.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.phaseDauer.Location = New System.Drawing.Point(164, 111)
        Me.phaseDauer.Name = "phaseDauer"
        Me.phaseDauer.Size = New System.Drawing.Size(128, 22)
        Me.phaseDauer.TabIndex = 3
        '
        'erlaeuterung
        '
        Me.erlaeuterung.BackColor = System.Drawing.SystemColors.Window
        Me.erlaeuterung.Enabled = False
        Me.erlaeuterung.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.erlaeuterung.Location = New System.Drawing.Point(18, 162)
        Me.erlaeuterung.MaximumSize = New System.Drawing.Size(420, 140)
        Me.erlaeuterung.MinimumSize = New System.Drawing.Size(420, 140)
        Me.erlaeuterung.Multiline = True
        Me.erlaeuterung.Name = "erlaeuterung"
        Me.erlaeuterung.ReadOnly = True
        Me.erlaeuterung.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.erlaeuterung.Size = New System.Drawing.Size(420, 140)
        Me.erlaeuterung.TabIndex = 0
        Me.erlaeuterung.Visible = False
        '
        'projectName
        '
        Me.projectName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.projectName.Location = New System.Drawing.Point(18, 26)
        Me.projectName.Name = "projectName"
        Me.projectName.Size = New System.Drawing.Size(420, 22)
        Me.projectName.TabIndex = 21
        '
        'frmPhaseInformation
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Window
        Me.ClientSize = New System.Drawing.Size(463, 330)
        Me.Controls.Add(Me.projectName)
        Me.Controls.Add(Me.erlaeuterung)
        Me.Controls.Add(Me.phaseDauer)
        Me.Controls.Add(Me.phaseEnde)
        Me.Controls.Add(Me.phaseStart)
        Me.Controls.Add(Me.phaseName)
        Me.Name = "frmPhaseInformation"
        Me.Text = "Phasen Information"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents erlaeuterung As System.Windows.Forms.TextBox
    Public WithEvents phaseName As System.Windows.Forms.TextBox
    Public WithEvents phaseStart As System.Windows.Forms.TextBox
    Public WithEvents phaseEnde As System.Windows.Forms.TextBox
    Public WithEvents phaseDauer As System.Windows.Forms.TextBox
    Public WithEvents projectName As System.Windows.Forms.TextBox
End Class
