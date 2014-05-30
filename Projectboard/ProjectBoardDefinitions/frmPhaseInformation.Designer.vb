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
        Me.projectName = New System.Windows.Forms.TextBox()
        Me.lessonsLearnedControl = New System.Windows.Forms.TabControl()
        Me.praemissen = New System.Windows.Forms.TabPage()
        Me.erlaeuterung = New System.Windows.Forms.TextBox()
        Me.sonderAblaeufe = New System.Windows.Forms.TabPage()
        Me.explSonderabl = New System.Windows.Forms.TextBox()
        Me.enabler = New System.Windows.Forms.TabPage()
        Me.explEnabler = New System.Windows.Forms.TextBox()
        Me.zusatzRisiken = New System.Windows.Forms.TabPage()
        Me.explRisiken = New System.Windows.Forms.TextBox()
        Me.teilnehmer = New System.Windows.Forms.TabPage()
        Me.lessonsLearnedControl.SuspendLayout()
        Me.praemissen.SuspendLayout()
        Me.sonderAblaeufe.SuspendLayout()
        Me.enabler.SuspendLayout()
        Me.zusatzRisiken.SuspendLayout()
        Me.SuspendLayout()
        '
        'phaseName
        '
        Me.phaseName.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.phaseName.Location = New System.Drawing.Point(24, 80)
        Me.phaseName.Margin = New System.Windows.Forms.Padding(4)
        Me.phaseName.Name = "phaseName"
        Me.phaseName.Size = New System.Drawing.Size(560, 34)
        Me.phaseName.TabIndex = 1
        '
        'phaseStart
        '
        Me.phaseStart.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.phaseStart.Location = New System.Drawing.Point(24, 137)
        Me.phaseStart.Margin = New System.Windows.Forms.Padding(4)
        Me.phaseStart.Name = "phaseStart"
        Me.phaseStart.Size = New System.Drawing.Size(175, 26)
        Me.phaseStart.TabIndex = 2
        '
        'phaseEnde
        '
        Me.phaseEnde.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.phaseEnde.Location = New System.Drawing.Point(408, 137)
        Me.phaseEnde.Margin = New System.Windows.Forms.Padding(4)
        Me.phaseEnde.Name = "phaseEnde"
        Me.phaseEnde.Size = New System.Drawing.Size(175, 26)
        Me.phaseEnde.TabIndex = 4
        '
        'phaseDauer
        '
        Me.phaseDauer.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.phaseDauer.Location = New System.Drawing.Point(219, 137)
        Me.phaseDauer.Margin = New System.Windows.Forms.Padding(4)
        Me.phaseDauer.Name = "phaseDauer"
        Me.phaseDauer.Size = New System.Drawing.Size(169, 26)
        Me.phaseDauer.TabIndex = 3
        '
        'projectName
        '
        Me.projectName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.projectName.Location = New System.Drawing.Point(24, 32)
        Me.projectName.Margin = New System.Windows.Forms.Padding(4)
        Me.projectName.Name = "projectName"
        Me.projectName.Size = New System.Drawing.Size(559, 26)
        Me.projectName.TabIndex = 21
        '
        'lessonsLearnedControl
        '
        Me.lessonsLearnedControl.Controls.Add(Me.praemissen)
        Me.lessonsLearnedControl.Controls.Add(Me.sonderAblaeufe)
        Me.lessonsLearnedControl.Controls.Add(Me.enabler)
        Me.lessonsLearnedControl.Controls.Add(Me.zusatzRisiken)
        Me.lessonsLearnedControl.Controls.Add(Me.teilnehmer)
        Me.lessonsLearnedControl.Location = New System.Drawing.Point(24, 201)
        Me.lessonsLearnedControl.Margin = New System.Windows.Forms.Padding(4)
        Me.lessonsLearnedControl.Name = "lessonsLearnedControl"
        Me.lessonsLearnedControl.SelectedIndex = 0
        Me.lessonsLearnedControl.Size = New System.Drawing.Size(560, 271)
        Me.lessonsLearnedControl.TabIndex = 22
        '
        'praemissen
        '
        Me.praemissen.Controls.Add(Me.erlaeuterung)
        Me.praemissen.Location = New System.Drawing.Point(4, 25)
        Me.praemissen.Margin = New System.Windows.Forms.Padding(4)
        Me.praemissen.Name = "praemissen"
        Me.praemissen.Padding = New System.Windows.Forms.Padding(4)
        Me.praemissen.Size = New System.Drawing.Size(552, 242)
        Me.praemissen.TabIndex = 0
        Me.praemissen.Text = "Prämissen"
        Me.praemissen.ToolTipText = "Anzeige der 26 Prämissen "
        Me.praemissen.UseVisualStyleBackColor = True
        '
        'erlaeuterung
        '
        Me.erlaeuterung.Location = New System.Drawing.Point(8, 7)
        Me.erlaeuterung.Margin = New System.Windows.Forms.Padding(4)
        Me.erlaeuterung.Multiline = True
        Me.erlaeuterung.Name = "erlaeuterung"
        Me.erlaeuterung.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.erlaeuterung.Size = New System.Drawing.Size(532, 223)
        Me.erlaeuterung.TabIndex = 0
        '
        'sonderAblaeufe
        '
        Me.sonderAblaeufe.Controls.Add(Me.explSonderabl)
        Me.sonderAblaeufe.Location = New System.Drawing.Point(4, 25)
        Me.sonderAblaeufe.Margin = New System.Windows.Forms.Padding(4)
        Me.sonderAblaeufe.Name = "sonderAblaeufe"
        Me.sonderAblaeufe.Padding = New System.Windows.Forms.Padding(4)
        Me.sonderAblaeufe.Size = New System.Drawing.Size(552, 242)
        Me.sonderAblaeufe.TabIndex = 1
        Me.sonderAblaeufe.Text = "Sonderabläufe"
        Me.sonderAblaeufe.ToolTipText = "Anzeige der Sonderabläufe"
        Me.sonderAblaeufe.UseVisualStyleBackColor = True
        '
        'explSonderabl
        '
        Me.explSonderabl.Location = New System.Drawing.Point(8, 7)
        Me.explSonderabl.Margin = New System.Windows.Forms.Padding(4)
        Me.explSonderabl.Multiline = True
        Me.explSonderabl.Name = "explSonderabl"
        Me.explSonderabl.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.explSonderabl.Size = New System.Drawing.Size(532, 223)
        Me.explSonderabl.TabIndex = 0
        '
        'enabler
        '
        Me.enabler.Controls.Add(Me.explEnabler)
        Me.enabler.Location = New System.Drawing.Point(4, 25)
        Me.enabler.Margin = New System.Windows.Forms.Padding(4)
        Me.enabler.Name = "enabler"
        Me.enabler.Size = New System.Drawing.Size(552, 242)
        Me.enabler.TabIndex = 2
        Me.enabler.Text = "Enabler"
        Me.enabler.ToolTipText = "Anzeige der Enabler"
        Me.enabler.UseVisualStyleBackColor = True
        '
        'explEnabler
        '
        Me.explEnabler.Location = New System.Drawing.Point(8, 7)
        Me.explEnabler.Margin = New System.Windows.Forms.Padding(4)
        Me.explEnabler.Multiline = True
        Me.explEnabler.Name = "explEnabler"
        Me.explEnabler.Size = New System.Drawing.Size(532, 223)
        Me.explEnabler.TabIndex = 0
        '
        'zusatzRisiken
        '
        Me.zusatzRisiken.Controls.Add(Me.explRisiken)
        Me.zusatzRisiken.Location = New System.Drawing.Point(4, 25)
        Me.zusatzRisiken.Margin = New System.Windows.Forms.Padding(4)
        Me.zusatzRisiken.Name = "zusatzRisiken"
        Me.zusatzRisiken.Size = New System.Drawing.Size(552, 242)
        Me.zusatzRisiken.TabIndex = 3
        Me.zusatzRisiken.Text = "Zusatzrisiken"
        Me.zusatzRisiken.ToolTipText = "Anzeige der Zusatzrisiken"
        Me.zusatzRisiken.UseVisualStyleBackColor = True
        '
        'explRisiken
        '
        Me.explRisiken.Location = New System.Drawing.Point(8, 7)
        Me.explRisiken.Margin = New System.Windows.Forms.Padding(4)
        Me.explRisiken.Multiline = True
        Me.explRisiken.Name = "explRisiken"
        Me.explRisiken.Size = New System.Drawing.Size(532, 223)
        Me.explRisiken.TabIndex = 0
        '
        'teilnehmer
        '
        Me.teilnehmer.Location = New System.Drawing.Point(4, 25)
        Me.teilnehmer.Margin = New System.Windows.Forms.Padding(4)
        Me.teilnehmer.Name = "teilnehmer"
        Me.teilnehmer.Size = New System.Drawing.Size(552, 242)
        Me.teilnehmer.TabIndex = 4
        Me.teilnehmer.Text = "Teilnehmer"
        Me.teilnehmer.UseVisualStyleBackColor = True
        '
        'frmPhaseInformation
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(120.0!, 120.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.BackColor = System.Drawing.SystemColors.Window
        Me.ClientSize = New System.Drawing.Size(619, 483)
        Me.Controls.Add(Me.lessonsLearnedControl)
        Me.Controls.Add(Me.projectName)
        Me.Controls.Add(Me.phaseDauer)
        Me.Controls.Add(Me.phaseEnde)
        Me.Controls.Add(Me.phaseStart)
        Me.Controls.Add(Me.phaseName)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "frmPhaseInformation"
        Me.Text = "Phasen Information"
        Me.TopMost = True
        Me.lessonsLearnedControl.ResumeLayout(False)
        Me.praemissen.ResumeLayout(False)
        Me.praemissen.PerformLayout()
        Me.sonderAblaeufe.ResumeLayout(False)
        Me.sonderAblaeufe.PerformLayout()
        Me.enabler.ResumeLayout(False)
        Me.enabler.PerformLayout()
        Me.zusatzRisiken.ResumeLayout(False)
        Me.zusatzRisiken.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents phaseName As System.Windows.Forms.TextBox
    Public WithEvents phaseStart As System.Windows.Forms.TextBox
    Public WithEvents phaseEnde As System.Windows.Forms.TextBox
    Public WithEvents phaseDauer As System.Windows.Forms.TextBox
    Public WithEvents projectName As System.Windows.Forms.TextBox
    Friend WithEvents lessonsLearnedControl As System.Windows.Forms.TabControl
    Friend WithEvents praemissen As System.Windows.Forms.TabPage
    Friend WithEvents sonderAblaeufe As System.Windows.Forms.TabPage
    Friend WithEvents enabler As System.Windows.Forms.TabPage
    Friend WithEvents zusatzRisiken As System.Windows.Forms.TabPage
    Public WithEvents erlaeuterung As System.Windows.Forms.TextBox
    Public WithEvents explSonderabl As System.Windows.Forms.TextBox
    Public WithEvents explEnabler As System.Windows.Forms.TextBox
    Public WithEvents explRisiken As System.Windows.Forms.TextBox
    Friend WithEvents teilnehmer As System.Windows.Forms.TabPage
End Class
