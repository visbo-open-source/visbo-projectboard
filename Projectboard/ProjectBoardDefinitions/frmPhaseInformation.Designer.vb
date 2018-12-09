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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmPhaseInformation))
        Me.phaseName = New System.Windows.Forms.TextBox()
        Me.phaseStart = New System.Windows.Forms.TextBox()
        Me.projectName = New System.Windows.Forms.TextBox()
        Me.breadCrumb = New System.Windows.Forms.TextBox()
        Me.showOrigItem = New System.Windows.Forms.CheckBox()
        Me.phaseEnde = New System.Windows.Forms.TextBox()
        Me.phaseDauer = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'phaseName
        '
        Me.phaseName.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.phaseName.Location = New System.Drawing.Point(24, 80)
        Me.phaseName.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.phaseName.Name = "phaseName"
        Me.phaseName.ReadOnly = True
        Me.phaseName.Size = New System.Drawing.Size(559, 30)
        Me.phaseName.TabIndex = 1
        '
        'phaseStart
        '
        Me.phaseStart.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.phaseStart.Location = New System.Drawing.Point(24, 150)
        Me.phaseStart.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.phaseStart.Name = "phaseStart"
        Me.phaseStart.ReadOnly = True
        Me.phaseStart.Size = New System.Drawing.Size(175, 26)
        Me.phaseStart.TabIndex = 2
        '
        'projectName
        '
        Me.projectName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.projectName.Location = New System.Drawing.Point(24, 32)
        Me.projectName.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.projectName.Name = "projectName"
        Me.projectName.ReadOnly = True
        Me.projectName.Size = New System.Drawing.Size(175, 26)
        Me.projectName.TabIndex = 21
        '
        'breadCrumb
        '
        Me.breadCrumb.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.breadCrumb.Location = New System.Drawing.Point(219, 32)
        Me.breadCrumb.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.breadCrumb.Name = "breadCrumb"
        Me.breadCrumb.ReadOnly = True
        Me.breadCrumb.Size = New System.Drawing.Size(364, 26)
        Me.breadCrumb.TabIndex = 22
        '
        'showOrigItem
        '
        Me.showOrigItem.AutoSize = True
        Me.showOrigItem.Location = New System.Drawing.Point(24, 120)
        Me.showOrigItem.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.showOrigItem.Name = "showOrigItem"
        Me.showOrigItem.Size = New System.Drawing.Size(120, 21)
        Me.showOrigItem.TabIndex = 28
        Me.showOrigItem.Text = "Original Name"
        Me.showOrigItem.UseVisualStyleBackColor = True
        '
        'phaseEnde
        '
        Me.phaseEnde.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.phaseEnde.Location = New System.Drawing.Point(408, 150)
        Me.phaseEnde.Margin = New System.Windows.Forms.Padding(4)
        Me.phaseEnde.Name = "phaseEnde"
        Me.phaseEnde.ReadOnly = True
        Me.phaseEnde.Size = New System.Drawing.Size(175, 26)
        Me.phaseEnde.TabIndex = 4
        '
        'phaseDauer
        '
        Me.phaseDauer.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.phaseDauer.Location = New System.Drawing.Point(219, 150)
        Me.phaseDauer.Margin = New System.Windows.Forms.Padding(4)
        Me.phaseDauer.Name = "phaseDauer"
        Me.phaseDauer.ReadOnly = True
        Me.phaseDauer.Size = New System.Drawing.Size(169, 26)
        Me.phaseDauer.TabIndex = 3
        '
        'frmPhaseInformation
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(120.0!, 120.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoSize = True
        Me.BackColor = System.Drawing.SystemColors.Window
        Me.ClientSize = New System.Drawing.Size(602, 205)
        Me.Controls.Add(Me.showOrigItem)
        Me.Controls.Add(Me.breadCrumb)
        Me.Controls.Add(Me.projectName)
        Me.Controls.Add(Me.phaseDauer)
        Me.Controls.Add(Me.phaseEnde)
        Me.Controls.Add(Me.phaseStart)
        Me.Controls.Add(Me.phaseName)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Name = "frmPhaseInformation"
        Me.Text = "Phasen Information"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents phaseName As System.Windows.Forms.TextBox
    Public WithEvents phaseStart As System.Windows.Forms.TextBox
    Public WithEvents projectName As System.Windows.Forms.TextBox
    Public WithEvents breadCrumb As System.Windows.Forms.TextBox
    Public WithEvents showOrigItem As System.Windows.Forms.CheckBox
    Public WithEvents phaseEnde As Windows.Forms.TextBox
    Public WithEvents phaseDauer As Windows.Forms.TextBox
End Class
