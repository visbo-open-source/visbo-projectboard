<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmProjectEditSettings
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmProjectEditSettings))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.adjustChilds = New System.Windows.Forms.CheckBox()
        Me.invoices = New System.Windows.Forms.CheckBox()
        Me.newCalculation = New System.Windows.Forms.CheckBox()
        Me.AdjustResourceNeeds = New System.Windows.Forms.CheckBox()
        Me.allowOverUtilization = New System.Windows.Forms.CheckBox()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.Controls.Add(Me.allowOverUtilization)
        Me.Panel1.Controls.Add(Me.adjustChilds)
        Me.Panel1.Controls.Add(Me.invoices)
        Me.Panel1.Controls.Add(Me.newCalculation)
        Me.Panel1.Controls.Add(Me.AdjustResourceNeeds)
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(480, 146)
        Me.Panel1.TabIndex = 0
        '
        'adjustChilds
        '
        Me.adjustChilds.AutoSize = True
        Me.adjustChilds.Location = New System.Drawing.Point(12, 87)
        Me.adjustChilds.Name = "adjustChilds"
        Me.adjustChilds.Size = New System.Drawing.Size(305, 19)
        Me.adjustChilds.TabIndex = 3
        Me.adjustChilds.Text = "Date changes also affect dates of subordinate tasks"
        Me.adjustChilds.UseVisualStyleBackColor = True
        '
        'invoices
        '
        Me.invoices.AutoSize = True
        Me.invoices.Location = New System.Drawing.Point(12, 62)
        Me.invoices.Name = "invoices"
        Me.invoices.Size = New System.Drawing.Size(181, 19)
        Me.invoices.TabIndex = 2
        Me.invoices.Text = "Show invoices and penalties"
        Me.invoices.UseVisualStyleBackColor = True
        '
        'newCalculation
        '
        Me.newCalculation.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.newCalculation.AutoSize = True
        Me.newCalculation.Location = New System.Drawing.Point(12, 37)
        Me.newCalculation.Name = "newCalculation"
        Me.newCalculation.Size = New System.Drawing.Size(430, 19)
        Me.newCalculation.TabIndex = 1
        Me.newCalculation.Text = "Distribute values automatically over time  (when phase dates are changed)"
        Me.newCalculation.UseVisualStyleBackColor = True
        '
        'AdjustResourceNeeds
        '
        Me.AdjustResourceNeeds.AutoSize = True
        Me.AdjustResourceNeeds.Location = New System.Drawing.Point(12, 12)
        Me.AdjustResourceNeeds.Name = "AdjustResourceNeeds"
        Me.AdjustResourceNeeds.Size = New System.Drawing.Size(455, 19)
        Me.AdjustResourceNeeds.TabIndex = 0
        Me.AdjustResourceNeeds.Text = "Adjust resource needs proportionally (when phases are extended or shortened)"
        Me.AdjustResourceNeeds.UseVisualStyleBackColor = True
        '
        'allowOverUtilization
        '
        Me.allowOverUtilization.AutoSize = True
        Me.allowOverUtilization.Location = New System.Drawing.Point(12, 114)
        Me.allowOverUtilization.Name = "allowOverUtilization"
        Me.allowOverUtilization.Size = New System.Drawing.Size(137, 19)
        Me.allowOverUtilization.TabIndex = 4
        Me.allowOverUtilization.Text = "Allow over-utilization"
        Me.allowOverUtilization.UseVisualStyleBackColor = True
        '
        'frmProjectEditSettings
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(482, 146)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(351, 64)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmProjectEditSettings"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Settings"
        Me.TopMost = True
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Panel1 As Panel
    Friend WithEvents adjustChilds As CheckBox
    Friend WithEvents invoices As CheckBox
    Friend WithEvents newCalculation As CheckBox
    Friend WithEvents AdjustResourceNeeds As CheckBox
    Friend WithEvents allowOverUtilization As CheckBox
End Class
