<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPPTTimeMachine
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmPPTTimeMachine))
        Me.lblMessage = New System.Windows.Forms.Label()
        Me.ToolTipTS = New System.Windows.Forms.ToolTip(Me.components)
        Me.currentDate = New System.Windows.Forms.DateTimePicker()
        Me.btnFastBack = New System.Windows.Forms.Button()
        Me.btnStart = New System.Windows.Forms.Button()
        Me.btnEnd = New System.Windows.Forms.Button()
        Me.btnFastForward = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lblMessage
        '
        Me.lblMessage.AutoSize = True
        Me.lblMessage.Location = New System.Drawing.Point(28, 236)
        Me.lblMessage.Name = "lblMessage"
        Me.lblMessage.Size = New System.Drawing.Size(50, 13)
        Me.lblMessage.TabIndex = 24
        Me.lblMessage.Text = "Message"
        '
        'currentDate
        '
        Me.currentDate.CausesValidation = False
        Me.currentDate.Checked = False
        Me.currentDate.Location = New System.Drawing.Point(30, 171)
        Me.currentDate.Name = "currentDate"
        Me.currentDate.Size = New System.Drawing.Size(192, 20)
        Me.currentDate.TabIndex = 26
        '
        'btnFastBack
        '
        Me.btnFastBack.Image = Global.VISBO_SmartInfo.My.Resources.Resources.navigate_left
        Me.btnFastBack.Location = New System.Drawing.Point(110, 197)
        Me.btnFastBack.Name = "btnFastBack"
        Me.btnFastBack.Size = New System.Drawing.Size(32, 32)
        Me.btnFastBack.TabIndex = 20
        Me.btnFastBack.UseVisualStyleBackColor = True
        '
        'btnStart
        '
        Me.btnStart.Image = Global.VISBO_SmartInfo.My.Resources.Resources.navigate_beginning1
        Me.btnStart.Location = New System.Drawing.Point(30, 197)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(32, 32)
        Me.btnStart.TabIndex = 19
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'btnEnd
        '
        Me.btnEnd.Image = Global.VISBO_SmartInfo.My.Resources.Resources.Calendar_icon_128x128_Last
        Me.btnEnd.Location = New System.Drawing.Point(30, 12)
        Me.btnEnd.Name = "btnEnd"
        Me.btnEnd.Size = New System.Drawing.Size(192, 152)
        Me.btnEnd.TabIndex = 18
        Me.btnEnd.UseVisualStyleBackColor = True
        '
        'btnFastForward
        '
        Me.btnFastForward.Image = Global.VISBO_SmartInfo.My.Resources.Resources.navigate_right
        Me.btnFastForward.Location = New System.Drawing.Point(190, 197)
        Me.btnFastForward.Name = "btnFastForward"
        Me.btnFastForward.Size = New System.Drawing.Size(32, 32)
        Me.btnFastForward.TabIndex = 17
        Me.btnFastForward.UseVisualStyleBackColor = True
        '
        'frmPPTTimeMachine
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.ClientSize = New System.Drawing.Size(260, 258)
        Me.Controls.Add(Me.currentDate)
        Me.Controls.Add(Me.lblMessage)
        Me.Controls.Add(Me.btnFastBack)
        Me.Controls.Add(Me.btnStart)
        Me.Controls.Add(Me.btnEnd)
        Me.Controls.Add(Me.btnFastForward)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmPPTTimeMachine"
        Me.Text = "VISBO Time Machine"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnFastBack As System.Windows.Forms.Button
    Friend WithEvents btnStart As System.Windows.Forms.Button
    Friend WithEvents btnEnd As System.Windows.Forms.Button
    Friend WithEvents btnFastForward As System.Windows.Forms.Button
    Friend WithEvents lblMessage As System.Windows.Forms.Label
    Friend WithEvents ToolTipTS As System.Windows.Forms.ToolTip
    Friend WithEvents currentDate As System.Windows.Forms.DateTimePicker
End Class
