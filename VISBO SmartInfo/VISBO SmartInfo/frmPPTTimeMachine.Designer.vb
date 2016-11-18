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
        Me.txtboxCurrentDate = New System.Windows.Forms.TextBox()
        Me.btnChangedPosition = New System.Windows.Forms.Button()
        Me.btnHome = New System.Windows.Forms.Button()
        Me.btnBack = New System.Windows.Forms.Button()
        Me.btnFastBack = New System.Windows.Forms.Button()
        Me.btnStart = New System.Windows.Forms.Button()
        Me.btnEnd = New System.Windows.Forms.Button()
        Me.btnFastForward = New System.Windows.Forms.Button()
        Me.btnForward = New System.Windows.Forms.Button()
        Me.lblMessage = New System.Windows.Forms.Label()
        Me.ProgressBarNavigate = New System.Windows.Forms.ProgressBar()
        Me.ToolTipTS = New System.Windows.Forms.ToolTip(Me.components)
        Me.SuspendLayout()
        '
        'txtboxCurrentDate
        '
        Me.txtboxCurrentDate.Location = New System.Drawing.Point(113, 42)
        Me.txtboxCurrentDate.Name = "txtboxCurrentDate"
        Me.txtboxCurrentDate.Size = New System.Drawing.Size(123, 20)
        Me.txtboxCurrentDate.TabIndex = 15
        '
        'btnChangedPosition
        '
        Me.btnChangedPosition.Image = Global.VISBO_SmartInfo.My.Resources.Resources.replace2
        Me.btnChangedPosition.Location = New System.Drawing.Point(164, 69)
        Me.btnChangedPosition.Name = "btnChangedPosition"
        Me.btnChangedPosition.Size = New System.Drawing.Size(24, 24)
        Me.btnChangedPosition.TabIndex = 23
        Me.btnChangedPosition.UseVisualStyleBackColor = True
        '
        'btnHome
        '
        Me.btnHome.Image = Global.VISBO_SmartInfo.My.Resources.Resources.home
        Me.btnHome.Location = New System.Drawing.Point(164, 10)
        Me.btnHome.Name = "btnHome"
        Me.btnHome.Size = New System.Drawing.Size(24, 24)
        Me.btnHome.TabIndex = 22
        Me.btnHome.UseVisualStyleBackColor = True
        '
        'btnBack
        '
        Me.btnBack.Image = Global.VISBO_SmartInfo.My.Resources.Resources.navigate_left
        Me.btnBack.Location = New System.Drawing.Point(72, 38)
        Me.btnBack.Name = "btnBack"
        Me.btnBack.Size = New System.Drawing.Size(24, 24)
        Me.btnBack.TabIndex = 21
        Me.btnBack.UseVisualStyleBackColor = True
        '
        'btnFastBack
        '
        Me.btnFastBack.Image = Global.VISBO_SmartInfo.My.Resources.Resources.navigate_left2
        Me.btnFastBack.Location = New System.Drawing.Point(42, 38)
        Me.btnFastBack.Name = "btnFastBack"
        Me.btnFastBack.Size = New System.Drawing.Size(24, 24)
        Me.btnFastBack.TabIndex = 20
        Me.btnFastBack.UseVisualStyleBackColor = True
        '
        'btnStart
        '
        Me.btnStart.Image = Global.VISBO_SmartInfo.My.Resources.Resources.navigate_beginning1
        Me.btnStart.Location = New System.Drawing.Point(13, 38)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(24, 24)
        Me.btnStart.TabIndex = 19
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'btnEnd
        '
        Me.btnEnd.Image = Global.VISBO_SmartInfo.My.Resources.Resources.navigate_end
        Me.btnEnd.Location = New System.Drawing.Point(311, 39)
        Me.btnEnd.Name = "btnEnd"
        Me.btnEnd.Size = New System.Drawing.Size(24, 24)
        Me.btnEnd.TabIndex = 18
        Me.btnEnd.UseVisualStyleBackColor = True
        '
        'btnFastForward
        '
        Me.btnFastForward.Image = Global.VISBO_SmartInfo.My.Resources.Resources.navigate_right2
        Me.btnFastForward.Location = New System.Drawing.Point(281, 39)
        Me.btnFastForward.Name = "btnFastForward"
        Me.btnFastForward.Size = New System.Drawing.Size(24, 24)
        Me.btnFastForward.TabIndex = 17
        Me.btnFastForward.UseVisualStyleBackColor = True
        '
        'btnForward
        '
        Me.btnForward.Image = Global.VISBO_SmartInfo.My.Resources.Resources.navigate_right
        Me.btnForward.Location = New System.Drawing.Point(252, 39)
        Me.btnForward.Name = "btnForward"
        Me.btnForward.Size = New System.Drawing.Size(24, 24)
        Me.btnForward.TabIndex = 16
        Me.btnForward.UseVisualStyleBackColor = True
        '
        'lblMessage
        '
        Me.lblMessage.AutoSize = True
        Me.lblMessage.Location = New System.Drawing.Point(12, 79)
        Me.lblMessage.Name = "lblMessage"
        Me.lblMessage.Size = New System.Drawing.Size(50, 13)
        Me.lblMessage.TabIndex = 24
        Me.lblMessage.Text = "Message"
        '
        'ProgressBarNavigate
        '
        Me.ProgressBarNavigate.Location = New System.Drawing.Point(138, 94)
        Me.ProgressBarNavigate.Name = "ProgressBarNavigate"
        Me.ProgressBarNavigate.Size = New System.Drawing.Size(79, 10)
        Me.ProgressBarNavigate.TabIndex = 25
        Me.ProgressBarNavigate.UseWaitCursor = True
        '
        'frmPPTTimeMachine
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(357, 106)
        Me.Controls.Add(Me.ProgressBarNavigate)
        Me.Controls.Add(Me.lblMessage)
        Me.Controls.Add(Me.btnChangedPosition)
        Me.Controls.Add(Me.btnHome)
        Me.Controls.Add(Me.btnBack)
        Me.Controls.Add(Me.btnFastBack)
        Me.Controls.Add(Me.btnStart)
        Me.Controls.Add(Me.btnEnd)
        Me.Controls.Add(Me.btnFastForward)
        Me.Controls.Add(Me.btnForward)
        Me.Controls.Add(Me.txtboxCurrentDate)
        Me.Name = "frmPPTTimeMachine"
        Me.Text = "VISBO Time Machine"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnChangedPosition As System.Windows.Forms.Button
    Friend WithEvents btnHome As System.Windows.Forms.Button
    Friend WithEvents btnBack As System.Windows.Forms.Button
    Friend WithEvents btnFastBack As System.Windows.Forms.Button
    Friend WithEvents btnStart As System.Windows.Forms.Button
    Friend WithEvents btnEnd As System.Windows.Forms.Button
    Friend WithEvents btnFastForward As System.Windows.Forms.Button
    Friend WithEvents btnForward As System.Windows.Forms.Button
    Friend WithEvents txtboxCurrentDate As System.Windows.Forms.TextBox
    Friend WithEvents lblMessage As System.Windows.Forms.Label
    Friend WithEvents ProgressBarNavigate As System.Windows.Forms.ProgressBar
    Friend WithEvents ToolTipTS As System.Windows.Forms.ToolTip
End Class
