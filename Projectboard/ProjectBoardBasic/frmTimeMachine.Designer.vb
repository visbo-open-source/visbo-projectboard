<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTimeMachine
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
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.btnChangedPosition = New System.Windows.Forms.Button()
        Me.btnHome = New System.Windows.Forms.Button()
        Me.btnBack = New System.Windows.Forms.Button()
        Me.btnFastBack = New System.Windows.Forms.Button()
        Me.btnStart = New System.Windows.Forms.Button()
        Me.btnEnd = New System.Windows.Forms.Button()
        Me.btnFastForward = New System.Windows.Forms.Button()
        Me.btnForward = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(112, 48)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(123, 20)
        Me.TextBox1.TabIndex = 6
        '
        'btnChangedPosition
        '
        Me.btnChangedPosition.Image = Global.ProjectBoardBasic.My.Resources.Resources.replace2
        Me.btnChangedPosition.Location = New System.Drawing.Point(163, 83)
        Me.btnChangedPosition.Name = "btnChangedPosition"
        Me.btnChangedPosition.Size = New System.Drawing.Size(24, 24)
        Me.btnChangedPosition.TabIndex = 14
        Me.btnChangedPosition.UseVisualStyleBackColor = True
        '
        'btnHome
        '
        Me.btnHome.Image = Global.ProjectBoardBasic.My.Resources.Resources.home
        Me.btnHome.Location = New System.Drawing.Point(163, 13)
        Me.btnHome.Name = "btnHome"
        Me.btnHome.Size = New System.Drawing.Size(24, 24)
        Me.btnHome.TabIndex = 13
        Me.btnHome.UseVisualStyleBackColor = True
        '
        'btnBack
        '
        Me.btnBack.Image = Global.ProjectBoardBasic.My.Resources.Resources.navigate_left
        Me.btnBack.Location = New System.Drawing.Point(71, 44)
        Me.btnBack.Name = "btnBack"
        Me.btnBack.Size = New System.Drawing.Size(24, 24)
        Me.btnBack.TabIndex = 12
        Me.btnBack.UseVisualStyleBackColor = True
        '
        'btnFastBack
        '
        Me.btnFastBack.Image = Global.ProjectBoardBasic.My.Resources.Resources.navigate_left2
        Me.btnFastBack.Location = New System.Drawing.Point(41, 44)
        Me.btnFastBack.Name = "btnFastBack"
        Me.btnFastBack.Size = New System.Drawing.Size(24, 24)
        Me.btnFastBack.TabIndex = 11
        Me.btnFastBack.UseVisualStyleBackColor = True
        '
        'btnStart
        '
        Me.btnStart.Image = Global.ProjectBoardBasic.My.Resources.Resources.navigate_beginning
        Me.btnStart.Location = New System.Drawing.Point(12, 44)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(24, 24)
        Me.btnStart.TabIndex = 10
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'btnEnd
        '
        Me.btnEnd.Image = Global.ProjectBoardBasic.My.Resources.Resources.navigate_end
        Me.btnEnd.Location = New System.Drawing.Point(310, 45)
        Me.btnEnd.Name = "btnEnd"
        Me.btnEnd.Size = New System.Drawing.Size(24, 24)
        Me.btnEnd.TabIndex = 9
        Me.btnEnd.UseVisualStyleBackColor = True
        '
        'btnFastForward
        '
        Me.btnFastForward.Image = Global.ProjectBoardBasic.My.Resources.Resources.navigate_right2
        Me.btnFastForward.Location = New System.Drawing.Point(280, 45)
        Me.btnFastForward.Name = "btnFastForward"
        Me.btnFastForward.Size = New System.Drawing.Size(24, 24)
        Me.btnFastForward.TabIndex = 8
        Me.btnFastForward.UseVisualStyleBackColor = True
        '
        'btnForward
        '
        Me.btnForward.Image = Global.ProjectBoardBasic.My.Resources.Resources.navigate_right
        Me.btnForward.Location = New System.Drawing.Point(251, 45)
        Me.btnForward.Name = "btnForward"
        Me.btnForward.Size = New System.Drawing.Size(24, 24)
        Me.btnForward.TabIndex = 7
        Me.btnForward.UseVisualStyleBackColor = True
        '
        'frmTimeMachine
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(348, 126)
        Me.Controls.Add(Me.btnChangedPosition)
        Me.Controls.Add(Me.btnHome)
        Me.Controls.Add(Me.btnBack)
        Me.Controls.Add(Me.btnFastBack)
        Me.Controls.Add(Me.btnStart)
        Me.Controls.Add(Me.btnEnd)
        Me.Controls.Add(Me.btnFastForward)
        Me.Controls.Add(Me.btnForward)
        Me.Controls.Add(Me.TextBox1)
        Me.Name = "frmTimeMachine"
        Me.Text = "VISBO Time Machine"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents btnForward As System.Windows.Forms.Button
    Friend WithEvents btnFastForward As System.Windows.Forms.Button
    Friend WithEvents btnEnd As System.Windows.Forms.Button
    Friend WithEvents btnBack As System.Windows.Forms.Button
    Friend WithEvents btnFastBack As System.Windows.Forms.Button
    Friend WithEvents btnStart As System.Windows.Forms.Button
    Friend WithEvents btnHome As System.Windows.Forms.Button
    Friend WithEvents btnChangedPosition As System.Windows.Forms.Button
End Class
