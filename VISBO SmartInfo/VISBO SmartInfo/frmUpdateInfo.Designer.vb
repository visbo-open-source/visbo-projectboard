<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmUpdateInfo
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
        Me.updateMsg = New System.Windows.Forms.Label()
        Me.update_btn = New System.Windows.Forms.Button()
        Me.cancel_btn = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'updateMsg
        '
        Me.updateMsg.AutoSize = True
        Me.updateMsg.Location = New System.Drawing.Point(35, 30)
        Me.updateMsg.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.updateMsg.Name = "updateMsg"
        Me.updateMsg.Size = New System.Drawing.Size(51, 17)
        Me.updateMsg.TabIndex = 0
        Me.updateMsg.Text = "Label1"
        '
        'update_btn
        '
        Me.update_btn.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.update_btn.Location = New System.Drawing.Point(35, 122)
        Me.update_btn.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.update_btn.Name = "update_btn"
        Me.update_btn.Size = New System.Drawing.Size(100, 28)
        Me.update_btn.TabIndex = 1
        Me.update_btn.Text = "Update"
        Me.update_btn.UseVisualStyleBackColor = True
        '
        'cancel_btn
        '
        Me.cancel_btn.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cancel_btn.Location = New System.Drawing.Point(212, 122)
        Me.cancel_btn.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cancel_btn.Name = "cancel_btn"
        Me.cancel_btn.Size = New System.Drawing.Size(100, 28)
        Me.cancel_btn.TabIndex = 2
        Me.cancel_btn.Text = "Cancel"
        Me.cancel_btn.UseVisualStyleBackColor = True
        '
        'frmUpdateInfo
        '
        Me.AcceptButton = Me.update_btn
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.CancelButton = Me.cancel_btn
        Me.ClientSize = New System.Drawing.Size(349, 171)
        Me.Controls.Add(Me.cancel_btn)
        Me.Controls.Add(Me.update_btn)
        Me.Controls.Add(Me.updateMsg)
        Me.Location = New System.Drawing.Point(30, 200)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmUpdateInfo"
        Me.Text = "Update VISBO Smart Slides?"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents update_btn As Windows.Forms.Button
    Friend WithEvents cancel_btn As Windows.Forms.Button
    Public WithEvents updateMsg As Windows.Forms.Label
End Class
