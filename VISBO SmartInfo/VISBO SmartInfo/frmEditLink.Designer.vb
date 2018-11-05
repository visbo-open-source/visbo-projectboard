<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEditLink
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
        Me.okBtn = New System.Windows.Forms.Button()
        Me.clearBtn = New System.Windows.Forms.Button()
        Me.linkValue = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'okBtn
        '
        Me.okBtn.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.okBtn.Location = New System.Drawing.Point(134, 95)
        Me.okBtn.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.okBtn.Name = "okBtn"
        Me.okBtn.Size = New System.Drawing.Size(112, 35)
        Me.okBtn.TabIndex = 0
        Me.okBtn.Text = "OK"
        Me.okBtn.UseVisualStyleBackColor = True
        '
        'clearBtn
        '
        Me.clearBtn.DialogResult = System.Windows.Forms.DialogResult.No
        Me.clearBtn.Location = New System.Drawing.Point(378, 95)
        Me.clearBtn.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.clearBtn.Name = "clearBtn"
        Me.clearBtn.Size = New System.Drawing.Size(112, 35)
        Me.clearBtn.TabIndex = 1
        Me.clearBtn.Text = "Clear"
        Me.clearBtn.UseVisualStyleBackColor = True
        '
        'linkValue
        '
        Me.linkValue.Location = New System.Drawing.Point(18, 40)
        Me.linkValue.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.linkValue.Name = "linkValue"
        Me.linkValue.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal
        Me.linkValue.Size = New System.Drawing.Size(644, 26)
        Me.linkValue.TabIndex = 2
        Me.linkValue.WordWrap = False
        '
        'frmEditLink
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(144.0!, 144.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(682, 160)
        Me.Controls.Add(Me.linkValue)
        Me.Controls.Add(Me.clearBtn)
        Me.Controls.Add(Me.okBtn)
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.MaximizeBox = False
        Me.Name = "frmEditLink"
        Me.Text = "Provide Link"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents okBtn As Windows.Forms.Button
    Friend WithEvents clearBtn As Windows.Forms.Button
    Friend WithEvents linkValue As Windows.Forms.TextBox
End Class
