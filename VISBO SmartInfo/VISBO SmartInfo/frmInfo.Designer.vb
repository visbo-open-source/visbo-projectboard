<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmInfo
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmInfo))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.positionDateButton = New System.Windows.Forms.Button()
        Me.deleteDate = New System.Windows.Forms.Button()
        Me.writeDate = New System.Windows.Forms.Button()
        Me.elemDate = New System.Windows.Forms.TextBox()
        Me.positionTextButton = New System.Windows.Forms.Button()
        Me.deleteText = New System.Windows.Forms.Button()
        Me.showOrginalName = New System.Windows.Forms.CheckBox()
        Me.elemName = New System.Windows.Forms.TextBox()
        Me.showAbbrev = New System.Windows.Forms.CheckBox()
        Me.writeText = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'positionDateButton
        '
        Me.positionDateButton.Image = Global.VISBO_SmartInfo.My.Resources.Resources.layout_south
        Me.positionDateButton.Location = New System.Drawing.Point(455, 64)
        Me.positionDateButton.Name = "positionDateButton"
        Me.positionDateButton.Size = New System.Drawing.Size(30, 26)
        Me.positionDateButton.TabIndex = 58
        Me.positionDateButton.UseVisualStyleBackColor = True
        '
        'deleteDate
        '
        Me.deleteDate.Image = Global.VISBO_SmartInfo.My.Resources.Resources.selection_delete
        Me.deleteDate.Location = New System.Drawing.Point(425, 64)
        Me.deleteDate.Name = "deleteDate"
        Me.deleteDate.Size = New System.Drawing.Size(30, 26)
        Me.deleteDate.TabIndex = 57
        Me.deleteDate.UseVisualStyleBackColor = True
        '
        'writeDate
        '
        Me.writeDate.Image = Global.VISBO_SmartInfo.My.Resources.Resources.pen_blue
        Me.writeDate.Location = New System.Drawing.Point(485, 64)
        Me.writeDate.Name = "writeDate"
        Me.writeDate.Size = New System.Drawing.Size(30, 26)
        Me.writeDate.TabIndex = 56
        Me.writeDate.UseVisualStyleBackColor = True
        '
        'elemDate
        '
        Me.elemDate.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.elemDate.Enabled = False
        Me.elemDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.elemDate.Location = New System.Drawing.Point(12, 65)
        Me.elemDate.Name = "elemDate"
        Me.elemDate.Size = New System.Drawing.Size(407, 22)
        Me.elemDate.TabIndex = 55
        '
        'positionTextButton
        '
        Me.positionTextButton.Image = Global.VISBO_SmartInfo.My.Resources.Resources.layout_north
        Me.positionTextButton.Location = New System.Drawing.Point(455, 32)
        Me.positionTextButton.Name = "positionTextButton"
        Me.positionTextButton.Size = New System.Drawing.Size(30, 26)
        Me.positionTextButton.TabIndex = 54
        Me.positionTextButton.UseVisualStyleBackColor = True
        '
        'deleteText
        '
        Me.deleteText.Image = Global.VISBO_SmartInfo.My.Resources.Resources.selection_delete
        Me.deleteText.Location = New System.Drawing.Point(425, 32)
        Me.deleteText.Name = "deleteText"
        Me.deleteText.Size = New System.Drawing.Size(30, 26)
        Me.deleteText.TabIndex = 53
        Me.deleteText.UseVisualStyleBackColor = True
        '
        'showOrginalName
        '
        Me.showOrginalName.AutoSize = True
        Me.showOrginalName.Location = New System.Drawing.Point(108, 10)
        Me.showOrginalName.Name = "showOrginalName"
        Me.showOrginalName.Size = New System.Drawing.Size(92, 17)
        Me.showOrginalName.TabIndex = 52
        Me.showOrginalName.Text = "Original-Name"
        Me.showOrginalName.UseVisualStyleBackColor = True
        Me.showOrginalName.Visible = False
        '
        'elemName
        '
        Me.elemName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.elemName.Enabled = False
        Me.elemName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.elemName.Location = New System.Drawing.Point(12, 33)
        Me.elemName.Name = "elemName"
        Me.elemName.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal
        Me.elemName.Size = New System.Drawing.Size(407, 22)
        Me.elemName.TabIndex = 49
        '
        'showAbbrev
        '
        Me.showAbbrev.AutoSize = True
        Me.showAbbrev.Location = New System.Drawing.Point(12, 10)
        Me.showAbbrev.Name = "showAbbrev"
        Me.showAbbrev.Size = New System.Drawing.Size(77, 17)
        Me.showAbbrev.TabIndex = 51
        Me.showAbbrev.Text = "Abkürzung"
        Me.showAbbrev.UseVisualStyleBackColor = True
        '
        'writeText
        '
        Me.writeText.Image = Global.VISBO_SmartInfo.My.Resources.Resources.pen_blue
        Me.writeText.Location = New System.Drawing.Point(485, 32)
        Me.writeText.Name = "writeText"
        Me.writeText.Size = New System.Drawing.Size(30, 26)
        Me.writeText.TabIndex = 50
        Me.writeText.UseVisualStyleBackColor = True
        '
        'frmInfo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(532, 101)
        Me.Controls.Add(Me.positionDateButton)
        Me.Controls.Add(Me.deleteDate)
        Me.Controls.Add(Me.writeDate)
        Me.Controls.Add(Me.elemDate)
        Me.Controls.Add(Me.positionTextButton)
        Me.Controls.Add(Me.deleteText)
        Me.Controls.Add(Me.showOrginalName)
        Me.Controls.Add(Me.elemName)
        Me.Controls.Add(Me.showAbbrev)
        Me.Controls.Add(Me.writeText)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmInfo"
        Me.Text = "Beschriften"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents positionDateButton As System.Windows.Forms.Button
    Friend WithEvents deleteDate As System.Windows.Forms.Button
    Friend WithEvents writeDate As System.Windows.Forms.Button
    Friend WithEvents elemDate As System.Windows.Forms.TextBox
    Friend WithEvents positionTextButton As System.Windows.Forms.Button
    Friend WithEvents deleteText As System.Windows.Forms.Button
    Friend WithEvents showOrginalName As System.Windows.Forms.CheckBox
    Friend WithEvents elemName As System.Windows.Forms.TextBox
    Friend WithEvents showAbbrev As System.Windows.Forms.CheckBox
    Friend WithEvents writeText As System.Windows.Forms.Button
End Class
