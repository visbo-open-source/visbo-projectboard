<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAddOrDeleteALine
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAddOrDeleteALine))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.DeleteALine = New System.Windows.Forms.Button()
        Me.AddALine = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.Panel1.Controls.Add(Me.DeleteALine)
        Me.Panel1.Controls.Add(Me.AddALine)
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(133, 67)
        Me.Panel1.TabIndex = 0
        '
        'DeleteALine
        '
        Me.DeleteALine.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DeleteALine.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.DeleteALine.FlatAppearance.BorderSize = 0
        Me.DeleteALine.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DeleteALine.Location = New System.Drawing.Point(5, 37)
        Me.DeleteALine.Margin = New System.Windows.Forms.Padding(0)
        Me.DeleteALine.Name = "DeleteALine"
        Me.DeleteALine.Size = New System.Drawing.Size(123, 27)
        Me.DeleteALine.TabIndex = 1
        Me.DeleteALine.Text = "Delete A Line"
        Me.DeleteALine.UseVisualStyleBackColor = False
        '
        'AddALine
        '
        Me.AddALine.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.AddALine.BackColor = System.Drawing.SystemColors.ButtonHighlight
        Me.AddALine.FlatAppearance.BorderSize = 0
        Me.AddALine.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AddALine.Location = New System.Drawing.Point(5, 6)
        Me.AddALine.Margin = New System.Windows.Forms.Padding(0)
        Me.AddALine.Name = "AddALine"
        Me.AddALine.Size = New System.Drawing.Size(123, 27)
        Me.AddALine.TabIndex = 0
        Me.AddALine.Text = "Add A Line"
        Me.AddALine.UseVisualStyleBackColor = False
        '
        'frmAddOrDeleteALine
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(134, 67)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmAddOrDeleteALine"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "+/- a Line"
        Me.TopMost = True
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Panel1 As Panel
    Friend WithEvents DeleteALine As Button
    Friend WithEvents AddALine As Button
End Class
