<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRenameProject
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmRenameProject))
        Me.oldName = New System.Windows.Forms.TextBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.newName = New System.Windows.Forms.TextBox()
        Me.Ok_Button = New System.Windows.Forms.Button()
        Me.Cancel_Button = New System.Windows.Forms.Button()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'oldName
        '
        Me.oldName.Enabled = False
        Me.oldName.Location = New System.Drawing.Point(12, 24)
        Me.oldName.Name = "oldName"
        Me.oldName.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal
        Me.oldName.Size = New System.Drawing.Size(433, 20)
        Me.oldName.TabIndex = 0
        Me.oldName.WordWrap = False
        '
        'PictureBox1
        '
        Me.PictureBox1.Enabled = False
        Me.PictureBox1.Location = New System.Drawing.Point(195, 54)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(35, 32)
        Me.PictureBox1.TabIndex = 1
        Me.PictureBox1.TabStop = False
        '
        'newName
        '
        Me.newName.Location = New System.Drawing.Point(12, 92)
        Me.newName.Name = "newName"
        Me.newName.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal
        Me.newName.Size = New System.Drawing.Size(433, 20)
        Me.newName.TabIndex = 2
        Me.newName.WordWrap = False
        '
        'Ok_Button
        '
        Me.Ok_Button.Location = New System.Drawing.Point(88, 137)
        Me.Ok_Button.Name = "Ok_Button"
        Me.Ok_Button.Size = New System.Drawing.Size(83, 23)
        Me.Ok_Button.TabIndex = 3
        Me.Ok_Button.Text = "OK"
        Me.Ok_Button.UseVisualStyleBackColor = True
        '
        'Cancel_Button
        '
        Me.Cancel_Button.Location = New System.Drawing.Point(247, 137)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(75, 23)
        Me.Cancel_Button.TabIndex = 4
        Me.Cancel_Button.Text = "Abbrechen"
        Me.Cancel_Button.UseVisualStyleBackColor = True
        '
        'frmRenameProject
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(457, 175)
        Me.Controls.Add(Me.Cancel_Button)
        Me.Controls.Add(Me.Ok_Button)
        Me.Controls.Add(Me.newName)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.oldName)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmRenameProject"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Rename"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Ok_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Public WithEvents oldName As System.Windows.Forms.TextBox
    Public WithEvents newName As System.Windows.Forms.TextBox
End Class
