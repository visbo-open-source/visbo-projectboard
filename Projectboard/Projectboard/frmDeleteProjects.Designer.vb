<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDeleteProjects
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
        Me.TreeViewProjekte = New System.Windows.Forms.TreeView()
        Me.DeleteButton = New System.Windows.Forms.Button()
        Me.AbbrechenButton = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'TreeViewProjekte
        '
        Me.TreeViewProjekte.Location = New System.Drawing.Point(45, 25)
        Me.TreeViewProjekte.Name = "TreeViewProjekte"
        Me.TreeViewProjekte.Size = New System.Drawing.Size(479, 357)
        Me.TreeViewProjekte.TabIndex = 0
        '
        'DeleteButton
        '
        Me.DeleteButton.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.DeleteButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DeleteButton.Location = New System.Drawing.Point(93, 419)
        Me.DeleteButton.Name = "DeleteButton"
        Me.DeleteButton.Size = New System.Drawing.Size(146, 31)
        Me.DeleteButton.TabIndex = 1
        Me.DeleteButton.Text = "Löschen"
        Me.DeleteButton.UseVisualStyleBackColor = True
        '
        'AbbrechenButton
        '
        Me.AbbrechenButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.AbbrechenButton.Location = New System.Drawing.Point(334, 419)
        Me.AbbrechenButton.Name = "AbbrechenButton"
        Me.AbbrechenButton.Size = New System.Drawing.Size(139, 31)
        Me.AbbrechenButton.TabIndex = 2
        Me.AbbrechenButton.Text = "Abbrechen"
        Me.AbbrechenButton.UseVisualStyleBackColor = True
        '
        'frmDeleteProjects
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(120.0!, 120.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(584, 474)
        Me.Controls.Add(Me.AbbrechenButton)
        Me.Controls.Add(Me.DeleteButton)
        Me.Controls.Add(Me.TreeViewProjekte)
        Me.Name = "frmDeleteProjects"
        Me.Text = "frmDeleteProjects"
        Me.ResumeLayout(False)

    End Sub
    Public WithEvents TreeViewProjekte As System.Windows.Forms.TreeView
    Public WithEvents DeleteButton As System.Windows.Forms.Button
    Public WithEvents AbbrechenButton As System.Windows.Forms.Button
End Class
