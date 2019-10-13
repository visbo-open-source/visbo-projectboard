<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSelectPhasesMilestones
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
        Me.TreeViewProjects = New System.Windows.Forms.TreeView()
        Me.Ok_Button = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'TreeViewProjects
        '
        Me.TreeViewProjects.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TreeViewProjects.Location = New System.Drawing.Point(9, 12)
        Me.TreeViewProjects.Name = "TreeViewProjects"
        Me.TreeViewProjects.Size = New System.Drawing.Size(471, 280)
        Me.TreeViewProjects.TabIndex = 0
        '
        'Ok_Button
        '
        Me.Ok_Button.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Ok_Button.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Ok_Button.Location = New System.Drawing.Point(165, 312)
        Me.Ok_Button.Name = "Ok_Button"
        Me.Ok_Button.Size = New System.Drawing.Size(157, 23)
        Me.Ok_Button.TabIndex = 4
        Me.Ok_Button.Text = "Auswahl bestätigen"
        Me.Ok_Button.UseVisualStyleBackColor = True
        '
        'frmSelectPhasesMilestones
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(487, 347)
        Me.Controls.Add(Me.Ok_Button)
        Me.Controls.Add(Me.TreeViewProjects)
        Me.Name = "frmSelectPhasesMilestones"
        Me.Text = "Auswahl von Projekten, Phasen und Meilensteinen"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TreeViewProjects As Windows.Forms.TreeView
    Friend WithEvents Ok_Button As Windows.Forms.Button
End Class
