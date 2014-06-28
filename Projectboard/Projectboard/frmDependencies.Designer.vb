<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDependencies
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.dependentProjectList = New System.Windows.Forms.ListBox()
        Me.ProjectList = New System.Windows.Forms.ListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.degree = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.AbbruchButton = New System.Windows.Forms.Button()
        Me.moveFromDependent = New System.Windows.Forms.Button()
        Me.copyFromDependent = New System.Windows.Forms.Button()
        Me.deleteFromDependent = New System.Windows.Forms.Button()
        Me.deleteFromProjects = New System.Windows.Forms.Button()
        Me.copyFromProjects = New System.Windows.Forms.Button()
        Me.moveFromProjects = New System.Windows.Forms.Button()
        Me.description = New System.Windows.Forms.TextBox()
        Me.statusMeldung = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'dependentProjectList
        '
        Me.dependentProjectList.AllowDrop = True
        Me.dependentProjectList.FormattingEnabled = True
        Me.dependentProjectList.HorizontalScrollbar = True
        Me.dependentProjectList.ItemHeight = 16
        Me.dependentProjectList.Location = New System.Drawing.Point(39, 70)
        Me.dependentProjectList.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.dependentProjectList.Name = "dependentProjectList"
        Me.dependentProjectList.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.dependentProjectList.Size = New System.Drawing.Size(259, 196)
        Me.dependentProjectList.Sorted = True
        Me.dependentProjectList.TabIndex = 0
        '
        'ProjectList
        '
        Me.ProjectList.AllowDrop = True
        Me.ProjectList.FormattingEnabled = True
        Me.ProjectList.ItemHeight = 16
        Me.ProjectList.Location = New System.Drawing.Point(491, 70)
        Me.ProjectList.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ProjectList.Name = "ProjectList"
        Me.ProjectList.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.ProjectList.Size = New System.Drawing.Size(259, 196)
        Me.ProjectList.Sorted = True
        Me.ProjectList.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(129, 10)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(574, 17)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "die sel. ""abhängigen Projekte"" sind von den sel. Projekten der anderen Gruppe abh" & _
    "ängig;"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(365, 124)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(60, 17)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "ist / sind"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(347, 178)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(94, 17)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "abhängig von"
        '
        'degree
        '
        Me.degree.FormattingEnabled = True
        Me.degree.Location = New System.Drawing.Point(315, 145)
        Me.degree.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.degree.Name = "degree"
        Me.degree.Size = New System.Drawing.Size(160, 24)
        Me.degree.TabIndex = 6
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(37, 50)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(131, 17)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "abhängige Projekte"
        '
        'OKButton
        '
        Me.OKButton.Location = New System.Drawing.Point(241, 409)
        Me.OKButton.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(100, 28)
        Me.OKButton.TabIndex = 8
        Me.OKButton.Text = "OK"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'AbbruchButton
        '
        Me.AbbruchButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.AbbruchButton.Location = New System.Drawing.Point(435, 409)
        Me.AbbruchButton.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.AbbruchButton.Name = "AbbruchButton"
        Me.AbbruchButton.Size = New System.Drawing.Size(100, 28)
        Me.AbbruchButton.TabIndex = 9
        Me.AbbruchButton.Text = "Abbrechen"
        Me.AbbruchButton.UseVisualStyleBackColor = True
        '
        'moveFromDependent
        '
        Me.moveFromDependent.Location = New System.Drawing.Point(83, 274)
        Me.moveFromDependent.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.moveFromDependent.Name = "moveFromDependent"
        Me.moveFromDependent.Size = New System.Drawing.Size(36, 28)
        Me.moveFromDependent.TabIndex = 10
        Me.moveFromDependent.Text = "m"
        Me.moveFromDependent.UseVisualStyleBackColor = True
        '
        'copyFromDependent
        '
        Me.copyFromDependent.Location = New System.Drawing.Point(127, 274)
        Me.copyFromDependent.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.copyFromDependent.Name = "copyFromDependent"
        Me.copyFromDependent.Size = New System.Drawing.Size(36, 28)
        Me.copyFromDependent.TabIndex = 11
        Me.copyFromDependent.Text = "c"
        Me.copyFromDependent.UseVisualStyleBackColor = True
        '
        'deleteFromDependent
        '
        Me.deleteFromDependent.Location = New System.Drawing.Point(171, 274)
        Me.deleteFromDependent.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.deleteFromDependent.Name = "deleteFromDependent"
        Me.deleteFromDependent.Size = New System.Drawing.Size(36, 28)
        Me.deleteFromDependent.TabIndex = 12
        Me.deleteFromDependent.Text = "d"
        Me.deleteFromDependent.UseVisualStyleBackColor = True
        '
        'deleteFromProjects
        '
        Me.deleteFromProjects.Location = New System.Drawing.Point(677, 274)
        Me.deleteFromProjects.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.deleteFromProjects.Name = "deleteFromProjects"
        Me.deleteFromProjects.Size = New System.Drawing.Size(36, 28)
        Me.deleteFromProjects.TabIndex = 15
        Me.deleteFromProjects.Text = "d"
        Me.deleteFromProjects.UseVisualStyleBackColor = True
        '
        'copyFromProjects
        '
        Me.copyFromProjects.Location = New System.Drawing.Point(633, 274)
        Me.copyFromProjects.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.copyFromProjects.Name = "copyFromProjects"
        Me.copyFromProjects.Size = New System.Drawing.Size(36, 28)
        Me.copyFromProjects.TabIndex = 14
        Me.copyFromProjects.Text = "c"
        Me.copyFromProjects.UseVisualStyleBackColor = True
        '
        'moveFromProjects
        '
        Me.moveFromProjects.Location = New System.Drawing.Point(589, 274)
        Me.moveFromProjects.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.moveFromProjects.Name = "moveFromProjects"
        Me.moveFromProjects.Size = New System.Drawing.Size(36, 28)
        Me.moveFromProjects.TabIndex = 13
        Me.moveFromProjects.Text = "m"
        Me.moveFromProjects.UseVisualStyleBackColor = True
        '
        'description
        '
        Me.description.Location = New System.Drawing.Point(39, 314)
        Me.description.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.description.Multiline = True
        Me.description.Name = "description"
        Me.description.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.description.Size = New System.Drawing.Size(708, 80)
        Me.description.TabIndex = 16
        Me.description.WordWrap = False
        '
        'statusMeldung
        '
        Me.statusMeldung.AutoSize = True
        Me.statusMeldung.Location = New System.Drawing.Point(17, 422)
        Me.statusMeldung.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.statusMeldung.Name = "statusMeldung"
        Me.statusMeldung.Size = New System.Drawing.Size(178, 17)
        Me.statusMeldung.TabIndex = 17
        Me.statusMeldung.Text = "ok, Abhängigkeiten erstellt!"
        '
        'frmDependencies
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(120.0!, 120.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.ClientSize = New System.Drawing.Size(791, 453)
        Me.Controls.Add(Me.statusMeldung)
        Me.Controls.Add(Me.description)
        Me.Controls.Add(Me.deleteFromProjects)
        Me.Controls.Add(Me.copyFromProjects)
        Me.Controls.Add(Me.moveFromProjects)
        Me.Controls.Add(Me.deleteFromDependent)
        Me.Controls.Add(Me.copyFromDependent)
        Me.Controls.Add(Me.moveFromDependent)
        Me.Controls.Add(Me.AbbruchButton)
        Me.Controls.Add(Me.OKButton)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.degree)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ProjectList)
        Me.Controls.Add(Me.dependentProjectList)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Name = "frmDependencies"
        Me.Text = "Abhängigkeiten zwischen Projekten definieren"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents degree As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents OKButton As System.Windows.Forms.Button
    Friend WithEvents AbbruchButton As System.Windows.Forms.Button
    Public WithEvents dependentProjectList As System.Windows.Forms.ListBox
    Public WithEvents ProjectList As System.Windows.Forms.ListBox
    Friend WithEvents moveFromDependent As System.Windows.Forms.Button
    Friend WithEvents copyFromDependent As System.Windows.Forms.Button
    Friend WithEvents deleteFromDependent As System.Windows.Forms.Button
    Friend WithEvents deleteFromProjects As System.Windows.Forms.Button
    Friend WithEvents copyFromProjects As System.Windows.Forms.Button
    Friend WithEvents moveFromProjects As System.Windows.Forms.Button
    Friend WithEvents description As System.Windows.Forms.TextBox
    Friend WithEvents statusMeldung As System.Windows.Forms.Label
End Class
