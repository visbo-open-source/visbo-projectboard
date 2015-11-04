<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCreateRolloutProject
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
        Me.components = New System.ComponentModel.Container()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.endDateMigration = New System.Windows.Forms.TextBox()
        Me.nrVeryComplexObjects = New System.Windows.Forms.TextBox()
        Me.nrComplexObjects = New System.Windows.Forms.TextBox()
        Me.nrMediumObjects = New System.Windows.Forms.TextBox()
        Me.nrSimpleObjects = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.endDateTraining = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.reqUsers = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.AbbrButton = New System.Windows.Forms.Button()
        Me.bestFit = New System.Windows.Forms.CheckBox()
        Me.BindingSource1 = New System.Windows.Forms.BindingSource(Me.components)
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.BindingSource1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.endDateMigration)
        Me.GroupBox1.Controls.Add(Me.nrVeryComplexObjects)
        Me.GroupBox1.Controls.Add(Me.nrComplexObjects)
        Me.GroupBox1.Controls.Add(Me.nrMediumObjects)
        Me.GroupBox1.Controls.Add(Me.nrSimpleObjects)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(36, 155)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox1.Size = New System.Drawing.Size(511, 224)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Migration Project"
        '
        'endDateMigration
        '
        Me.endDateMigration.Location = New System.Drawing.Point(309, 185)
        Me.endDateMigration.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.endDateMigration.Name = "endDateMigration"
        Me.endDateMigration.Size = New System.Drawing.Size(132, 22)
        Me.endDateMigration.TabIndex = 10
        '
        'nrVeryComplexObjects
        '
        Me.nrVeryComplexObjects.Location = New System.Drawing.Point(309, 151)
        Me.nrVeryComplexObjects.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.nrVeryComplexObjects.Name = "nrVeryComplexObjects"
        Me.nrVeryComplexObjects.Size = New System.Drawing.Size(132, 22)
        Me.nrVeryComplexObjects.TabIndex = 9
        '
        'nrComplexObjects
        '
        Me.nrComplexObjects.Location = New System.Drawing.Point(309, 118)
        Me.nrComplexObjects.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.nrComplexObjects.Name = "nrComplexObjects"
        Me.nrComplexObjects.Size = New System.Drawing.Size(132, 22)
        Me.nrComplexObjects.TabIndex = 8
        '
        'nrMediumObjects
        '
        Me.nrMediumObjects.Location = New System.Drawing.Point(309, 86)
        Me.nrMediumObjects.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.nrMediumObjects.Name = "nrMediumObjects"
        Me.nrMediumObjects.Size = New System.Drawing.Size(132, 22)
        Me.nrMediumObjects.TabIndex = 7
        '
        'nrSimpleObjects
        '
        Me.nrSimpleObjects.Location = New System.Drawing.Point(309, 54)
        Me.nrSimpleObjects.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.nrSimpleObjects.Name = "nrSimpleObjects"
        Me.nrSimpleObjects.Size = New System.Drawing.Size(132, 22)
        Me.nrSimpleObjects.TabIndex = 6
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(152, 188)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(99, 17)
        Me.Label6.TabIndex = 5
        Me.Label6.Text = "latest Enddate"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(111, 156)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(139, 17)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "very complex objects"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(141, 123)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(108, 17)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "complex objects"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(145, 91)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(106, 17)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "medium objects"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(155, 59)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(97, 17)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "simple objects"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(41, 27)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(133, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "required number of "
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.endDateTraining)
        Me.GroupBox2.Controls.Add(Me.Label11)
        Me.GroupBox2.Controls.Add(Me.reqUsers)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Location = New System.Drawing.Point(36, 399)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox2.Size = New System.Drawing.Size(511, 96)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Training Project"
        '
        'endDateTraining
        '
        Me.endDateTraining.Location = New System.Drawing.Point(309, 59)
        Me.endDateTraining.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.endDateTraining.Name = "endDateTraining"
        Me.endDateTraining.Size = New System.Drawing.Size(132, 22)
        Me.endDateTraining.TabIndex = 3
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(152, 64)
        Me.Label11.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(99, 17)
        Me.Label11.TabIndex = 2
        Me.Label11.Text = "latest Enddate"
        '
        'reqUsers
        '
        Me.reqUsers.Location = New System.Drawing.Point(309, 27)
        Me.reqUsers.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.reqUsers.Name = "reqUsers"
        Me.reqUsers.Size = New System.Drawing.Size(132, 22)
        Me.reqUsers.TabIndex = 1
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(41, 32)
        Me.Label7.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(216, 17)
        Me.Label7.TabIndex = 0
        Me.Label7.Text = "required number of trained users"
        '
        'OKButton
        '
        Me.OKButton.Location = New System.Drawing.Point(81, 546)
        Me.OKButton.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(100, 28)
        Me.OKButton.TabIndex = 2
        Me.OKButton.Text = "OK"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'AbbrButton
        '
        Me.AbbrButton.Location = New System.Drawing.Point(377, 545)
        Me.AbbrButton.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.AbbrButton.Name = "AbbrButton"
        Me.AbbrButton.Size = New System.Drawing.Size(100, 28)
        Me.AbbrButton.TabIndex = 3
        Me.AbbrButton.Text = "Abbrechen"
        Me.AbbrButton.UseVisualStyleBackColor = True
        '
        'bestFit
        '
        Me.bestFit.AutoSize = True
        Me.bestFit.Location = New System.Drawing.Point(345, 502)
        Me.bestFit.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.bestFit.Name = "bestFit"
        Me.bestFit.Size = New System.Drawing.Size(76, 21)
        Me.bestFit.TabIndex = 4
        Me.bestFit.Text = "best Fit"
        Me.bestFit.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.TextBox2)
        Me.GroupBox3.Controls.Add(Me.TextBox1)
        Me.GroupBox3.Controls.Add(Me.Label9)
        Me.GroupBox3.Controls.Add(Me.Label8)
        Me.GroupBox3.Location = New System.Drawing.Point(36, 22)
        Me.GroupBox3.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox3.Size = New System.Drawing.Size(511, 112)
        Me.GroupBox3.TabIndex = 5
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "General Information"
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(180, 63)
        Me.TextBox2.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(261, 22)
        Me.TextBox2.TabIndex = 4
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(180, 32)
        Me.TextBox1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(261, 22)
        Me.TextBox1.TabIndex = 3
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(60, 68)
        Me.Label9.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(94, 17)
        Me.Label9.TabIndex = 1
        Me.Label9.Text = "Business Unit"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(108, 37)
        Me.Label8.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(45, 17)
        Me.Label8.TabIndex = 0
        Me.Label8.Text = "Name"
        '
        'frmCreateRolloutProject
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(120.0!, 120.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.ClientSize = New System.Drawing.Size(607, 597)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.bestFit)
        Me.Controls.Add(Me.AbbrButton)
        Me.Controls.Add(Me.OKButton)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Name = "frmCreateRolloutProject"
        Me.Text = "Create new Rollout Project"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.BindingSource1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents endDateMigration As System.Windows.Forms.TextBox
    Friend WithEvents nrVeryComplexObjects As System.Windows.Forms.TextBox
    Friend WithEvents nrComplexObjects As System.Windows.Forms.TextBox
    Friend WithEvents nrMediumObjects As System.Windows.Forms.TextBox
    Friend WithEvents nrSimpleObjects As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents endDateTraining As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents reqUsers As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents OKButton As System.Windows.Forms.Button
    Friend WithEvents AbbrButton As System.Windows.Forms.Button
    Friend WithEvents bestFit As System.Windows.Forms.CheckBox
    Friend WithEvents BindingSource1 As System.Windows.Forms.BindingSource
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
End Class
