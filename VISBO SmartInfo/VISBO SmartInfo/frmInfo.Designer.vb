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
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.fullBreadCrumb = New System.Windows.Forms.TextBox()
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
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.lblAmpeln = New System.Windows.Forms.Label()
        Me.shwRedLight = New System.Windows.Forms.CheckBox()
        Me.shwYellowLight = New System.Windows.Forms.CheckBox()
        Me.shwGreenLight = New System.Windows.Forms.CheckBox()
        Me.shwOhneLight = New System.Windows.Forms.CheckBox()
        Me.listboxNames = New System.Windows.Forms.ListBox()
        Me.filterText = New System.Windows.Forms.TextBox()
        Me.rdbBreadcrumb = New System.Windows.Forms.RadioButton()
        Me.rdbAbbrev = New System.Windows.Forms.RadioButton()
        Me.rdbOriginalName = New System.Windows.Forms.RadioButton()
        Me.rdbName = New System.Windows.Forms.RadioButton()
        Me.searchIcon = New System.Windows.Forms.Button()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(12, 18)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(542, 186)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.fullBreadCrumb)
        Me.TabPage1.Controls.Add(Me.positionDateButton)
        Me.TabPage1.Controls.Add(Me.deleteDate)
        Me.TabPage1.Controls.Add(Me.writeDate)
        Me.TabPage1.Controls.Add(Me.elemDate)
        Me.TabPage1.Controls.Add(Me.positionTextButton)
        Me.TabPage1.Controls.Add(Me.deleteText)
        Me.TabPage1.Controls.Add(Me.showOrginalName)
        Me.TabPage1.Controls.Add(Me.elemName)
        Me.TabPage1.Controls.Add(Me.showAbbrev)
        Me.TabPage1.Controls.Add(Me.writeText)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(534, 160)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Information"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'fullBreadCrumb
        '
        Me.fullBreadCrumb.BackColor = System.Drawing.SystemColors.Window
        Me.fullBreadCrumb.Enabled = False
        Me.fullBreadCrumb.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fullBreadCrumb.Location = New System.Drawing.Point(20, 74)
        Me.fullBreadCrumb.Name = "fullBreadCrumb"
        Me.fullBreadCrumb.ReadOnly = True
        Me.fullBreadCrumb.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal
        Me.fullBreadCrumb.Size = New System.Drawing.Size(401, 20)
        Me.fullBreadCrumb.TabIndex = 28
        Me.fullBreadCrumb.Visible = False
        '
        'positionDateButton
        '
        Me.positionDateButton.Image = Global.VISBO_SmartInfo.My.Resources.Resources.layout_south
        Me.positionDateButton.Location = New System.Drawing.Point(457, 100)
        Me.positionDateButton.Name = "positionDateButton"
        Me.positionDateButton.Size = New System.Drawing.Size(30, 26)
        Me.positionDateButton.TabIndex = 27
        Me.positionDateButton.UseVisualStyleBackColor = True
        '
        'deleteDate
        '
        Me.deleteDate.Image = Global.VISBO_SmartInfo.My.Resources.Resources.selection_delete
        Me.deleteDate.Location = New System.Drawing.Point(427, 100)
        Me.deleteDate.Name = "deleteDate"
        Me.deleteDate.Size = New System.Drawing.Size(30, 26)
        Me.deleteDate.TabIndex = 26
        Me.deleteDate.UseVisualStyleBackColor = True
        '
        'writeDate
        '
        Me.writeDate.Image = Global.VISBO_SmartInfo.My.Resources.Resources.pen_blue
        Me.writeDate.Location = New System.Drawing.Point(487, 100)
        Me.writeDate.Name = "writeDate"
        Me.writeDate.Size = New System.Drawing.Size(30, 26)
        Me.writeDate.TabIndex = 25
        Me.writeDate.UseVisualStyleBackColor = True
        '
        'elemDate
        '
        Me.elemDate.Enabled = False
        Me.elemDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.elemDate.Location = New System.Drawing.Point(20, 99)
        Me.elemDate.Name = "elemDate"
        Me.elemDate.Size = New System.Drawing.Size(401, 26)
        Me.elemDate.TabIndex = 24
        '
        'positionTextButton
        '
        Me.positionTextButton.Image = Global.VISBO_SmartInfo.My.Resources.Resources.layout_north
        Me.positionTextButton.Location = New System.Drawing.Point(457, 46)
        Me.positionTextButton.Name = "positionTextButton"
        Me.positionTextButton.Size = New System.Drawing.Size(30, 26)
        Me.positionTextButton.TabIndex = 23
        Me.positionTextButton.UseVisualStyleBackColor = True
        '
        'deleteText
        '
        Me.deleteText.Image = Global.VISBO_SmartInfo.My.Resources.Resources.selection_delete
        Me.deleteText.Location = New System.Drawing.Point(427, 46)
        Me.deleteText.Name = "deleteText"
        Me.deleteText.Size = New System.Drawing.Size(30, 26)
        Me.deleteText.TabIndex = 22
        Me.deleteText.UseVisualStyleBackColor = True
        '
        'showOrginalName
        '
        Me.showOrginalName.AutoSize = True
        Me.showOrginalName.Location = New System.Drawing.Point(121, 22)
        Me.showOrginalName.Name = "showOrginalName"
        Me.showOrginalName.Size = New System.Drawing.Size(92, 17)
        Me.showOrginalName.TabIndex = 21
        Me.showOrginalName.Text = "Original-Name"
        Me.showOrginalName.UseVisualStyleBackColor = True
        '
        'elemName
        '
        Me.elemName.Enabled = False
        Me.elemName.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.elemName.Location = New System.Drawing.Point(20, 45)
        Me.elemName.Name = "elemName"
        Me.elemName.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal
        Me.elemName.Size = New System.Drawing.Size(401, 26)
        Me.elemName.TabIndex = 18
        '
        'showAbbrev
        '
        Me.showAbbrev.AutoSize = True
        Me.showAbbrev.Location = New System.Drawing.Point(20, 22)
        Me.showAbbrev.Name = "showAbbrev"
        Me.showAbbrev.Size = New System.Drawing.Size(77, 17)
        Me.showAbbrev.TabIndex = 20
        Me.showAbbrev.Text = "Abkürzung"
        Me.showAbbrev.UseVisualStyleBackColor = True
        '
        'writeText
        '
        Me.writeText.Image = Global.VISBO_SmartInfo.My.Resources.Resources.pen_blue
        Me.writeText.Location = New System.Drawing.Point(487, 46)
        Me.writeText.Name = "writeText"
        Me.writeText.Size = New System.Drawing.Size(30, 26)
        Me.writeText.TabIndex = 19
        Me.writeText.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(534, 163)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Messen"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'lblAmpeln
        '
        Me.lblAmpeln.AutoSize = True
        Me.lblAmpeln.Location = New System.Drawing.Point(13, 223)
        Me.lblAmpeln.Name = "lblAmpeln"
        Me.lblAmpeln.Size = New System.Drawing.Size(79, 13)
        Me.lblAmpeln.TabIndex = 1
        Me.lblAmpeln.Text = "Ampeln zeigen:"
        '
        'shwRedLight
        '
        Me.shwRedLight.AutoSize = True
        Me.shwRedLight.BackColor = System.Drawing.Color.Firebrick
        Me.shwRedLight.Location = New System.Drawing.Point(194, 222)
        Me.shwRedLight.Name = "shwRedLight"
        Me.shwRedLight.Size = New System.Drawing.Size(15, 14)
        Me.shwRedLight.TabIndex = 36
        Me.shwRedLight.UseVisualStyleBackColor = False
        '
        'shwYellowLight
        '
        Me.shwYellowLight.AutoSize = True
        Me.shwYellowLight.BackColor = System.Drawing.Color.Yellow
        Me.shwYellowLight.Location = New System.Drawing.Point(163, 222)
        Me.shwYellowLight.Name = "shwYellowLight"
        Me.shwYellowLight.Size = New System.Drawing.Size(15, 14)
        Me.shwYellowLight.TabIndex = 35
        Me.shwYellowLight.UseVisualStyleBackColor = False
        '
        'shwGreenLight
        '
        Me.shwGreenLight.AutoSize = True
        Me.shwGreenLight.BackColor = System.Drawing.Color.LawnGreen
        Me.shwGreenLight.Location = New System.Drawing.Point(130, 222)
        Me.shwGreenLight.Name = "shwGreenLight"
        Me.shwGreenLight.Size = New System.Drawing.Size(15, 14)
        Me.shwGreenLight.TabIndex = 34
        Me.shwGreenLight.UseVisualStyleBackColor = False
        '
        'shwOhneLight
        '
        Me.shwOhneLight.AutoSize = True
        Me.shwOhneLight.Location = New System.Drawing.Point(101, 222)
        Me.shwOhneLight.Name = "shwOhneLight"
        Me.shwOhneLight.Size = New System.Drawing.Size(15, 14)
        Me.shwOhneLight.TabIndex = 33
        Me.shwOhneLight.UseVisualStyleBackColor = True
        '
        'listboxNames
        '
        Me.listboxNames.FormattingEnabled = True
        Me.listboxNames.HorizontalScrollbar = True
        Me.listboxNames.Location = New System.Drawing.Point(12, 307)
        Me.listboxNames.Name = "listboxNames"
        Me.listboxNames.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.listboxNames.Size = New System.Drawing.Size(538, 186)
        Me.listboxNames.Sorted = True
        Me.listboxNames.TabIndex = 38
        '
        'filterText
        '
        Me.filterText.Location = New System.Drawing.Point(12, 277)
        Me.filterText.Name = "filterText"
        Me.filterText.Size = New System.Drawing.Size(538, 20)
        Me.filterText.TabIndex = 37
        '
        'rdbBreadcrumb
        '
        Me.rdbBreadcrumb.AutoSize = True
        Me.rdbBreadcrumb.Location = New System.Drawing.Point(336, 258)
        Me.rdbBreadcrumb.Name = "rdbBreadcrumb"
        Me.rdbBreadcrumb.Size = New System.Drawing.Size(81, 17)
        Me.rdbBreadcrumb.TabIndex = 42
        Me.rdbBreadcrumb.Text = "voller Name"
        Me.rdbBreadcrumb.UseVisualStyleBackColor = True
        Me.rdbBreadcrumb.Visible = False
        '
        'rdbAbbrev
        '
        Me.rdbAbbrev.AutoSize = True
        Me.rdbAbbrev.Location = New System.Drawing.Point(227, 258)
        Me.rdbAbbrev.Name = "rdbAbbrev"
        Me.rdbAbbrev.Size = New System.Drawing.Size(76, 17)
        Me.rdbAbbrev.TabIndex = 41
        Me.rdbAbbrev.Text = "Abkürzung"
        Me.rdbAbbrev.UseVisualStyleBackColor = True
        Me.rdbAbbrev.Visible = False
        '
        'rdbOriginalName
        '
        Me.rdbOriginalName.AutoSize = True
        Me.rdbOriginalName.Location = New System.Drawing.Point(102, 258)
        Me.rdbOriginalName.Name = "rdbOriginalName"
        Me.rdbOriginalName.Size = New System.Drawing.Size(91, 17)
        Me.rdbOriginalName.TabIndex = 40
        Me.rdbOriginalName.Text = "Original Name"
        Me.rdbOriginalName.UseVisualStyleBackColor = True
        Me.rdbOriginalName.Visible = False
        '
        'rdbName
        '
        Me.rdbName.AutoSize = True
        Me.rdbName.Location = New System.Drawing.Point(15, 258)
        Me.rdbName.Name = "rdbName"
        Me.rdbName.Size = New System.Drawing.Size(53, 17)
        Me.rdbName.TabIndex = 39
        Me.rdbName.Text = "Name"
        Me.rdbName.UseVisualStyleBackColor = True
        Me.rdbName.Visible = False
        '
        'searchIcon
        '
        Me.searchIcon.Image = Global.VISBO_SmartInfo.My.Resources.Resources.view
        Me.searchIcon.Location = New System.Drawing.Point(470, 210)
        Me.searchIcon.Name = "searchIcon"
        Me.searchIcon.Size = New System.Drawing.Size(40, 40)
        Me.searchIcon.TabIndex = 29
        Me.searchIcon.UseVisualStyleBackColor = True
        '
        'frmInfo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(566, 511)
        Me.Controls.Add(Me.rdbBreadcrumb)
        Me.Controls.Add(Me.rdbAbbrev)
        Me.Controls.Add(Me.rdbOriginalName)
        Me.Controls.Add(Me.rdbName)
        Me.Controls.Add(Me.listboxNames)
        Me.Controls.Add(Me.filterText)
        Me.Controls.Add(Me.searchIcon)
        Me.Controls.Add(Me.shwRedLight)
        Me.Controls.Add(Me.shwYellowLight)
        Me.Controls.Add(Me.shwGreenLight)
        Me.Controls.Add(Me.shwOhneLight)
        Me.Controls.Add(Me.lblAmpeln)
        Me.Controls.Add(Me.TabControl1)
        Me.Name = "frmInfo"
        Me.Text = "VISBO Smart-Info"
        Me.TopMost = True
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents positionTextButton As System.Windows.Forms.Button
    Friend WithEvents deleteText As System.Windows.Forms.Button
    Friend WithEvents showOrginalName As System.Windows.Forms.CheckBox
    Friend WithEvents elemName As System.Windows.Forms.TextBox
    Friend WithEvents showAbbrev As System.Windows.Forms.CheckBox
    Friend WithEvents writeText As System.Windows.Forms.Button
    Friend WithEvents positionDateButton As System.Windows.Forms.Button
    Friend WithEvents deleteDate As System.Windows.Forms.Button
    Friend WithEvents writeDate As System.Windows.Forms.Button
    Friend WithEvents elemDate As System.Windows.Forms.TextBox
    Friend WithEvents fullBreadCrumb As System.Windows.Forms.TextBox
    Friend WithEvents lblAmpeln As System.Windows.Forms.Label
    Friend WithEvents shwRedLight As System.Windows.Forms.CheckBox
    Friend WithEvents shwYellowLight As System.Windows.Forms.CheckBox
    Friend WithEvents shwGreenLight As System.Windows.Forms.CheckBox
    Friend WithEvents shwOhneLight As System.Windows.Forms.CheckBox
    Friend WithEvents searchIcon As System.Windows.Forms.Button
    Friend WithEvents listboxNames As System.Windows.Forms.ListBox
    Friend WithEvents filterText As System.Windows.Forms.TextBox
    Friend WithEvents rdbBreadcrumb As System.Windows.Forms.RadioButton
    Friend WithEvents rdbAbbrev As System.Windows.Forms.RadioButton
    Friend WithEvents rdbOriginalName As System.Windows.Forms.RadioButton
    Friend WithEvents rdbName As System.Windows.Forms.RadioButton
End Class
