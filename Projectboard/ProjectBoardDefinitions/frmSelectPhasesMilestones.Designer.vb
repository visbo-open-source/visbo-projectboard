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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSelectPhasesMilestones))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.TreeViewProjects = New System.Windows.Forms.TreeView()
        Me.zeitLabel = New System.Windows.Forms.Label()
        Me.vonDate = New System.Windows.Forms.DateTimePicker()
        Me.bisDate = New System.Windows.Forms.DateTimePicker()
        Me.einstellungen = New System.Windows.Forms.LinkLabel()
        Me.OK_Button = New System.Windows.Forms.Button()
        Me.collapseTree = New System.Windows.Forms.PictureBox()
        Me.expandTree = New System.Windows.Forms.PictureBox()
        Me.resetSelections = New System.Windows.Forms.PictureBox()
        Me.SelectionSet = New System.Windows.Forms.PictureBox()
        Me.rdbProjStruktProj = New System.Windows.Forms.RadioButton()
        Me.rdbProjStruktTyp = New System.Windows.Forms.RadioButton()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Panel1.SuspendLayout()
        CType(Me.collapseTree, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.expandTree, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.resetSelections, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SelectionSet, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.Controls.Add(Me.TreeViewProjects)
        Me.Panel1.Controls.Add(Me.zeitLabel)
        Me.Panel1.Controls.Add(Me.vonDate)
        Me.Panel1.Controls.Add(Me.bisDate)
        Me.Panel1.Controls.Add(Me.einstellungen)
        Me.Panel1.Controls.Add(Me.OK_Button)
        Me.Panel1.Controls.Add(Me.collapseTree)
        Me.Panel1.Controls.Add(Me.expandTree)
        Me.Panel1.Controls.Add(Me.resetSelections)
        Me.Panel1.Controls.Add(Me.SelectionSet)
        Me.Panel1.Controls.Add(Me.rdbProjStruktProj)
        Me.Panel1.Controls.Add(Me.rdbProjStruktTyp)
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Location = New System.Drawing.Point(0, 2)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(516, 427)
        Me.Panel1.TabIndex = 102
        '
        'TreeViewProjects
        '
        Me.TreeViewProjects.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TreeViewProjects.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TreeViewProjects.Location = New System.Drawing.Point(8, 39)
        Me.TreeViewProjects.Name = "TreeViewProjects"
        Me.TreeViewProjects.Size = New System.Drawing.Size(499, 318)
        Me.TreeViewProjects.TabIndex = 108
        '
        'zeitLabel
        '
        Me.zeitLabel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.zeitLabel.AutoSize = True
        Me.zeitLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte), True)
        Me.zeitLabel.Location = New System.Drawing.Point(202, 369)
        Me.zeitLabel.Name = "zeitLabel"
        Me.zeitLabel.Size = New System.Drawing.Size(63, 16)
        Me.zeitLabel.TabIndex = 107
        Me.zeitLabel.Text = "Zeitraum:"
        Me.zeitLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'vonDate
        '
        Me.vonDate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.vonDate.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.vonDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.vonDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.vonDate.Location = New System.Drawing.Point(269, 366)
        Me.vonDate.Name = "vonDate"
        Me.vonDate.Size = New System.Drawing.Size(108, 22)
        Me.vonDate.TabIndex = 106
        '
        'bisDate
        '
        Me.bisDate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.bisDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bisDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.bisDate.Location = New System.Drawing.Point(398, 366)
        Me.bisDate.Name = "bisDate"
        Me.bisDate.Size = New System.Drawing.Size(107, 22)
        Me.bisDate.TabIndex = 105
        '
        'einstellungen
        '
        Me.einstellungen.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.einstellungen.AutoSize = True
        Me.einstellungen.Location = New System.Drawing.Point(398, 402)
        Me.einstellungen.Name = "einstellungen"
        Me.einstellungen.Size = New System.Drawing.Size(70, 13)
        Me.einstellungen.TabIndex = 104
        Me.einstellungen.TabStop = True
        Me.einstellungen.Text = "Einstellungen"
        '
        'OK_Button
        '
        Me.OK_Button.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.OK_Button.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.OK_Button.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OK_Button.Location = New System.Drawing.Point(95, 397)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(265, 23)
        Me.OK_Button.TabIndex = 103
        Me.OK_Button.Text = "Auswahl bestätigen"
        Me.OK_Button.UseVisualStyleBackColor = True
        '
        'collapseTree
        '
        Me.collapseTree.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.collapseTree.BackColor = System.Drawing.SystemColors.Control
        Me.collapseTree.Image = CType(resources.GetObject("collapseTree.Image"), System.Drawing.Image)
        Me.collapseTree.Location = New System.Drawing.Point(53, 372)
        Me.collapseTree.Name = "collapseTree"
        Me.collapseTree.Size = New System.Drawing.Size(16, 16)
        Me.collapseTree.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.collapseTree.TabIndex = 102
        Me.collapseTree.TabStop = False
        '
        'expandTree
        '
        Me.expandTree.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.expandTree.BackColor = System.Drawing.SystemColors.Control
        Me.expandTree.Image = CType(resources.GetObject("expandTree.Image"), System.Drawing.Image)
        Me.expandTree.Location = New System.Drawing.Point(75, 372)
        Me.expandTree.Name = "expandTree"
        Me.expandTree.Size = New System.Drawing.Size(16, 16)
        Me.expandTree.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.expandTree.TabIndex = 100
        Me.expandTree.TabStop = False
        '
        'resetSelections
        '
        Me.resetSelections.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.resetSelections.BackColor = System.Drawing.SystemColors.Control
        Me.resetSelections.Image = CType(resources.GetObject("resetSelections.Image"), System.Drawing.Image)
        Me.resetSelections.InitialImage = Nothing
        Me.resetSelections.Location = New System.Drawing.Point(31, 372)
        Me.resetSelections.Name = "resetSelections"
        Me.resetSelections.Size = New System.Drawing.Size(16, 16)
        Me.resetSelections.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.resetSelections.TabIndex = 101
        Me.resetSelections.TabStop = False
        '
        'SelectionSet
        '
        Me.SelectionSet.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.SelectionSet.BackColor = System.Drawing.SystemColors.Control
        Me.SelectionSet.ErrorImage = CType(resources.GetObject("SelectionSet.ErrorImage"), System.Drawing.Image)
        Me.SelectionSet.Image = CType(resources.GetObject("SelectionSet.Image"), System.Drawing.Image)
        Me.SelectionSet.InitialImage = Nothing
        Me.SelectionSet.Location = New System.Drawing.Point(9, 372)
        Me.SelectionSet.Name = "SelectionSet"
        Me.SelectionSet.Size = New System.Drawing.Size(16, 16)
        Me.SelectionSet.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.SelectionSet.TabIndex = 99
        Me.SelectionSet.TabStop = False
        '
        'rdbProjStruktProj
        '
        Me.rdbProjStruktProj.AutoSize = True
        Me.rdbProjStruktProj.Checked = True
        Me.rdbProjStruktProj.Location = New System.Drawing.Point(165, 16)
        Me.rdbProjStruktProj.Name = "rdbProjStruktProj"
        Me.rdbProjStruktProj.Size = New System.Drawing.Size(140, 17)
        Me.rdbProjStruktProj.TabIndex = 67
        Me.rdbProjStruktProj.TabStop = True
        Me.rdbProjStruktProj.Text = "Projekt-Struktur (Projekt)"
        Me.rdbProjStruktProj.UseVisualStyleBackColor = True
        '
        'rdbProjStruktTyp
        '
        Me.rdbProjStruktTyp.AutoSize = True
        Me.rdbProjStruktTyp.Location = New System.Drawing.Point(9, 16)
        Me.rdbProjStruktTyp.Name = "rdbProjStruktTyp"
        Me.rdbProjStruktTyp.Size = New System.Drawing.Size(125, 17)
        Me.rdbProjStruktTyp.TabIndex = 68
        Me.rdbProjStruktTyp.Text = "Projekt-Struktur (Typ)"
        Me.rdbProjStruktTyp.UseVisualStyleBackColor = True
        '
        'Panel3
        '
        Me.Panel3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel3.Location = New System.Drawing.Point(9, 13)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(497, 25)
        Me.Panel3.TabIndex = 98
        '
        'frmSelectPhasesMilestones
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(516, 427)
        Me.Controls.Add(Me.Panel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmSelectPhasesMilestones"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Auswahl von Projekten, Phasen und Meilensteinen"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.collapseTree, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.expandTree, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.resetSelections, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SelectionSet, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As Windows.Forms.Panel
    Friend WithEvents TreeViewProjects As Windows.Forms.TreeView
    Friend WithEvents zeitLabel As Windows.Forms.Label
    Public WithEvents vonDate As Windows.Forms.DateTimePicker
    Public WithEvents bisDate As Windows.Forms.DateTimePicker
    Friend WithEvents einstellungen As Windows.Forms.LinkLabel
    Friend WithEvents OK_Button As Windows.Forms.Button
    Friend WithEvents collapseTree As Windows.Forms.PictureBox
    Friend WithEvents expandTree As Windows.Forms.PictureBox
    Friend WithEvents resetSelections As Windows.Forms.PictureBox
    Friend WithEvents SelectionSet As Windows.Forms.PictureBox
    Friend WithEvents rdbProjStruktProj As Windows.Forms.RadioButton
    Friend WithEvents rdbProjStruktTyp As Windows.Forms.RadioButton
    Friend WithEvents Panel3 As Windows.Forms.Panel
End Class
