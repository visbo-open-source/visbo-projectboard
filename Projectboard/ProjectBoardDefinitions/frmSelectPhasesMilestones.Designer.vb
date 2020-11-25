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
        Me.TreeViewProjects = New System.Windows.Forms.TreeView()
        Me.Ok_Button = New System.Windows.Forms.Button()
        Me.SelectionSet = New System.Windows.Forms.PictureBox()
        Me.resetSelections = New System.Windows.Forms.PictureBox()
        Me.collapseTree = New System.Windows.Forms.PictureBox()
        Me.expandTree = New System.Windows.Forms.PictureBox()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.rdbProjStruktTyp = New System.Windows.Forms.RadioButton()
        Me.rdbProjStruktProj = New System.Windows.Forms.RadioButton()
        Me.bisDate = New System.Windows.Forms.DateTimePicker()
        Me.vonDate = New System.Windows.Forms.DateTimePicker()
        Me.zeitLabel = New System.Windows.Forms.Label()
        Me.einstellungen = New System.Windows.Forms.LinkLabel()
        CType(Me.SelectionSet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.resetSelections, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.collapseTree, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.expandTree, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'TreeViewProjects
        '
        Me.TreeViewProjects.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TreeViewProjects.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TreeViewProjects.Location = New System.Drawing.Point(9, 35)
        Me.TreeViewProjects.Name = "TreeViewProjects"
        Me.TreeViewProjects.Size = New System.Drawing.Size(471, 261)
        Me.TreeViewProjects.TabIndex = 0
        '
        'Ok_Button
        '
        Me.Ok_Button.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Ok_Button.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Ok_Button.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Ok_Button.Location = New System.Drawing.Point(165, 334)
        Me.Ok_Button.Name = "Ok_Button"
        Me.Ok_Button.Size = New System.Drawing.Size(157, 23)
        Me.Ok_Button.TabIndex = 4
        Me.Ok_Button.Text = "Auswahl bestätigen"
        Me.Ok_Button.UseVisualStyleBackColor = True
        '
        'SelectionSet
        '
        Me.SelectionSet.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.SelectionSet.BackColor = System.Drawing.SystemColors.Control
        Me.SelectionSet.ErrorImage = CType(resources.GetObject("SelectionSet.ErrorImage"), System.Drawing.Image)
        Me.SelectionSet.Image = CType(resources.GetObject("SelectionSet.Image"), System.Drawing.Image)
        Me.SelectionSet.InitialImage = Nothing
        Me.SelectionSet.Location = New System.Drawing.Point(9, 303)
        Me.SelectionSet.Name = "SelectionSet"
        Me.SelectionSet.Size = New System.Drawing.Size(16, 16)
        Me.SelectionSet.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.SelectionSet.TabIndex = 90
        Me.SelectionSet.TabStop = False
        '
        'resetSelections
        '
        Me.resetSelections.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.resetSelections.BackColor = System.Drawing.SystemColors.Control
        Me.resetSelections.Image = CType(resources.GetObject("resetSelections.Image"), System.Drawing.Image)
        Me.resetSelections.InitialImage = Nothing
        Me.resetSelections.Location = New System.Drawing.Point(32, 303)
        Me.resetSelections.Name = "resetSelections"
        Me.resetSelections.Size = New System.Drawing.Size(16, 16)
        Me.resetSelections.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.resetSelections.TabIndex = 95
        Me.resetSelections.TabStop = False
        '
        'collapseTree
        '
        Me.collapseTree.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.collapseTree.BackColor = System.Drawing.SystemColors.Control
        Me.collapseTree.Image = CType(resources.GetObject("collapseTree.Image"), System.Drawing.Image)
        Me.collapseTree.Location = New System.Drawing.Point(55, 303)
        Me.collapseTree.Name = "collapseTree"
        Me.collapseTree.Size = New System.Drawing.Size(16, 16)
        Me.collapseTree.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.collapseTree.TabIndex = 96
        Me.collapseTree.TabStop = False
        '
        'expandTree
        '
        Me.expandTree.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.expandTree.BackColor = System.Drawing.SystemColors.Control
        Me.expandTree.Image = CType(resources.GetObject("expandTree.Image"), System.Drawing.Image)
        Me.expandTree.Location = New System.Drawing.Point(78, 303)
        Me.expandTree.Name = "expandTree"
        Me.expandTree.Size = New System.Drawing.Size(16, 16)
        Me.expandTree.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.expandTree.TabIndex = 95
        Me.expandTree.TabStop = False
        '
        'Panel3
        '
        Me.Panel3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel3.Controls.Add(Me.rdbProjStruktTyp)
        Me.Panel3.Controls.Add(Me.rdbProjStruktProj)
        Me.Panel3.Location = New System.Drawing.Point(4, 4)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(442, 25)
        Me.Panel3.TabIndex = 97
        '
        'rdbProjStruktTyp
        '
        Me.rdbProjStruktTyp.AutoSize = True
        Me.rdbProjStruktTyp.Location = New System.Drawing.Point(7, 3)
        Me.rdbProjStruktTyp.Name = "rdbProjStruktTyp"
        Me.rdbProjStruktTyp.Size = New System.Drawing.Size(125, 17)
        Me.rdbProjStruktTyp.TabIndex = 68
        Me.rdbProjStruktTyp.Text = "Projekt-Struktur (Typ)"
        Me.rdbProjStruktTyp.UseVisualStyleBackColor = True
        '
        'rdbProjStruktProj
        '
        Me.rdbProjStruktProj.AutoSize = True
        Me.rdbProjStruktProj.Checked = True
        Me.rdbProjStruktProj.Location = New System.Drawing.Point(139, 3)
        Me.rdbProjStruktProj.Name = "rdbProjStruktProj"
        Me.rdbProjStruktProj.Size = New System.Drawing.Size(140, 17)
        Me.rdbProjStruktProj.TabIndex = 67
        Me.rdbProjStruktProj.TabStop = True
        Me.rdbProjStruktProj.Text = "Projekt-Struktur (Projekt)"
        Me.rdbProjStruktProj.UseVisualStyleBackColor = True
        '
        'bisDate
        '
        Me.bisDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bisDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.bisDate.Location = New System.Drawing.Point(331, 301)
        Me.bisDate.Name = "bisDate"
        Me.bisDate.Size = New System.Drawing.Size(107, 22)
        Me.bisDate.TabIndex = 99
        '
        'vonDate
        '
        Me.vonDate.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.vonDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.vonDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.vonDate.Location = New System.Drawing.Point(201, 301)
        Me.vonDate.Name = "vonDate"
        Me.vonDate.Size = New System.Drawing.Size(108, 22)
        Me.vonDate.TabIndex = 98
        '
        'zeitLabel
        '
        Me.zeitLabel.AutoSize = True
        Me.zeitLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte), True)
        Me.zeitLabel.Location = New System.Drawing.Point(116, 303)
        Me.zeitLabel.Name = "zeitLabel"
        Me.zeitLabel.Size = New System.Drawing.Size(63, 16)
        Me.zeitLabel.TabIndex = 100
        Me.zeitLabel.Text = "Zeitraum:"
        Me.zeitLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'einstellungen
        '
        Me.einstellungen.AutoSize = True
        Me.einstellungen.Location = New System.Drawing.Point(348, 339)
        Me.einstellungen.Name = "einstellungen"
        Me.einstellungen.Size = New System.Drawing.Size(70, 13)
        Me.einstellungen.TabIndex = 101
        Me.einstellungen.TabStop = True
        Me.einstellungen.Text = "Einstellungen"
        '
        'frmSelectPhasesMilestones
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(487, 369)
        Me.Controls.Add(Me.einstellungen)
        Me.Controls.Add(Me.zeitLabel)
        Me.Controls.Add(Me.bisDate)
        Me.Controls.Add(Me.vonDate)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.collapseTree)
        Me.Controls.Add(Me.expandTree)
        Me.Controls.Add(Me.resetSelections)
        Me.Controls.Add(Me.SelectionSet)
        Me.Controls.Add(Me.Ok_Button)
        Me.Controls.Add(Me.TreeViewProjects)
        Me.Name = "frmSelectPhasesMilestones"
        Me.Text = "Auswahl von Projekten, Phasen und Meilensteinen"
        CType(Me.SelectionSet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.resetSelections, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.collapseTree, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.expandTree, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents TreeViewProjects As Windows.Forms.TreeView
    Friend WithEvents Ok_Button As Windows.Forms.Button
    Friend WithEvents SelectionSet As Windows.Forms.PictureBox
    Friend WithEvents resetSelections As Windows.Forms.PictureBox
    Friend WithEvents collapseTree As Windows.Forms.PictureBox
    Friend WithEvents expandTree As Windows.Forms.PictureBox
    Friend WithEvents Panel3 As Windows.Forms.Panel
    Friend WithEvents rdbProjStruktTyp As Windows.Forms.RadioButton
    Friend WithEvents rdbProjStruktProj As Windows.Forms.RadioButton
    Friend WithEvents zeitLabel As Windows.Forms.Label
    Public WithEvents bisDate As Windows.Forms.DateTimePicker
    Public WithEvents vonDate As Windows.Forms.DateTimePicker
    Friend WithEvents einstellungen As Windows.Forms.LinkLabel
End Class
