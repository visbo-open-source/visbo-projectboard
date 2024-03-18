<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAllocateRessources
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAllocateRessources))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.lblContext = New System.Windows.Forms.Label()
        Me.lblUnit = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.explanationLbl = New System.Windows.Forms.Label()
        Me.CancelBtn = New System.Windows.Forms.Button()
        Me.okBtn = New System.Windows.Forms.Button()
        Me.candidatesTable = New System.Windows.Forms.DataGridView()
        Me.colLblPerson = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ColLblFreeCapacity = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ColLblIsExtern = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colLblAmount = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.lblSum = New System.Windows.Forms.Label()
        Me.lblOrgaUnitSkill = New System.Windows.Forms.Label()
        Me.lblAllowOverloads = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        CType(Me.candidatesTable, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.Controls.Add(Me.lblAllowOverloads)
        Me.Panel1.Controls.Add(Me.lblContext)
        Me.Panel1.Controls.Add(Me.lblUnit)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.explanationLbl)
        Me.Panel1.Controls.Add(Me.CancelBtn)
        Me.Panel1.Controls.Add(Me.okBtn)
        Me.Panel1.Controls.Add(Me.candidatesTable)
        Me.Panel1.Controls.Add(Me.lblSum)
        Me.Panel1.Controls.Add(Me.lblOrgaUnitSkill)
        Me.Panel1.Location = New System.Drawing.Point(-1, 1)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(562, 352)
        Me.Panel1.TabIndex = 0
        '
        'lblContext
        '
        Me.lblContext.AutoSize = True
        Me.lblContext.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblContext.Location = New System.Drawing.Point(23, 61)
        Me.lblContext.Name = "lblContext"
        Me.lblContext.Size = New System.Drawing.Size(175, 16)
        Me.lblContext.TabIndex = 9
        Me.lblContext.Text = "considering loaded projects"
        Me.lblContext.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblUnit
        '
        Me.lblUnit.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblUnit.AutoSize = True
        Me.lblUnit.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUnit.Location = New System.Drawing.Point(507, 12)
        Me.lblUnit.Name = "lblUnit"
        Me.lblUnit.Size = New System.Drawing.Size(31, 20)
        Me.lblUnit.TabIndex = 8
        Me.lblUnit.Text = "PD"
        Me.lblUnit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(21, 286)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(233, 13)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Names with color are already in the project team"
        '
        'explanationLbl
        '
        Me.explanationLbl.AutoSize = True
        Me.explanationLbl.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.explanationLbl.Location = New System.Drawing.Point(23, 42)
        Me.explanationLbl.Name = "explanationLbl"
        Me.explanationLbl.Size = New System.Drawing.Size(285, 16)
        Me.explanationLbl.TabIndex = 6
        Me.explanationLbl.Text = "Candidates with free Capacity [PD] in timespan"
        Me.explanationLbl.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'CancelBtn
        '
        Me.CancelBtn.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.CancelBtn.Location = New System.Drawing.Point(332, 321)
        Me.CancelBtn.Name = "CancelBtn"
        Me.CancelBtn.Size = New System.Drawing.Size(75, 23)
        Me.CancelBtn.TabIndex = 5
        Me.CancelBtn.Text = "Cancel"
        Me.CancelBtn.UseVisualStyleBackColor = True
        '
        'okBtn
        '
        Me.okBtn.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.okBtn.Location = New System.Drawing.Point(110, 321)
        Me.okBtn.Name = "okBtn"
        Me.okBtn.Size = New System.Drawing.Size(75, 23)
        Me.okBtn.TabIndex = 4
        Me.okBtn.Text = "OK"
        Me.okBtn.UseVisualStyleBackColor = True
        '
        'candidatesTable
        '
        Me.candidatesTable.AllowUserToAddRows = False
        Me.candidatesTable.AllowUserToDeleteRows = False
        Me.candidatesTable.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.candidatesTable.BackgroundColor = System.Drawing.SystemColors.Control
        Me.candidatesTable.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.candidatesTable.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None
        Me.candidatesTable.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.candidatesTable.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colLblPerson, Me.ColLblFreeCapacity, Me.ColLblIsExtern, Me.colLblAmount})
        Me.candidatesTable.Location = New System.Drawing.Point(24, 83)
        Me.candidatesTable.Name = "candidatesTable"
        Me.candidatesTable.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.candidatesTable.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.candidatesTable.Size = New System.Drawing.Size(516, 200)
        Me.candidatesTable.TabIndex = 3
        '
        'colLblPerson
        '
        Me.colLblPerson.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.colLblPerson.HeaderText = "Name"
        Me.colLblPerson.Name = "colLblPerson"
        Me.colLblPerson.ReadOnly = True
        Me.colLblPerson.Width = 220
        '
        'ColLblFreeCapacity
        '
        Me.ColLblFreeCapacity.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.ColLblFreeCapacity.HeaderText = "Free Capacity in PD [1 PD = 8 hrs]"
        Me.ColLblFreeCapacity.Name = "ColLblFreeCapacity"
        Me.ColLblFreeCapacity.ReadOnly = True
        Me.ColLblFreeCapacity.ToolTipText = "shows the amount of free capacity in person days "
        Me.ColLblFreeCapacity.Width = 120
        '
        'ColLblIsExtern
        '
        Me.ColLblIsExtern.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.ColLblIsExtern.HeaderText = "is Extern"
        Me.ColLblIsExtern.Name = "ColLblIsExtern"
        Me.ColLblIsExtern.ReadOnly = True
        Me.ColLblIsExtern.Width = 40
        '
        'colLblAmount
        '
        Me.colLblAmount.HeaderText = "will do how much PD?"
        Me.colLblAmount.Name = "colLblAmount"
        Me.colLblAmount.ToolTipText = "will do how many person days "
        Me.colLblAmount.Width = 90
        '
        'lblSum
        '
        Me.lblSum.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblSum.AutoSize = True
        Me.lblSum.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSum.Location = New System.Drawing.Point(426, 12)
        Me.lblSum.Name = "lblSum"
        Me.lblSum.Size = New System.Drawing.Size(83, 20)
        Me.lblSum.TabIndex = 2
        Me.lblSum.Text = "<Amount>"
        Me.lblSum.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblOrgaUnitSkill
        '
        Me.lblOrgaUnitSkill.AutoSize = True
        Me.lblOrgaUnitSkill.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOrgaUnitSkill.Location = New System.Drawing.Point(20, 12)
        Me.lblOrgaUnitSkill.Name = "lblOrgaUnitSkill"
        Me.lblOrgaUnitSkill.Size = New System.Drawing.Size(105, 20)
        Me.lblOrgaUnitSkill.TabIndex = 0
        Me.lblOrgaUnitSkill.Text = "<Name, Skill>"
        Me.lblOrgaUnitSkill.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblAllowOverloads
        '
        Me.lblAllowOverloads.AutoSize = True
        Me.lblAllowOverloads.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAllowOverloads.ForeColor = System.Drawing.Color.Maroon
        Me.lblAllowOverloads.Location = New System.Drawing.Point(432, 286)
        Me.lblAllowOverloads.Name = "lblAllowOverloads"
        Me.lblAllowOverloads.Size = New System.Drawing.Size(108, 13)
        Me.lblAllowOverloads.TabIndex = 10
        Me.lblAllowOverloads.Text = "Overload is allowed ! "
        '
        'frmAllocateRessources
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(558, 351)
        Me.Controls.Add(Me.Panel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmAllocateRessources"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Allocate Ressources"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.candidatesTable, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Panel1 As Windows.Forms.Panel
    Friend WithEvents candidatesTable As Windows.Forms.DataGridView
    Friend WithEvents lblSum As Windows.Forms.Label
    Friend WithEvents lblOrgaUnitSkill As Windows.Forms.Label
    Friend WithEvents explanationLbl As Windows.Forms.Label
    Friend WithEvents CancelBtn As Windows.Forms.Button
    Friend WithEvents okBtn As Windows.Forms.Button
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents colLblPerson As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLblFreeCapacity As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ColLblIsExtern As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colLblAmount As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents lblUnit As Windows.Forms.Label
    Friend WithEvents lblContext As Windows.Forms.Label
    Friend WithEvents lblAllowOverloads As Windows.Forms.Label
End Class
