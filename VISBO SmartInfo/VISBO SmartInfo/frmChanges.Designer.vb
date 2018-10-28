<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmChanges
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmChanges))
        Me.changeListTable = New System.Windows.Forms.DataGridView()
        Me.colPname = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colElemName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ts1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ts2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colDiff = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.changeListTable, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'changeListTable
        '
        Me.changeListTable.AllowUserToAddRows = False
        Me.changeListTable.AllowUserToDeleteRows = False
        Me.changeListTable.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.changeListTable.BackgroundColor = System.Drawing.SystemColors.ControlLightLight
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.changeListTable.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.changeListTable.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.changeListTable.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colPname, Me.colElemName, Me.ts1, Me.ts2, Me.colDiff})
        Me.changeListTable.Dock = System.Windows.Forms.DockStyle.Fill
        Me.changeListTable.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.changeListTable.Location = New System.Drawing.Point(0, 0)
        Me.changeListTable.Name = "changeListTable"
        Me.changeListTable.ReadOnly = True
        Me.changeListTable.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.changeListTable.ShowCellErrors = False
        Me.changeListTable.ShowCellToolTips = False
        Me.changeListTable.ShowEditingIcon = False
        Me.changeListTable.ShowRowErrors = False
        Me.changeListTable.Size = New System.Drawing.Size(799, 86)
        Me.changeListTable.TabIndex = 0
        '
        'colPname
        '
        Me.colPname.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.colPname.DefaultCellStyle = DataGridViewCellStyle2
        Me.colPname.HeaderText = "Project"
        Me.colPname.MinimumWidth = 20
        Me.colPname.Name = "colPname"
        Me.colPname.ReadOnly = True
        Me.colPname.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Programmatic
        '
        'colElemName
        '
        Me.colElemName.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.colElemName.DefaultCellStyle = DataGridViewCellStyle3
        Me.colElemName.HeaderText = "Element"
        Me.colElemName.MinimumWidth = 20
        Me.colElemName.Name = "colElemName"
        Me.colElemName.ReadOnly = True
        Me.colElemName.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ts1
        '
        Me.ts1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ts1.DefaultCellStyle = DataGridViewCellStyle4
        Me.ts1.HeaderText = "Version 23.06.17"
        Me.ts1.MinimumWidth = 20
        Me.ts1.Name = "ts1"
        Me.ts1.ReadOnly = True
        Me.ts1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'ts2
        '
        Me.ts2.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ts2.DefaultCellStyle = DataGridViewCellStyle5
        Me.ts2.HeaderText = "Version 01.07.18"
        Me.ts2.MinimumWidth = 20
        Me.ts2.Name = "ts2"
        Me.ts2.ReadOnly = True
        Me.ts2.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'colDiff
        '
        Me.colDiff.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.colDiff.DefaultCellStyle = DataGridViewCellStyle6
        Me.colDiff.HeaderText = "Difference End-Date"
        Me.colDiff.MinimumWidth = 30
        Me.colDiff.Name = "colDiff"
        Me.colDiff.ReadOnly = True
        '
        'frmChanges
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(799, 86)
        Me.Controls.Add(Me.changeListTable)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(24, 922)
        Me.MaximizeBox = False
        Me.MinimumSize = New System.Drawing.Size(100, 38)
        Me.Name = "frmChanges"
        Me.Text = "Changes"
        Me.TopMost = True
        CType(Me.changeListTable, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents changeListTable As System.Windows.Forms.DataGridView
    Friend WithEvents colPname As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colElemName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ts1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ts2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colDiff As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
