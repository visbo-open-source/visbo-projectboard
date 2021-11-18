<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmChooseCustomUserRole
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmChooseCustomUserRole))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.dgv_customUserRoles = New System.Windows.Forms.DataGridView()
        Me.userRole = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.specifics = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        CType(Me.dgv_customUserRoles, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.Controls.Add(Me.dgv_customUserRoles)
        Me.Panel1.Controls.Add(Me.btnOK)
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(366, 243)
        Me.Panel1.TabIndex = 0
        '
        'dgv_customUserRoles
        '
        Me.dgv_customUserRoles.AllowUserToAddRows = False
        Me.dgv_customUserRoles.AllowUserToDeleteRows = False
        Me.dgv_customUserRoles.AllowUserToResizeColumns = False
        Me.dgv_customUserRoles.AllowUserToResizeRows = False
        Me.dgv_customUserRoles.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgv_customUserRoles.BackgroundColor = System.Drawing.SystemColors.ButtonHighlight
        Me.dgv_customUserRoles.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgv_customUserRoles.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.userRole, Me.specifics})
        Me.dgv_customUserRoles.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.dgv_customUserRoles.Location = New System.Drawing.Point(2, 0)
        Me.dgv_customUserRoles.MultiSelect = False
        Me.dgv_customUserRoles.Name = "dgv_customUserRoles"
        Me.dgv_customUserRoles.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgv_customUserRoles.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.dgv_customUserRoles.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgv_customUserRoles.Size = New System.Drawing.Size(362, 197)
        Me.dgv_customUserRoles.TabIndex = 3
        '
        'userRole
        '
        Me.userRole.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.userRole.HeaderText = "Rolle"
        Me.userRole.Name = "userRole"
        '
        'specifics
        '
        Me.specifics.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.specifics.HeaderText = "Details"
        Me.specifics.Name = "specifics"
        '
        'btnOK
        '
        Me.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.btnOK.Location = New System.Drawing.Point(110, 204)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(144, 34)
        Me.btnOK.TabIndex = 2
        Me.btnOK.Text = "Auswählen"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'frmChooseCustomUserRole
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(364, 244)
        Me.ControlBox = False
        Me.Controls.Add(Me.Panel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmChooseCustomUserRole"
        Me.Text = "Wählen Sie Ihre Rolle"
        Me.TopMost = True
        Me.Panel1.ResumeLayout(False)
        CType(Me.dgv_customUserRoles, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Panel1 As Windows.Forms.Panel
    Friend WithEvents dgv_customUserRoles As Windows.Forms.DataGridView
    Friend WithEvents userRole As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents specifics As Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnOK As Windows.Forms.Button
End Class
