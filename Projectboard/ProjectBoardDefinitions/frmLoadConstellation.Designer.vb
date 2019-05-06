<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLoadConstellation
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmLoadConstellation))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.loadAsSummary = New System.Windows.Forms.CheckBox()
        Me.requiredDate = New System.Windows.Forms.DateTimePicker()
        Me.lblStandvom = New System.Windows.Forms.Label()
        Me.addToSession = New System.Windows.Forms.CheckBox()
        Me.Abbrechen = New System.Windows.Forms.Button()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.Controls.Add(Me.loadAsSummary)
        Me.Panel1.Controls.Add(Me.requiredDate)
        Me.Panel1.Controls.Add(Me.lblStandvom)
        Me.Panel1.Controls.Add(Me.addToSession)
        Me.Panel1.Controls.Add(Me.Abbrechen)
        Me.Panel1.Controls.Add(Me.OKButton)
        Me.Panel1.Controls.Add(Me.ListBox1)
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(577, 380)
        Me.Panel1.TabIndex = 0
        '
        'loadAsSummary
        '
        Me.loadAsSummary.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.loadAsSummary.AutoSize = True
        Me.loadAsSummary.Cursor = System.Windows.Forms.Cursors.Default
        Me.loadAsSummary.Location = New System.Drawing.Point(266, 293)
        Me.loadAsSummary.Margin = New System.Windows.Forms.Padding(4)
        Me.loadAsSummary.Name = "loadAsSummary"
        Me.loadAsSummary.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.loadAsSummary.Size = New System.Drawing.Size(299, 21)
        Me.loadAsSummary.TabIndex = 47
        Me.loadAsSummary.Text = "Summary Projekt berechnen und anzeigen"
        Me.loadAsSummary.UseVisualStyleBackColor = True
        '
        'requiredDate
        '
        Me.requiredDate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.requiredDate.Location = New System.Drawing.Point(311, 17)
        Me.requiredDate.Margin = New System.Windows.Forms.Padding(4)
        Me.requiredDate.Name = "requiredDate"
        Me.requiredDate.Size = New System.Drawing.Size(249, 22)
        Me.requiredDate.TabIndex = 46
        '
        'lblStandvom
        '
        Me.lblStandvom.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblStandvom.AutoSize = True
        Me.lblStandvom.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStandvom.Location = New System.Drawing.Point(227, 21)
        Me.lblStandvom.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblStandvom.Name = "lblStandvom"
        Me.lblStandvom.Size = New System.Drawing.Size(79, 17)
        Me.lblStandvom.TabIndex = 45
        Me.lblStandvom.Text = "Stand vom:"
        '
        'addToSession
        '
        Me.addToSession.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.addToSession.AutoSize = True
        Me.addToSession.Checked = True
        Me.addToSession.CheckState = System.Windows.Forms.CheckState.Checked
        Me.addToSession.Cursor = System.Windows.Forms.Cursors.Default
        Me.addToSession.Location = New System.Drawing.Point(11, 293)
        Me.addToSession.Margin = New System.Windows.Forms.Padding(4)
        Me.addToSession.Name = "addToSession"
        Me.addToSession.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.addToSession.Size = New System.Drawing.Size(178, 21)
        Me.addToSession.TabIndex = 44
        Me.addToSession.Text = "zur Session hinzufügen"
        Me.addToSession.UseVisualStyleBackColor = True
        '
        'Abbrechen
        '
        Me.Abbrechen.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Abbrechen.Location = New System.Drawing.Point(332, 336)
        Me.Abbrechen.Margin = New System.Windows.Forms.Padding(4)
        Me.Abbrechen.Name = "Abbrechen"
        Me.Abbrechen.Size = New System.Drawing.Size(100, 28)
        Me.Abbrechen.TabIndex = 43
        Me.Abbrechen.Text = "Abbrechen"
        Me.Abbrechen.UseVisualStyleBackColor = True
        '
        'OKButton
        '
        Me.OKButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.OKButton.Location = New System.Drawing.Point(75, 336)
        Me.OKButton.Margin = New System.Windows.Forms.Padding(4)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(100, 28)
        Me.OKButton.TabIndex = 42
        Me.OKButton.Text = "OK"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'ListBox1
        '
        Me.ListBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ListBox1.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.HorizontalScrollbar = True
        Me.ListBox1.ItemHeight = 19
        Me.ListBox1.Location = New System.Drawing.Point(12, 49)
        Me.ListBox1.Margin = New System.Windows.Forms.Padding(4)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.ListBox1.Size = New System.Drawing.Size(550, 213)
        Me.ListBox1.TabIndex = 41
        '
        'frmLoadConstellation
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(120.0!, 120.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(582, 381)
        Me.Controls.Add(Me.Panel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "frmLoadConstellation"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Portfolio laden"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Panel1 As Windows.Forms.Panel
    Public WithEvents loadAsSummary As Windows.Forms.CheckBox
    Public WithEvents requiredDate As Windows.Forms.DateTimePicker
    Public WithEvents lblStandvom As Windows.Forms.Label
    Public WithEvents addToSession As Windows.Forms.CheckBox
    Public WithEvents Abbrechen As Windows.Forms.Button
    Public WithEvents OKButton As Windows.Forms.Button
    Public WithEvents ListBox1 As Windows.Forms.ListBox
End Class
