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
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.Abbrechen = New System.Windows.Forms.Button()
        Me.addToSession = New System.Windows.Forms.CheckBox()
        Me.lblStandvom = New System.Windows.Forms.Label()
        Me.requiredDate = New System.Windows.Forms.DateTimePicker()
        Me.loadAsSummary = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'ListBox1
        '
        Me.ListBox1.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.HorizontalScrollbar = True
        Me.ListBox1.ItemHeight = 16
        Me.ListBox1.Location = New System.Drawing.Point(13, 42)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.ListBox1.Size = New System.Drawing.Size(340, 180)
        Me.ListBox1.TabIndex = 0
        '
        'OKButton
        '
        Me.OKButton.Location = New System.Drawing.Point(65, 269)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(80, 22)
        Me.OKButton.TabIndex = 1
        Me.OKButton.Text = "OK"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'Abbrechen
        '
        Me.Abbrechen.Location = New System.Drawing.Point(204, 269)
        Me.Abbrechen.Name = "Abbrechen"
        Me.Abbrechen.Size = New System.Drawing.Size(80, 22)
        Me.Abbrechen.TabIndex = 2
        Me.Abbrechen.Text = "Abbrechen"
        Me.Abbrechen.UseVisualStyleBackColor = True
        '
        'addToSession
        '
        Me.addToSession.AutoSize = True
        Me.addToSession.Checked = True
        Me.addToSession.CheckState = System.Windows.Forms.CheckState.Checked
        Me.addToSession.Cursor = System.Windows.Forms.Cursors.Default
        Me.addToSession.Location = New System.Drawing.Point(12, 237)
        Me.addToSession.Name = "addToSession"
        Me.addToSession.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.addToSession.Size = New System.Drawing.Size(135, 17)
        Me.addToSession.TabIndex = 3
        Me.addToSession.Text = "zur Session hinzufügen"
        Me.addToSession.UseVisualStyleBackColor = True
        '
        'lblStandvom
        '
        Me.lblStandvom.AutoSize = True
        Me.lblStandvom.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStandvom.Location = New System.Drawing.Point(86, 19)
        Me.lblStandvom.Name = "lblStandvom"
        Me.lblStandvom.Size = New System.Drawing.Size(61, 13)
        Me.lblStandvom.TabIndex = 38
        Me.lblStandvom.Text = "Stand vom:"
        '
        'requiredDate
        '
        Me.requiredDate.Location = New System.Drawing.Point(153, 16)
        Me.requiredDate.Name = "requiredDate"
        Me.requiredDate.Size = New System.Drawing.Size(200, 20)
        Me.requiredDate.TabIndex = 39
        '
        'loadAsSummary
        '
        Me.loadAsSummary.AutoSize = True
        Me.loadAsSummary.Cursor = System.Windows.Forms.Cursors.Default
        Me.loadAsSummary.Location = New System.Drawing.Point(193, 237)
        Me.loadAsSummary.Name = "loadAsSummary"
        Me.loadAsSummary.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.loadAsSummary.Size = New System.Drawing.Size(152, 17)
        Me.loadAsSummary.TabIndex = 40
        Me.loadAsSummary.Text = "nur Summary Projekt laden"
        Me.loadAsSummary.UseVisualStyleBackColor = True
        '
        'frmLoadConstellation
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(366, 305)
        Me.Controls.Add(Me.loadAsSummary)
        Me.Controls.Add(Me.requiredDate)
        Me.Controls.Add(Me.lblStandvom)
        Me.Controls.Add(Me.addToSession)
        Me.Controls.Add(Me.Abbrechen)
        Me.Controls.Add(Me.OKButton)
        Me.Controls.Add(Me.ListBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmLoadConstellation"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Portfolio laden"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents ListBox1 As System.Windows.Forms.ListBox
    Public WithEvents OKButton As System.Windows.Forms.Button
    Public WithEvents Abbrechen As System.Windows.Forms.Button
    Public WithEvents addToSession As System.Windows.Forms.CheckBox
    Public WithEvents lblStandvom As System.Windows.Forms.Label
    Public WithEvents requiredDate As System.Windows.Forms.DateTimePicker
    Public WithEvents loadAsSummary As Windows.Forms.CheckBox
End Class
