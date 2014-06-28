<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmShowProjCharacteristics
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
        Me.timeSlider = New System.Windows.Forms.TrackBar()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.compareBeauftragung = New System.Windows.Forms.Button()
        Me.snapshotDate = New System.Windows.Forms.Label()
        Me.compareCurrent = New System.Windows.Forms.Button()
        Me.movetoBeauftragung = New System.Windows.Forms.Button()
        Me.typSelection = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.movetoNext = New System.Windows.Forms.Button()
        Me.movetoPrevious = New System.Windows.Forms.Button()
        Me.showMore = New System.Windows.Forms.Label()
        CType(Me.timeSlider, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'timeSlider
        '
        Me.timeSlider.Location = New System.Drawing.Point(131, 58)
        Me.timeSlider.Margin = New System.Windows.Forms.Padding(4)
        Me.timeSlider.Name = "timeSlider"
        Me.timeSlider.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.timeSlider.RightToLeftLayout = True
        Me.timeSlider.Size = New System.Drawing.Size(551, 56)
        Me.timeSlider.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(37, 60)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(79, 17)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Time Slider"
        '
        'compareBeauftragung
        '
        Me.compareBeauftragung.BackColor = System.Drawing.SystemColors.Control
        Me.compareBeauftragung.Location = New System.Drawing.Point(333, 183)
        Me.compareBeauftragung.Margin = New System.Windows.Forms.Padding(4)
        Me.compareBeauftragung.Name = "compareBeauftragung"
        Me.compareBeauftragung.Size = New System.Drawing.Size(140, 43)
        Me.compareBeauftragung.TabIndex = 4
        Me.compareBeauftragung.Text = "mit Beauftragung vergleichen"
        Me.compareBeauftragung.UseVisualStyleBackColor = False
        Me.compareBeauftragung.Visible = False
        '
        'snapshotDate
        '
        Me.snapshotDate.AutoSize = True
        Me.snapshotDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.snapshotDate.Location = New System.Drawing.Point(329, 25)
        Me.snapshotDate.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.snapshotDate.Name = "snapshotDate"
        Me.snapshotDate.Size = New System.Drawing.Size(54, 20)
        Me.snapshotDate.TabIndex = 7
        Me.snapshotDate.Text = "Heute"
        '
        'compareCurrent
        '
        Me.compareCurrent.BackColor = System.Drawing.SystemColors.Control
        Me.compareCurrent.Location = New System.Drawing.Point(531, 183)
        Me.compareCurrent.Margin = New System.Windows.Forms.Padding(4)
        Me.compareCurrent.Name = "compareCurrent"
        Me.compareCurrent.Size = New System.Drawing.Size(140, 43)
        Me.compareCurrent.TabIndex = 9
        Me.compareCurrent.Text = "mit aktuellem Stand vergleichen"
        Me.compareCurrent.UseVisualStyleBackColor = False
        Me.compareCurrent.Visible = False
        '
        'movetoBeauftragung
        '
        Me.movetoBeauftragung.BackColor = System.Drawing.SystemColors.Control
        Me.movetoBeauftragung.Location = New System.Drawing.Point(133, 183)
        Me.movetoBeauftragung.Margin = New System.Windows.Forms.Padding(4)
        Me.movetoBeauftragung.Name = "movetoBeauftragung"
        Me.movetoBeauftragung.Size = New System.Drawing.Size(140, 43)
        Me.movetoBeauftragung.TabIndex = 11
        Me.movetoBeauftragung.Text = "Positioniere auf Beauftragung"
        Me.movetoBeauftragung.UseVisualStyleBackColor = False
        Me.movetoBeauftragung.Visible = False
        '
        'typSelection
        '
        Me.typSelection.FormattingEnabled = True
        Me.typSelection.Location = New System.Drawing.Point(261, 134)
        Me.typSelection.Margin = New System.Windows.Forms.Padding(4)
        Me.typSelection.Name = "typSelection"
        Me.typSelection.Size = New System.Drawing.Size(284, 24)
        Me.typSelection.TabIndex = 14
        Me.typSelection.Visible = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(256, 116)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(89, 17)
        Me.Label3.TabIndex = 15
        Me.Label3.Text = "Änderung in "
        Me.Label3.Visible = False
        '
        'movetoNext
        '
        Me.movetoNext.BackColor = System.Drawing.SystemColors.Control
        Me.movetoNext.Image = Global.ExcelWorkbook1.My.Resources.Resources.Pfeil_rechts_32x32
        Me.movetoNext.Location = New System.Drawing.Point(576, 116)
        Me.movetoNext.Margin = New System.Windows.Forms.Padding(4)
        Me.movetoNext.Name = "movetoNext"
        Me.movetoNext.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.movetoNext.Size = New System.Drawing.Size(49, 46)
        Me.movetoNext.TabIndex = 13
        Me.movetoNext.UseVisualStyleBackColor = False
        Me.movetoNext.Visible = False
        '
        'movetoPrevious
        '
        Me.movetoPrevious.BackColor = System.Drawing.SystemColors.Control
        Me.movetoPrevious.Image = Global.ExcelWorkbook1.My.Resources.Resources.Pfeil_links_32x32
        Me.movetoPrevious.Location = New System.Drawing.Point(183, 116)
        Me.movetoPrevious.Margin = New System.Windows.Forms.Padding(4)
        Me.movetoPrevious.Name = "movetoPrevious"
        Me.movetoPrevious.Size = New System.Drawing.Size(49, 46)
        Me.movetoPrevious.TabIndex = 12
        Me.movetoPrevious.UseVisualStyleBackColor = False
        Me.movetoPrevious.Visible = False
        '
        'showMore
        '
        Me.showMore.AutoSize = True
        Me.showMore.Location = New System.Drawing.Point(41, 116)
        Me.showMore.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.showMore.Name = "showMore"
        Me.showMore.Size = New System.Drawing.Size(56, 17)
        Me.showMore.TabIndex = 16
        Me.showMore.Text = "more ..."
        '
        'frmShowProjCharacteristics
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(120.0!, 120.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.BackColor = System.Drawing.SystemColors.Window
        Me.ClientSize = New System.Drawing.Size(760, 254)
        Me.Controls.Add(Me.showMore)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.typSelection)
        Me.Controls.Add(Me.movetoNext)
        Me.Controls.Add(Me.movetoPrevious)
        Me.Controls.Add(Me.movetoBeauftragung)
        Me.Controls.Add(Me.compareCurrent)
        Me.Controls.Add(Me.snapshotDate)
        Me.Controls.Add(Me.compareBeauftragung)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.timeSlider)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "frmShowProjCharacteristics"
        Me.Text = "Projekt Historie"
        Me.TopMost = True
        CType(Me.timeSlider, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents timeSlider As System.Windows.Forms.TrackBar
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents compareBeauftragung As System.Windows.Forms.Button
    Friend WithEvents snapshotDate As System.Windows.Forms.Label
    Friend WithEvents compareCurrent As System.Windows.Forms.Button
    Friend WithEvents movetoBeauftragung As System.Windows.Forms.Button
    Friend WithEvents movetoPrevious As System.Windows.Forms.Button
    Friend WithEvents movetoNext As System.Windows.Forms.Button
    Friend WithEvents typSelection As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents showMore As System.Windows.Forms.Label
End Class
