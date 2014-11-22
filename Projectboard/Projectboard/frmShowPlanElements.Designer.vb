<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmShowPlanElements
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmShowPlanElements))
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.AbbrButton = New System.Windows.Forms.Button()
        Me.filterBox = New System.Windows.Forms.TextBox()
        Me.headerLine = New System.Windows.Forms.Label()
        Me.pictureZoom = New System.Windows.Forms.PictureBox()
        Me.rdbPhases = New System.Windows.Forms.RadioButton()
        Me.rdbMilestones = New System.Windows.Forms.RadioButton()
        Me.rdbRoles = New System.Windows.Forms.RadioButton()
        Me.rdbCosts = New System.Windows.Forms.RadioButton()
        Me.pictureCosts = New System.Windows.Forms.PictureBox()
        Me.pictureRoles = New System.Windows.Forms.PictureBox()
        Me.picturePhasen = New System.Windows.Forms.PictureBox()
        Me.pictureMilestones = New System.Windows.Forms.PictureBox()
        Me.chkbxShowObjects = New System.Windows.Forms.CheckBox()
        Me.chkbxCreateCharts = New System.Windows.Forms.CheckBox()
        Me.chkbxOneChart = New System.Windows.Forms.CheckBox()
        Me.repVorlagenDropbox = New System.Windows.Forms.ComboBox()
        Me.labelPPTVorlage = New System.Windows.Forms.Label()
        Me.statusLabel = New System.Windows.Forms.Label()
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        CType(Me.pictureZoom, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pictureCosts, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pictureRoles, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picturePhasen, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pictureMilestones, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ListBox1
        '
        Me.ListBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.ItemHeight = 16
        Me.ListBox1.Location = New System.Drawing.Point(12, 109)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple
        Me.ListBox1.Size = New System.Drawing.Size(349, 196)
        Me.ListBox1.Sorted = True
        Me.ListBox1.TabIndex = 0
        '
        'OKButton
        '
        Me.OKButton.Location = New System.Drawing.Point(76, 384)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(75, 23)
        Me.OKButton.TabIndex = 9
        Me.OKButton.Text = "Anzeigen"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'AbbrButton
        '
        Me.AbbrButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.AbbrButton.Location = New System.Drawing.Point(201, 384)
        Me.AbbrButton.Name = "AbbrButton"
        Me.AbbrButton.Size = New System.Drawing.Size(75, 23)
        Me.AbbrButton.TabIndex = 10
        Me.AbbrButton.Text = "Abbrechen"
        Me.AbbrButton.UseVisualStyleBackColor = True
        '
        'filterBox
        '
        Me.filterBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.filterBox.Location = New System.Drawing.Point(130, 76)
        Me.filterBox.Name = "filterBox"
        Me.filterBox.Size = New System.Drawing.Size(176, 22)
        Me.filterBox.TabIndex = 11
        '
        'headerLine
        '
        Me.headerLine.AutoSize = True
        Me.headerLine.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.headerLine.Location = New System.Drawing.Point(12, 76)
        Me.headerLine.Name = "headerLine"
        Me.headerLine.Size = New System.Drawing.Size(49, 16)
        Me.headerLine.TabIndex = 12
        Me.headerLine.Text = "Label1"
        '
        'pictureZoom
        '
        Me.pictureZoom.Image = Global.ExcelWorkbook1.My.Resources.Resources.zoom_out
        Me.pictureZoom.Location = New System.Drawing.Point(312, 69)
        Me.pictureZoom.Name = "pictureZoom"
        Me.pictureZoom.Size = New System.Drawing.Size(33, 33)
        Me.pictureZoom.TabIndex = 20
        Me.pictureZoom.TabStop = False
        '
        'rdbPhases
        '
        Me.rdbPhases.AutoSize = True
        Me.rdbPhases.Location = New System.Drawing.Point(19, 32)
        Me.rdbPhases.Name = "rdbPhases"
        Me.rdbPhases.Size = New System.Drawing.Size(14, 13)
        Me.rdbPhases.TabIndex = 2
        Me.rdbPhases.TabStop = True
        Me.rdbPhases.UseVisualStyleBackColor = True
        '
        'rdbMilestones
        '
        Me.rdbMilestones.AutoSize = True
        Me.rdbMilestones.Location = New System.Drawing.Point(110, 32)
        Me.rdbMilestones.Name = "rdbMilestones"
        Me.rdbMilestones.Size = New System.Drawing.Size(14, 13)
        Me.rdbMilestones.TabIndex = 3
        Me.rdbMilestones.TabStop = True
        Me.rdbMilestones.UseVisualStyleBackColor = True
        '
        'rdbRoles
        '
        Me.rdbRoles.AutoSize = True
        Me.rdbRoles.Location = New System.Drawing.Point(201, 32)
        Me.rdbRoles.Name = "rdbRoles"
        Me.rdbRoles.Size = New System.Drawing.Size(14, 13)
        Me.rdbRoles.TabIndex = 4
        Me.rdbRoles.TabStop = True
        Me.rdbRoles.UseVisualStyleBackColor = True
        '
        'rdbCosts
        '
        Me.rdbCosts.AutoSize = True
        Me.rdbCosts.Location = New System.Drawing.Point(292, 32)
        Me.rdbCosts.Name = "rdbCosts"
        Me.rdbCosts.Size = New System.Drawing.Size(14, 13)
        Me.rdbCosts.TabIndex = 5
        Me.rdbCosts.TabStop = True
        Me.rdbCosts.UseVisualStyleBackColor = True
        '
        'pictureCosts
        '
        Me.pictureCosts.Image = Global.ExcelWorkbook1.My.Resources.Resources.money2
        Me.pictureCosts.Location = New System.Drawing.Point(312, 21)
        Me.pictureCosts.Name = "pictureCosts"
        Me.pictureCosts.Size = New System.Drawing.Size(33, 33)
        Me.pictureCosts.TabIndex = 17
        Me.pictureCosts.TabStop = False
        '
        'pictureRoles
        '
        Me.pictureRoles.Image = Global.ExcelWorkbook1.My.Resources.Resources.businessmen
        Me.pictureRoles.Location = New System.Drawing.Point(221, 21)
        Me.pictureRoles.Name = "pictureRoles"
        Me.pictureRoles.Size = New System.Drawing.Size(33, 33)
        Me.pictureRoles.TabIndex = 14
        Me.pictureRoles.TabStop = False
        '
        'picturePhasen
        '
        Me.picturePhasen.Image = CType(resources.GetObject("picturePhasen.Image"), System.Drawing.Image)
        Me.picturePhasen.Location = New System.Drawing.Point(39, 21)
        Me.picturePhasen.Name = "picturePhasen"
        Me.picturePhasen.Size = New System.Drawing.Size(33, 33)
        Me.picturePhasen.TabIndex = 18
        Me.picturePhasen.TabStop = False
        '
        'pictureMilestones
        '
        Me.pictureMilestones.Image = CType(resources.GetObject("pictureMilestones.Image"), System.Drawing.Image)
        Me.pictureMilestones.Location = New System.Drawing.Point(130, 21)
        Me.pictureMilestones.Name = "pictureMilestones"
        Me.pictureMilestones.Size = New System.Drawing.Size(33, 33)
        Me.pictureMilestones.TabIndex = 19
        Me.pictureMilestones.TabStop = False
        '
        'chkbxShowObjects
        '
        Me.chkbxShowObjects.AutoSize = True
        Me.chkbxShowObjects.Location = New System.Drawing.Point(15, 311)
        Me.chkbxShowObjects.Name = "chkbxShowObjects"
        Me.chkbxShowObjects.Size = New System.Drawing.Size(136, 17)
        Me.chkbxShowObjects.TabIndex = 6
        Me.chkbxShowObjects.Text = "Planelemente anzeigen"
        Me.chkbxShowObjects.UseVisualStyleBackColor = True
        Me.chkbxShowObjects.Visible = False
        '
        'chkbxCreateCharts
        '
        Me.chkbxCreateCharts.AutoSize = True
        Me.chkbxCreateCharts.Location = New System.Drawing.Point(66, 311)
        Me.chkbxCreateCharts.Name = "chkbxCreateCharts"
        Me.chkbxCreateCharts.Size = New System.Drawing.Size(97, 17)
        Me.chkbxCreateCharts.TabIndex = 7
        Me.chkbxCreateCharts.Text = "Chart anzeigen"
        Me.chkbxCreateCharts.UseVisualStyleBackColor = True
        Me.chkbxCreateCharts.Visible = False
        '
        'chkbxOneChart
        '
        Me.chkbxOneChart.AutoSize = True
        Me.chkbxOneChart.Location = New System.Drawing.Point(243, 311)
        Me.chkbxOneChart.Name = "chkbxOneChart"
        Me.chkbxOneChart.Size = New System.Drawing.Size(118, 17)
        Me.chkbxOneChart.TabIndex = 8
        Me.chkbxOneChart.Text = "Alles in einem Chart"
        Me.chkbxOneChart.UseVisualStyleBackColor = True
        Me.chkbxOneChart.Visible = False
        '
        'repVorlagenDropbox
        '
        Me.repVorlagenDropbox.FormattingEnabled = True
        Me.repVorlagenDropbox.Location = New System.Drawing.Point(162, 343)
        Me.repVorlagenDropbox.Name = "repVorlagenDropbox"
        Me.repVorlagenDropbox.Size = New System.Drawing.Size(199, 21)
        Me.repVorlagenDropbox.TabIndex = 9
        Me.repVorlagenDropbox.Visible = False
        '
        'labelPPTVorlage
        '
        Me.labelPPTVorlage.AutoSize = True
        Me.labelPPTVorlage.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelPPTVorlage.Location = New System.Drawing.Point(12, 343)
        Me.labelPPTVorlage.Name = "labelPPTVorlage"
        Me.labelPPTVorlage.Size = New System.Drawing.Size(126, 16)
        Me.labelPPTVorlage.TabIndex = 10
        Me.labelPPTVorlage.Text = "Powerpoint Vorlage"
        Me.labelPPTVorlage.Visible = False
        '
        'statusLabel
        '
        Me.statusLabel.AutoSize = True
        Me.statusLabel.Location = New System.Drawing.Point(16, 417)
        Me.statusLabel.Name = "statusLabel"
        Me.statusLabel.Size = New System.Drawing.Size(39, 13)
        Me.statusLabel.TabIndex = 21
        Me.statusLabel.Text = "Label1"
        '
        'BackgroundWorker1
        '
        Me.BackgroundWorker1.WorkerReportsProgress = True
        Me.BackgroundWorker1.WorkerSupportsCancellation = True
        '
        'frmShowPlanElements
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(375, 439)
        Me.Controls.Add(Me.pictureMilestones)
        Me.Controls.Add(Me.chkbxOneChart)
        Me.Controls.Add(Me.picturePhasen)
        Me.Controls.Add(Me.labelPPTVorlage)
        Me.Controls.Add(Me.pictureRoles)
        Me.Controls.Add(Me.chkbxCreateCharts)
        Me.Controls.Add(Me.pictureCosts)
        Me.Controls.Add(Me.rdbCosts)
        Me.Controls.Add(Me.statusLabel)
        Me.Controls.Add(Me.rdbRoles)
        Me.Controls.Add(Me.chkbxShowObjects)
        Me.Controls.Add(Me.rdbMilestones)
        Me.Controls.Add(Me.repVorlagenDropbox)
        Me.Controls.Add(Me.rdbPhases)
        Me.Controls.Add(Me.pictureZoom)
        Me.Controls.Add(Me.headerLine)
        Me.Controls.Add(Me.filterBox)
        Me.Controls.Add(Me.AbbrButton)
        Me.Controls.Add(Me.OKButton)
        Me.Controls.Add(Me.ListBox1)
        Me.Name = "frmShowPlanElements"
        Me.Text = "Visualisieren von Plan-Objekten"
        Me.TopMost = True
        CType(Me.pictureZoom, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pictureCosts, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pictureRoles, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picturePhasen, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pictureMilestones, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents OKButton As System.Windows.Forms.Button
    Friend WithEvents AbbrButton As System.Windows.Forms.Button
    Friend WithEvents filterBox As System.Windows.Forms.TextBox
    Friend WithEvents headerLine As System.Windows.Forms.Label
    Friend WithEvents pictureZoom As System.Windows.Forms.PictureBox
    Friend WithEvents rdbPhases As System.Windows.Forms.RadioButton
    Friend WithEvents rdbMilestones As System.Windows.Forms.RadioButton
    Friend WithEvents rdbRoles As System.Windows.Forms.RadioButton
    Friend WithEvents rdbCosts As System.Windows.Forms.RadioButton
    Friend WithEvents pictureCosts As System.Windows.Forms.PictureBox
    Friend WithEvents pictureRoles As System.Windows.Forms.PictureBox
    Friend WithEvents picturePhasen As System.Windows.Forms.PictureBox
    Friend WithEvents pictureMilestones As System.Windows.Forms.PictureBox
    Friend WithEvents chkbxShowObjects As System.Windows.Forms.CheckBox
    Friend WithEvents chkbxCreateCharts As System.Windows.Forms.CheckBox
    Friend WithEvents chkbxOneChart As System.Windows.Forms.CheckBox
    Friend WithEvents repVorlagenDropbox As System.Windows.Forms.ComboBox
    Friend WithEvents labelPPTVorlage As System.Windows.Forms.Label
    Friend WithEvents statusLabel As System.Windows.Forms.Label
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
End Class
