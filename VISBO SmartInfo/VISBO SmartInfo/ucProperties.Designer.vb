<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ucProperties
    Inherits System.Windows.Forms.UserControl

    'UserControl überschreibt den Löschvorgang, um die Komponentenliste zu bereinigen.
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
        Me.components = New System.ComponentModel.Container()
        Me.eleName = New System.Windows.Forms.Label()
        Me.labelDate = New System.Windows.Forms.Label()
        Me.eleDatum = New System.Windows.Forms.Label()
        Me.labelRespons = New System.Windows.Forms.Label()
        Me.eleRespons = New System.Windows.Forms.Label()
        Me.labelAmpel = New System.Windows.Forms.Label()
        Me.eleAmpel = New System.Windows.Forms.TextBox()
        Me.percentDone = New System.Windows.Forms.Label()
        Me.labelDeliver = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.eleDeliverables = New System.Windows.Forms.RichTextBox()
        Me.eleAmpelText = New System.Windows.Forms.RichTextBox()
        Me.stdLinks = New System.Windows.Forms.GroupBox()
        Me.dreiDlnk = New System.Windows.Forms.PictureBox()
        Me.survlnk = New System.Windows.Forms.PictureBox()
        Me.medialnk = New System.Windows.Forms.PictureBox()
        Me.doclnk = New System.Windows.Forms.PictureBox()
        Me.myLinks = New System.Windows.Forms.GroupBox()
        Me.mydreiDlnk = New System.Windows.Forms.PictureBox()
        Me.mysurvlnk = New System.Windows.Forms.PictureBox()
        Me.mymedialnk = New System.Windows.Forms.PictureBox()
        Me.mydoclnk = New System.Windows.Forms.PictureBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.stdLinks.SuspendLayout()
        CType(Me.dreiDlnk, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.survlnk, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.medialnk, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.doclnk, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.myLinks.SuspendLayout()
        CType(Me.mydreiDlnk, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mysurvlnk, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mymedialnk, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mydoclnk, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'eleName
        '
        Me.eleName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.eleName.AutoEllipsis = True
        Me.eleName.AutoSize = True
        Me.eleName.Enabled = False
        Me.eleName.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.eleName.Location = New System.Drawing.Point(5, 12)
        Me.eleName.Name = "eleName"
        Me.eleName.Size = New System.Drawing.Size(300, 20)
        Me.eleName.TabIndex = 1
        Me.eleName.Text = "Name:                                                "
        Me.eleName.UseWaitCursor = True
        '
        'labelDate
        '
        Me.labelDate.AutoSize = True
        Me.labelDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelDate.Location = New System.Drawing.Point(5, 66)
        Me.labelDate.Name = "labelDate"
        Me.labelDate.Size = New System.Drawing.Size(42, 17)
        Me.labelDate.TabIndex = 2
        Me.labelDate.Text = "Date:"
        '
        'eleDatum
        '
        Me.eleDatum.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.eleDatum.AutoSize = True
        Me.eleDatum.Enabled = False
        Me.eleDatum.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
        Me.eleDatum.Location = New System.Drawing.Point(93, 66)
        Me.eleDatum.Name = "eleDatum"
        Me.eleDatum.Size = New System.Drawing.Size(204, 17)
        Me.eleDatum.TabIndex = 3
        Me.eleDatum.Text = "                                                 "
        '
        'labelRespons
        '
        Me.labelRespons.AutoSize = True
        Me.labelRespons.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelRespons.Location = New System.Drawing.Point(6, 101)
        Me.labelRespons.Name = "labelRespons"
        Me.labelRespons.Size = New System.Drawing.Size(90, 17)
        Me.labelRespons.TabIndex = 4
        Me.labelRespons.Text = "Responsible:"
        '
        'eleRespons
        '
        Me.eleRespons.AutoSize = True
        Me.eleRespons.Enabled = False
        Me.eleRespons.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.eleRespons.Location = New System.Drawing.Point(94, 101)
        Me.eleRespons.Name = "eleRespons"
        Me.eleRespons.Size = New System.Drawing.Size(92, 17)
        Me.eleRespons.TabIndex = 5
        Me.eleRespons.Text = "                     "
        '
        'labelAmpel
        '
        Me.labelAmpel.AutoSize = True
        Me.labelAmpel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelAmpel.Location = New System.Drawing.Point(6, 138)
        Me.labelAmpel.Name = "labelAmpel"
        Me.labelAmpel.Size = New System.Drawing.Size(82, 17)
        Me.labelAmpel.TabIndex = 6
        Me.labelAmpel.Text = "Traffic light:"
        '
        'eleAmpel
        '
        Me.eleAmpel.BackColor = System.Drawing.Color.Silver
        Me.eleAmpel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.eleAmpel.Cursor = System.Windows.Forms.Cursors.Default
        Me.eleAmpel.Enabled = False
        Me.eleAmpel.Location = New System.Drawing.Point(94, 137)
        Me.eleAmpel.Name = "eleAmpel"
        Me.eleAmpel.ReadOnly = True
        Me.eleAmpel.Size = New System.Drawing.Size(23, 20)
        Me.eleAmpel.TabIndex = 7
        '
        'percentDone
        '
        Me.percentDone.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.percentDone.AutoSize = True
        Me.percentDone.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.percentDone.Location = New System.Drawing.Point(247, 138)
        Me.percentDone.Name = "percentDone"
        Me.percentDone.Size = New System.Drawing.Size(0, 17)
        Me.percentDone.TabIndex = 13
        Me.percentDone.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'labelDeliver
        '
        Me.labelDeliver.AutoSize = True
        Me.labelDeliver.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelDeliver.Location = New System.Drawing.Point(6, 316)
        Me.labelDeliver.Name = "labelDeliver"
        Me.labelDeliver.Size = New System.Drawing.Size(90, 17)
        Me.labelDeliver.TabIndex = 9
        Me.labelDeliver.Text = "Deliverables:"
        '
        'Panel1
        '
        Me.Panel1.AutoScroll = True
        Me.Panel1.AutoSize = True
        Me.Panel1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.Panel1.Location = New System.Drawing.Point(10, 365)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(0, 0)
        Me.Panel1.TabIndex = 12
        '
        'eleDeliverables
        '
        Me.eleDeliverables.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.eleDeliverables.BackColor = System.Drawing.Color.White
        Me.eleDeliverables.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.eleDeliverables.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.eleDeliverables.Location = New System.Drawing.Point(10, 339)
        Me.eleDeliverables.Name = "eleDeliverables"
        Me.eleDeliverables.ReadOnly = True
        Me.eleDeliverables.Size = New System.Drawing.Size(276, 143)
        Me.eleDeliverables.TabIndex = 14
        Me.eleDeliverables.Text = ""
        '
        'eleAmpelText
        '
        Me.eleAmpelText.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.eleAmpelText.BackColor = System.Drawing.Color.White
        Me.eleAmpelText.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.eleAmpelText.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.eleAmpelText.HideSelection = False
        Me.eleAmpelText.Location = New System.Drawing.Point(10, 161)
        Me.eleAmpelText.Name = "eleAmpelText"
        Me.eleAmpelText.ReadOnly = True
        Me.eleAmpelText.Size = New System.Drawing.Size(276, 143)
        Me.eleAmpelText.TabIndex = 15
        Me.eleAmpelText.Text = ""
        '
        'stdLinks
        '
        Me.stdLinks.Controls.Add(Me.dreiDlnk)
        Me.stdLinks.Controls.Add(Me.survlnk)
        Me.stdLinks.Controls.Add(Me.medialnk)
        Me.stdLinks.Controls.Add(Me.doclnk)
        Me.stdLinks.Location = New System.Drawing.Point(10, 507)
        Me.stdLinks.Name = "stdLinks"
        Me.stdLinks.Size = New System.Drawing.Size(276, 55)
        Me.stdLinks.TabIndex = 21
        Me.stdLinks.TabStop = False
        Me.stdLinks.Text = "standard Connections"
        '
        'dreiDlnk
        '
        Me.dreiDlnk.Image = Global.VISBO_SmartInfo.My.Resources.Resources._3d
        Me.dreiDlnk.Location = New System.Drawing.Point(226, 21)
        Me.dreiDlnk.Name = "dreiDlnk"
        Me.dreiDlnk.Size = New System.Drawing.Size(26, 27)
        Me.dreiDlnk.TabIndex = 2
        Me.dreiDlnk.TabStop = False
        '
        'survlnk
        '
        Me.survlnk.Image = Global.VISBO_SmartInfo.My.Resources.Resources.surveillance_camera
        Me.survlnk.Location = New System.Drawing.Point(153, 21)
        Me.survlnk.Name = "survlnk"
        Me.survlnk.Size = New System.Drawing.Size(26, 27)
        Me.survlnk.TabIndex = 2
        Me.survlnk.TabStop = False
        '
        'medialnk
        '
        Me.medialnk.Image = Global.VISBO_SmartInfo.My.Resources.Resources.camera
        Me.medialnk.Location = New System.Drawing.Point(80, 21)
        Me.medialnk.Name = "medialnk"
        Me.medialnk.Size = New System.Drawing.Size(26, 27)
        Me.medialnk.TabIndex = 1
        Me.medialnk.TabStop = False
        '
        'doclnk
        '
        Me.doclnk.Image = Global.VISBO_SmartInfo.My.Resources.Resources.documents
        Me.doclnk.Location = New System.Drawing.Point(7, 21)
        Me.doclnk.Name = "doclnk"
        Me.doclnk.Size = New System.Drawing.Size(26, 27)
        Me.doclnk.TabIndex = 0
        Me.doclnk.TabStop = False
        '
        'myLinks
        '
        Me.myLinks.Controls.Add(Me.mydreiDlnk)
        Me.myLinks.Controls.Add(Me.mysurvlnk)
        Me.myLinks.Controls.Add(Me.mymedialnk)
        Me.myLinks.Controls.Add(Me.mydoclnk)
        Me.myLinks.Location = New System.Drawing.Point(10, 570)
        Me.myLinks.Name = "myLinks"
        Me.myLinks.Size = New System.Drawing.Size(276, 55)
        Me.myLinks.TabIndex = 22
        Me.myLinks.TabStop = False
        Me.myLinks.Text = "my Connections"
        '
        'mydreiDlnk
        '
        Me.mydreiDlnk.Image = Global.VISBO_SmartInfo.My.Resources.Resources._3d_plus
        Me.mydreiDlnk.Location = New System.Drawing.Point(226, 20)
        Me.mydreiDlnk.Name = "mydreiDlnk"
        Me.mydreiDlnk.Size = New System.Drawing.Size(26, 27)
        Me.mydreiDlnk.TabIndex = 6
        Me.mydreiDlnk.TabStop = False
        '
        'mysurvlnk
        '
        Me.mysurvlnk.Image = Global.VISBO_SmartInfo.My.Resources.Resources.surveillance_camera_plus
        Me.mysurvlnk.Location = New System.Drawing.Point(153, 20)
        Me.mysurvlnk.Name = "mysurvlnk"
        Me.mysurvlnk.Size = New System.Drawing.Size(26, 27)
        Me.mysurvlnk.TabIndex = 5
        Me.mysurvlnk.TabStop = False
        '
        'mymedialnk
        '
        Me.mymedialnk.Image = Global.VISBO_SmartInfo.My.Resources.Resources.camera_plus
        Me.mymedialnk.Location = New System.Drawing.Point(80, 20)
        Me.mymedialnk.Name = "mymedialnk"
        Me.mymedialnk.Size = New System.Drawing.Size(26, 27)
        Me.mymedialnk.TabIndex = 4
        Me.mymedialnk.TabStop = False
        '
        'mydoclnk
        '
        Me.mydoclnk.Image = Global.VISBO_SmartInfo.My.Resources.Resources.documents_plus
        Me.mydoclnk.Location = New System.Drawing.Point(7, 20)
        Me.mydoclnk.Name = "mydoclnk"
        Me.mydoclnk.Size = New System.Drawing.Size(26, 27)
        Me.mydoclnk.TabIndex = 3
        Me.mydoclnk.TabStop = False
        '
        'ucProperties
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.Controls.Add(Me.eleAmpelText)
        Me.Controls.Add(Me.eleDeliverables)
        Me.Controls.Add(Me.labelDeliver)
        Me.Controls.Add(Me.percentDone)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.eleAmpel)
        Me.Controls.Add(Me.labelAmpel)
        Me.Controls.Add(Me.eleRespons)
        Me.Controls.Add(Me.labelRespons)
        Me.Controls.Add(Me.eleDatum)
        Me.Controls.Add(Me.labelDate)
        Me.Controls.Add(Me.eleName)
        Me.Controls.Add(Me.stdLinks)
        Me.Controls.Add(Me.myLinks)
        Me.Name = "ucProperties"
        Me.Size = New System.Drawing.Size(299, 861)
        Me.stdLinks.ResumeLayout(False)
        CType(Me.dreiDlnk, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.survlnk, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.medialnk, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.doclnk, System.ComponentModel.ISupportInitialize).EndInit()
        Me.myLinks.ResumeLayout(False)
        CType(Me.mydreiDlnk, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mysurvlnk, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mymedialnk, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mydoclnk, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents eleName As System.Windows.Forms.Label
    Friend WithEvents labelDate As System.Windows.Forms.Label
    Friend WithEvents eleDatum As System.Windows.Forms.Label
    Friend WithEvents labelRespons As System.Windows.Forms.Label
    Friend WithEvents eleRespons As System.Windows.Forms.Label
    Friend WithEvents labelAmpel As System.Windows.Forms.Label
    Friend WithEvents eleAmpel As System.Windows.Forms.TextBox
    Friend WithEvents percentDone As System.Windows.Forms.Label
    Friend WithEvents labelDeliver As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents eleDeliverables As Windows.Forms.RichTextBox
    Friend WithEvents eleAmpelText As Windows.Forms.RichTextBox
    Friend WithEvents stdLinks As Windows.Forms.GroupBox
    Friend WithEvents myLinks As Windows.Forms.GroupBox
    Friend WithEvents medialnk As Windows.Forms.PictureBox
    Friend WithEvents doclnk As Windows.Forms.PictureBox
    Friend WithEvents survlnk As Windows.Forms.PictureBox
    Friend WithEvents dreiDlnk As Windows.Forms.PictureBox
    Friend WithEvents mydoclnk As Windows.Forms.PictureBox
    Friend WithEvents mymedialnk As Windows.Forms.PictureBox
    Friend WithEvents mysurvlnk As Windows.Forms.PictureBox
    Friend WithEvents mydreiDlnk As Windows.Forms.PictureBox
    Friend WithEvents ToolTip1 As Windows.Forms.ToolTip
End Class
