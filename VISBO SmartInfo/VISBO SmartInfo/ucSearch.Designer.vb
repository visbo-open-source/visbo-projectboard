<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ucSearch
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
        Me.cathegoryList = New System.Windows.Forms.ComboBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.selListboxNames = New System.Windows.Forms.ListBox()
        Me.CheckBxMarker = New System.Windows.Forms.CheckBox()
        Me.PictureMarker = New System.Windows.Forms.PictureBox()
        Me.listboxNames = New System.Windows.Forms.ListBox()
        Me.filterText = New System.Windows.Forms.TextBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.shwOhneLight = New System.Windows.Forms.CheckBox()
        Me.shwRedLight = New System.Windows.Forms.CheckBox()
        Me.shwGreenLight = New System.Windows.Forms.CheckBox()
        Me.shwYellowLight = New System.Windows.Forms.CheckBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Panel1.SuspendLayout()
        CType(Me.PictureMarker, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cathegoryList
        '
        Me.cathegoryList.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cathegoryList.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cathegoryList.FormattingEnabled = True
        Me.cathegoryList.Location = New System.Drawing.Point(8, 79)
        Me.cathegoryList.Name = "cathegoryList"
        Me.cathegoryList.Size = New System.Drawing.Size(257, 24)
        Me.cathegoryList.TabIndex = 1
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.AutoScroll = True
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.selListboxNames)
        Me.Panel1.Controls.Add(Me.CheckBxMarker)
        Me.Panel1.Controls.Add(Me.PictureMarker)
        Me.Panel1.Controls.Add(Me.listboxNames)
        Me.Panel1.Controls.Add(Me.filterText)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Controls.Add(Me.PictureBox1)
        Me.Panel1.Controls.Add(Me.cathegoryList)
        Me.Panel1.Location = New System.Drawing.Point(0, 3)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(276, 740)
        Me.Panel1.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(7, 475)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(53, 13)
        Me.Label2.TabIndex = 47
        Me.Label2.Text = "Elements:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(7, 114)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(77, 13)
        Me.Label1.TabIndex = 46
        Me.Label1.Text = "Search results:"
        '
        'selListboxNames
        '
        Me.selListboxNames.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.selListboxNames.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.selListboxNames.FormattingEnabled = True
        Me.selListboxNames.HorizontalScrollbar = True
        Me.selListboxNames.ItemHeight = 16
        Me.selListboxNames.Location = New System.Drawing.Point(8, 493)
        Me.selListboxNames.MinimumSize = New System.Drawing.Size(4, 50)
        Me.selListboxNames.Name = "selListboxNames"
        Me.selListboxNames.Size = New System.Drawing.Size(257, 212)
        Me.selListboxNames.TabIndex = 45
        '
        'CheckBxMarker
        '
        Me.CheckBxMarker.AutoSize = True
        Me.CheckBxMarker.Location = New System.Drawing.Point(26, 5)
        Me.CheckBxMarker.Name = "CheckBxMarker"
        Me.CheckBxMarker.Size = New System.Drawing.Size(15, 14)
        Me.CheckBxMarker.TabIndex = 44
        Me.CheckBxMarker.UseVisualStyleBackColor = True
        '
        'PictureMarker
        '
        Me.PictureMarker.Image = Global.VISBO_SmartInfo.My.Resources.Resources.arrow_down_blue
        Me.PictureMarker.Location = New System.Drawing.Point(8, 3)
        Me.PictureMarker.Name = "PictureMarker"
        Me.PictureMarker.Size = New System.Drawing.Size(16, 16)
        Me.PictureMarker.TabIndex = 43
        Me.PictureMarker.TabStop = False
        '
        'listboxNames
        '
        Me.listboxNames.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.listboxNames.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.listboxNames.FormattingEnabled = True
        Me.listboxNames.HorizontalScrollbar = True
        Me.listboxNames.ItemHeight = 16
        Me.listboxNames.Location = New System.Drawing.Point(8, 134)
        Me.listboxNames.MinimumSize = New System.Drawing.Size(4, 50)
        Me.listboxNames.Name = "listboxNames"
        Me.listboxNames.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.listboxNames.Size = New System.Drawing.Size(257, 324)
        Me.listboxNames.Sorted = True
        Me.listboxNames.TabIndex = 42
        '
        'filterText
        '
        Me.filterText.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.filterText.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.filterText.Location = New System.Drawing.Point(8, 33)
        Me.filterText.Name = "filterText"
        Me.filterText.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.filterText.Size = New System.Drawing.Size(230, 23)
        Me.filterText.TabIndex = 39
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.Controls.Add(Me.shwOhneLight)
        Me.Panel2.Controls.Add(Me.shwRedLight)
        Me.Panel2.Controls.Add(Me.shwGreenLight)
        Me.Panel2.Controls.Add(Me.shwYellowLight)
        Me.Panel2.Location = New System.Drawing.Point(181, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(92, 27)
        Me.Panel2.TabIndex = 38
        '
        'shwOhneLight
        '
        Me.shwOhneLight.AutoSize = True
        Me.shwOhneLight.Location = New System.Drawing.Point(3, 3)
        Me.shwOhneLight.Name = "shwOhneLight"
        Me.shwOhneLight.Size = New System.Drawing.Size(15, 14)
        Me.shwOhneLight.TabIndex = 34
        Me.shwOhneLight.UseVisualStyleBackColor = True
        '
        'shwRedLight
        '
        Me.shwRedLight.AutoSize = True
        Me.shwRedLight.BackColor = System.Drawing.Color.Firebrick
        Me.shwRedLight.Location = New System.Drawing.Point(67, 3)
        Me.shwRedLight.Name = "shwRedLight"
        Me.shwRedLight.Size = New System.Drawing.Size(15, 14)
        Me.shwRedLight.TabIndex = 37
        Me.shwRedLight.UseVisualStyleBackColor = False
        '
        'shwGreenLight
        '
        Me.shwGreenLight.AutoSize = True
        Me.shwGreenLight.BackColor = System.Drawing.Color.LawnGreen
        Me.shwGreenLight.Location = New System.Drawing.Point(25, 3)
        Me.shwGreenLight.Name = "shwGreenLight"
        Me.shwGreenLight.Size = New System.Drawing.Size(15, 14)
        Me.shwGreenLight.TabIndex = 35
        Me.shwGreenLight.UseVisualStyleBackColor = False
        '
        'shwYellowLight
        '
        Me.shwYellowLight.AutoSize = True
        Me.shwYellowLight.BackColor = System.Drawing.Color.Yellow
        Me.shwYellowLight.Location = New System.Drawing.Point(46, 3)
        Me.shwYellowLight.Name = "shwYellowLight"
        Me.shwYellowLight.Size = New System.Drawing.Size(15, 14)
        Me.shwYellowLight.TabIndex = 36
        Me.shwYellowLight.UseVisualStyleBackColor = False
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.BackColor = System.Drawing.SystemColors.Control
        Me.PictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox1.Image = Global.VISBO_SmartInfo.My.Resources.Resources.view1
        Me.PictureBox1.Location = New System.Drawing.Point(237, 33)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(28, 25)
        Me.PictureBox1.TabIndex = 2
        Me.PictureBox1.TabStop = False
        '
        'ucSearch
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.Controls.Add(Me.Panel1)
        Me.Name = "ucSearch"
        Me.Size = New System.Drawing.Size(279, 746)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.PictureMarker, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents cathegoryList As System.Windows.Forms.ComboBox
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents shwOhneLight As System.Windows.Forms.CheckBox
    Friend WithEvents shwGreenLight As System.Windows.Forms.CheckBox
    Friend WithEvents shwYellowLight As System.Windows.Forms.CheckBox
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents shwRedLight As System.Windows.Forms.CheckBox
    Friend WithEvents filterText As System.Windows.Forms.TextBox
    Friend WithEvents listboxNames As System.Windows.Forms.ListBox
    Friend WithEvents PictureMarker As System.Windows.Forms.PictureBox
    Friend WithEvents CheckBxMarker As System.Windows.Forms.CheckBox
    Friend WithEvents selListboxNames As System.Windows.Forms.ListBox
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents Label1 As Windows.Forms.Label
End Class
