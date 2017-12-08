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
        Me.SuspendLayout()
        '
        'eleName
        '
        Me.eleName.AutoEllipsis = True
        Me.eleName.AutoSize = True
        Me.eleName.Enabled = False
        Me.eleName.Font = New System.Drawing.Font("Segoe UI Emoji", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.eleName.Location = New System.Drawing.Point(3, 12)
        Me.eleName.MaximumSize = New System.Drawing.Size(0, 42)
        Me.eleName.Name = "eleName"
        Me.eleName.Size = New System.Drawing.Size(130, 21)
        Me.eleName.TabIndex = 1
        Me.eleName.Text = "Name:              "
        Me.eleName.UseWaitCursor = True
        '
        'labelDate
        '
        Me.labelDate.AutoSize = True
        Me.labelDate.Font = New System.Drawing.Font("Segoe UI Emoji", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelDate.Location = New System.Drawing.Point(3, 52)
        Me.labelDate.Name = "labelDate"
        Me.labelDate.Size = New System.Drawing.Size(41, 19)
        Me.labelDate.TabIndex = 2
        Me.labelDate.Text = "Date:"
        '
        'eleDatum
        '
        Me.eleDatum.AutoSize = True
        Me.eleDatum.Enabled = False
        Me.eleDatum.Font = New System.Drawing.Font("Segoe UI Emoji", 10.0!)
        Me.eleDatum.Location = New System.Drawing.Point(93, 52)
        Me.eleDatum.Name = "eleDatum"
        Me.eleDatum.Size = New System.Drawing.Size(205, 19)
        Me.eleDatum.TabIndex = 3
        Me.eleDatum.Text = "                                                 "
        '
        'labelRespons
        '
        Me.labelRespons.AutoSize = True
        Me.labelRespons.Font = New System.Drawing.Font("Segoe UI Emoji", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelRespons.Location = New System.Drawing.Point(4, 86)
        Me.labelRespons.Name = "labelRespons"
        Me.labelRespons.Size = New System.Drawing.Size(84, 19)
        Me.labelRespons.TabIndex = 4
        Me.labelRespons.Text = "Responsible:"
        '
        'eleRespons
        '
        Me.eleRespons.AutoSize = True
        Me.eleRespons.Enabled = False
        Me.eleRespons.Font = New System.Drawing.Font("Segoe UI Emoji", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.eleRespons.Location = New System.Drawing.Point(94, 88)
        Me.eleRespons.Name = "eleRespons"
        Me.eleRespons.Size = New System.Drawing.Size(93, 19)
        Me.eleRespons.TabIndex = 5
        Me.eleRespons.Text = "                     "
        '
        'labelAmpel
        '
        Me.labelAmpel.AutoSize = True
        Me.labelAmpel.Font = New System.Drawing.Font("Segoe UI Emoji", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelAmpel.Location = New System.Drawing.Point(3, 122)
        Me.labelAmpel.Name = "labelAmpel"
        Me.labelAmpel.Size = New System.Drawing.Size(79, 19)
        Me.labelAmpel.TabIndex = 6
        Me.labelAmpel.Text = "Traffic light:"
        '
        'eleAmpel
        '
        Me.eleAmpel.BackColor = System.Drawing.Color.Silver
        Me.eleAmpel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.eleAmpel.Cursor = System.Windows.Forms.Cursors.Default
        Me.eleAmpel.Enabled = False
        Me.eleAmpel.Location = New System.Drawing.Point(94, 121)
        Me.eleAmpel.Name = "eleAmpel"
        Me.eleAmpel.ReadOnly = True
        Me.eleAmpel.Size = New System.Drawing.Size(23, 20)
        Me.eleAmpel.TabIndex = 7
        '
        'percentDone
        '
        Me.percentDone.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.percentDone.AutoSize = True
        Me.percentDone.Font = New System.Drawing.Font("Segoe UI Emoji", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.percentDone.Location = New System.Drawing.Point(247, 128)
        Me.percentDone.Name = "percentDone"
        Me.percentDone.Size = New System.Drawing.Size(0, 19)
        Me.percentDone.TabIndex = 13
        Me.percentDone.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'labelDeliver
        '
        Me.labelDeliver.AutoSize = True
        Me.labelDeliver.Font = New System.Drawing.Font("Segoe UI Emoji", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelDeliver.Location = New System.Drawing.Point(4, 296)
        Me.labelDeliver.Name = "labelDeliver"
        Me.labelDeliver.Size = New System.Drawing.Size(85, 19)
        Me.labelDeliver.TabIndex = 9
        Me.labelDeliver.Text = "Deliverables:"
        '
        'Panel1
        '
        Me.Panel1.AutoScroll = True
        Me.Panel1.AutoSize = True
        Me.Panel1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.Panel1.Location = New System.Drawing.Point(10, 356)
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
        Me.eleDeliverables.Location = New System.Drawing.Point(10, 318)
        Me.eleDeliverables.Name = "eleDeliverables"
        Me.eleDeliverables.ReadOnly = True
        Me.eleDeliverables.Size = New System.Drawing.Size(276, 139)
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
        Me.eleAmpelText.Location = New System.Drawing.Point(10, 144)
        Me.eleAmpelText.Name = "eleAmpelText"
        Me.eleAmpelText.ReadOnly = True
        Me.eleAmpelText.Size = New System.Drawing.Size(276, 139)
        Me.eleAmpelText.TabIndex = 15
        Me.eleAmpelText.Text = ""
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
        Me.Name = "ucProperties"
        Me.Size = New System.Drawing.Size(299, 839)
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
End Class
