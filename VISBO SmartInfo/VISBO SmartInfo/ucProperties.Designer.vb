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
        Me.eleAmpelText = New System.Windows.Forms.TextBox()
        Me.percentDone = New System.Windows.Forms.Label()
        Me.eleDeliverables = New System.Windows.Forms.TextBox()
        Me.labelDeliver = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'eleName
        '
        Me.eleName.AutoSize = True
        Me.eleName.Enabled = False
        Me.eleName.Font = New System.Drawing.Font("Segoe UI Emoji", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.eleName.Location = New System.Drawing.Point(3, 12)
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
        Me.labelAmpel.Location = New System.Drawing.Point(6, 122)
        Me.labelAmpel.Name = "labelAmpel"
        Me.labelAmpel.Size = New System.Drawing.Size(77, 19)
        Me.labelAmpel.TabIndex = 6
        Me.labelAmpel.Text = "traffic light:"
        '
        'eleAmpel
        '
        Me.eleAmpel.BackColor = System.Drawing.Color.Silver
        Me.eleAmpel.Enabled = False
        Me.eleAmpel.Location = New System.Drawing.Point(94, 121)
        Me.eleAmpel.Name = "eleAmpel"
        Me.eleAmpel.Size = New System.Drawing.Size(23, 20)
        Me.eleAmpel.TabIndex = 7
        '
        'eleAmpelText
        '
        Me.eleAmpelText.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.eleAmpelText.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.eleAmpelText.Enabled = False
        Me.eleAmpelText.Font = New System.Drawing.Font("Segoe UI Emoji", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.eleAmpelText.Location = New System.Drawing.Point(9, 155)
        Me.eleAmpelText.Multiline = True
        Me.eleAmpelText.Name = "eleAmpelText"
        Me.eleAmpelText.ReadOnly = True
        Me.eleAmpelText.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.eleAmpelText.Size = New System.Drawing.Size(281, 176)
        Me.eleAmpelText.TabIndex = 8
        '
        'percentDone
        '
        Me.percentDone.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.percentDone.AutoSize = True
        Me.percentDone.Font = New System.Drawing.Font("Segoe UI Emoji", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.percentDone.Location = New System.Drawing.Point(251, 128)
        Me.percentDone.Name = "percentDone"
        Me.percentDone.Size = New System.Drawing.Size(0, 19)
        Me.percentDone.TabIndex = 13
        Me.percentDone.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'eleDeliverables
        '
        Me.eleDeliverables.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.eleDeliverables.BackColor = System.Drawing.SystemColors.Window
        Me.eleDeliverables.Enabled = False
        Me.eleDeliverables.Font = New System.Drawing.Font("Segoe UI Emoji", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.eleDeliverables.Location = New System.Drawing.Point(0, 34)
        Me.eleDeliverables.Multiline = True
        Me.eleDeliverables.Name = "eleDeliverables"
        Me.eleDeliverables.ReadOnly = True
        Me.eleDeliverables.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.eleDeliverables.Size = New System.Drawing.Size(280, 486)
        Me.eleDeliverables.TabIndex = 10
        '
        'labelDeliver
        '
        Me.labelDeliver.AutoSize = True
        Me.labelDeliver.Dock = System.Windows.Forms.DockStyle.Top
        Me.labelDeliver.Font = New System.Drawing.Font("Segoe UI Emoji", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelDeliver.Location = New System.Drawing.Point(0, 0)
        Me.labelDeliver.Name = "labelDeliver"
        Me.labelDeliver.Size = New System.Drawing.Size(85, 19)
        Me.labelDeliver.TabIndex = 9
        Me.labelDeliver.Text = "Deliverables:"
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.AutoScroll = True
        Me.Panel1.AutoSize = True
        Me.Panel1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.Panel1.Controls.Add(Me.labelDeliver)
        Me.Panel1.Controls.Add(Me.eleDeliverables)
        Me.Panel1.Location = New System.Drawing.Point(10, 356)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(280, 520)
        Me.Panel1.TabIndex = 12
        '
        'ucProperties
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.Controls.Add(Me.percentDone)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.eleAmpelText)
        Me.Controls.Add(Me.eleAmpel)
        Me.Controls.Add(Me.labelAmpel)
        Me.Controls.Add(Me.eleRespons)
        Me.Controls.Add(Me.labelRespons)
        Me.Controls.Add(Me.eleDatum)
        Me.Controls.Add(Me.labelDate)
        Me.Controls.Add(Me.eleName)
        Me.Name = "ucProperties"
        Me.Size = New System.Drawing.Size(303, 900)
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
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
    Friend WithEvents eleAmpelText As System.Windows.Forms.TextBox
    Friend WithEvents percentDone As System.Windows.Forms.Label
    Friend WithEvents eleDeliverables As System.Windows.Forms.TextBox
    Friend WithEvents labelDeliver As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel

End Class
