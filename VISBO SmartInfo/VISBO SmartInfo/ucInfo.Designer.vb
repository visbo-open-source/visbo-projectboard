<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class ucInfo
    Inherits System.Windows.Forms.UserControl

    'UserControl überschreibt den Löschvorgang, um die Komponentenliste zu bereinigen.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ucInfo))
        Me.labelDate = New System.Windows.Forms.Label()
        Me.labelRespons = New System.Windows.Forms.Label()
        Me.labelAmpel = New System.Windows.Forms.Label()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.eleAmpelText = New System.Windows.Forms.Label()
        Me.eleDeliverables = New System.Windows.Forms.TextBox()
        Me.labelDeliver = New System.Windows.Forms.Label()
        Me.eleDatum = New System.Windows.Forms.Label()
        Me.eleRespons = New System.Windows.Forms.Label()
        Me.eleAmpel = New System.Windows.Forms.Label()
        Me.FlowLayoutPanel1 = New System.Windows.Forms.FlowLayoutPanel()
        Me.eleName = New System.Windows.Forms.Label()
        Me.eleType = New System.Windows.Forms.Label()
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.FlowLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'labelDate
        '
        Me.labelDate.AutoSize = True
        Me.labelDate.Location = New System.Drawing.Point(3, 8)
        Me.labelDate.Margin = New System.Windows.Forms.Padding(3, 8, 3, 0)
        Me.labelDate.Name = "labelDate"
        Me.labelDate.Size = New System.Drawing.Size(41, 13)
        Me.labelDate.TabIndex = 3
        Me.labelDate.Text = "Datum:"
        '
        'labelRespons
        '
        Me.labelRespons.AutoSize = True
        Me.labelRespons.Location = New System.Drawing.Point(3, 29)
        Me.labelRespons.Margin = New System.Windows.Forms.Padding(3, 8, 3, 0)
        Me.labelRespons.Name = "labelRespons"
        Me.labelRespons.Size = New System.Drawing.Size(57, 13)
        Me.labelRespons.TabIndex = 4
        Me.labelRespons.Text = "Zuständig:"
        '
        'labelAmpel
        '
        Me.labelAmpel.AutoSize = True
        Me.labelAmpel.Location = New System.Drawing.Point(3, 50)
        Me.labelAmpel.Margin = New System.Windows.Forms.Padding(3, 8, 3, 0)
        Me.labelAmpel.Name = "labelAmpel"
        Me.labelAmpel.Size = New System.Drawing.Size(39, 13)
        Me.labelAmpel.TabIndex = 5
        Me.labelAmpel.Text = "Ampel:"
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.AutoScrollMargin = New System.Drawing.Size(10, 10)
        Me.TableLayoutPanel1.BackColor = System.Drawing.SystemColors.Control
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 80.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel1.Controls.Add(Me.labelAmpel, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.labelRespons, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.labelDate, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.eleAmpelText, 0, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.eleDeliverables, 1, 5)
        Me.TableLayoutPanel1.Controls.Add(Me.labelDeliver, 0, 5)
        Me.TableLayoutPanel1.Controls.Add(Me.eleDatum, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.eleRespons, 1, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.eleAmpel, 1, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.RichTextBox1, 0, 7)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(3, 43)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 9
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(229, 457)
        Me.TableLayoutPanel1.TabIndex = 6
        '
        'eleAmpelText
        '
        Me.eleAmpelText.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.eleAmpelText.AutoSize = True
        Me.eleAmpelText.BackColor = System.Drawing.Color.Transparent
        Me.TableLayoutPanel1.SetColumnSpan(Me.eleAmpelText, 2)
        Me.eleAmpelText.Location = New System.Drawing.Point(3, 71)
        Me.eleAmpelText.Margin = New System.Windows.Forms.Padding(3, 8, 3, 0)
        Me.eleAmpelText.MaximumSize = New System.Drawing.Size(220, 0)
        Me.eleAmpelText.Name = "eleAmpelText"
        Me.eleAmpelText.Size = New System.Drawing.Size(220, 104)
        Me.eleAmpelText.TabIndex = 6
        Me.eleAmpelText.Text = resources.GetString("eleAmpelText.Text")
        '
        'eleDeliverables
        '
        Me.eleDeliverables.AcceptsReturn = True
        Me.eleDeliverables.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.eleDeliverables.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TableLayoutPanel1.SetColumnSpan(Me.eleDeliverables, 2)
        Me.eleDeliverables.Location = New System.Drawing.Point(3, 199)
        Me.eleDeliverables.Multiline = True
        Me.eleDeliverables.Name = "eleDeliverables"
        Me.eleDeliverables.ReadOnly = True
        Me.eleDeliverables.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.eleDeliverables.Size = New System.Drawing.Size(223, 111)
        Me.eleDeliverables.TabIndex = 7
        Me.eleDeliverables.Text = resources.GetString("eleDeliverables.Text")
        '
        'labelDeliver
        '
        Me.labelDeliver.AutoSize = True
        Me.labelDeliver.Location = New System.Drawing.Point(3, 183)
        Me.labelDeliver.Margin = New System.Windows.Forms.Padding(3, 8, 3, 0)
        Me.labelDeliver.Name = "labelDeliver"
        Me.labelDeliver.Size = New System.Drawing.Size(68, 13)
        Me.labelDeliver.TabIndex = 8
        Me.labelDeliver.Text = "Deliverables:"
        '
        'eleDatum
        '
        Me.eleDatum.AutoSize = True
        Me.eleDatum.Location = New System.Drawing.Point(83, 8)
        Me.eleDatum.Margin = New System.Windows.Forms.Padding(3, 8, 3, 0)
        Me.eleDatum.Name = "eleDatum"
        Me.eleDatum.Size = New System.Drawing.Size(39, 13)
        Me.eleDatum.TabIndex = 9
        Me.eleDatum.Text = "Label8"
        '
        'eleRespons
        '
        Me.eleRespons.AutoSize = True
        Me.eleRespons.Location = New System.Drawing.Point(83, 29)
        Me.eleRespons.Margin = New System.Windows.Forms.Padding(3, 8, 3, 0)
        Me.eleRespons.Name = "eleRespons"
        Me.eleRespons.Size = New System.Drawing.Size(39, 13)
        Me.eleRespons.TabIndex = 10
        Me.eleRespons.Text = "Label9"
        '
        'eleAmpel
        '
        Me.eleAmpel.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.eleAmpel.AutoSize = True
        Me.eleAmpel.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.eleAmpel.Location = New System.Drawing.Point(83, 50)
        Me.eleAmpel.Margin = New System.Windows.Forms.Padding(3, 8, 3, 0)
        Me.eleAmpel.Name = "eleAmpel"
        Me.eleAmpel.Size = New System.Drawing.Size(143, 13)
        Me.eleAmpel.TabIndex = 11
        '
        'FlowLayoutPanel1
        '
        Me.FlowLayoutPanel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.FlowLayoutPanel1.AutoSize = True
        Me.FlowLayoutPanel1.Controls.Add(Me.eleName)
        Me.FlowLayoutPanel1.Controls.Add(Me.eleType)
        Me.FlowLayoutPanel1.FlowDirection = System.Windows.Forms.FlowDirection.TopDown
        Me.FlowLayoutPanel1.Location = New System.Drawing.Point(3, 2)
        Me.FlowLayoutPanel1.Name = "FlowLayoutPanel1"
        Me.FlowLayoutPanel1.Size = New System.Drawing.Size(228, 39)
        Me.FlowLayoutPanel1.TabIndex = 7
        '
        'eleName
        '
        Me.eleName.AutoSize = True
        Me.eleName.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.eleName.Location = New System.Drawing.Point(3, 0)
        Me.eleName.Name = "eleName"
        Me.eleName.Size = New System.Drawing.Size(49, 17)
        Me.eleName.TabIndex = 0
        Me.eleName.Text = "Name"
        '
        'eleType
        '
        Me.eleType.AutoSize = True
        Me.eleType.Location = New System.Drawing.Point(3, 20)
        Me.eleType.Margin = New System.Windows.Forms.Padding(3, 3, 3, 0)
        Me.eleType.Name = "eleType"
        Me.eleType.Size = New System.Drawing.Size(191, 13)
        Me.eleType.TabIndex = 4
        Me.eleType.Text = "Objekttyp (zb. Meilenstein, Phase, etc.)"
        '
        'RichTextBox1
        '
        Me.RichTextBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RichTextBox1.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.RichTextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TableLayoutPanel1.SetColumnSpan(Me.RichTextBox1, 2)
        Me.RichTextBox1.Location = New System.Drawing.Point(3, 316)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.Size = New System.Drawing.Size(223, 64)
        Me.RichTextBox1.TabIndex = 12
        Me.RichTextBox1.Text = resources.GetString("RichTextBox1.Text")
        '
        'ucInfo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.FlowLayoutPanel1)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.MinimumSize = New System.Drawing.Size(232, 0)
        Me.Name = "ucInfo"
        Me.Size = New System.Drawing.Size(232, 520)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.FlowLayoutPanel1.ResumeLayout(False)
        Me.FlowLayoutPanel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents labelDate As Windows.Forms.Label
    Friend WithEvents labelRespons As Windows.Forms.Label
    Friend WithEvents labelAmpel As Windows.Forms.Label
    Friend WithEvents TableLayoutPanel1 As Windows.Forms.TableLayoutPanel
    Friend WithEvents eleAmpelText As Windows.Forms.Label
    Friend WithEvents eleDeliverables As Windows.Forms.TextBox
    Friend WithEvents FlowLayoutPanel1 As Windows.Forms.FlowLayoutPanel
    Friend WithEvents eleName As Windows.Forms.Label
    Friend WithEvents eleType As Windows.Forms.Label
    Friend WithEvents labelDeliver As Windows.Forms.Label
    Friend WithEvents eleDatum As Windows.Forms.Label
    Friend WithEvents eleRespons As Windows.Forms.Label
    Friend WithEvents eleAmpel As Windows.Forms.Label
    Friend WithEvents RichTextBox1 As Windows.Forms.RichTextBox
End Class
