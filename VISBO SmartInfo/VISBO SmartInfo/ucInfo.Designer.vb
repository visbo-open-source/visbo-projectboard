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
        Me.labelDeliver = New System.Windows.Forms.Label()
        Me.eleDatum = New System.Windows.Forms.Label()
        Me.eleRespons = New System.Windows.Forms.Label()
        Me.eleDeliverables = New System.Windows.Forms.RichTextBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.percentDone = New System.Windows.Forms.Label()
        Me.eleAmpel = New System.Windows.Forms.Label()
        Me.FlowLayoutPanel1 = New System.Windows.Forms.FlowLayoutPanel()
        Me.eleName = New System.Windows.Forms.Label()
        Me.eleType = New System.Windows.Forms.Label()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.FlowLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'labelDate
        '
        Me.labelDate.AutoSize = True
        Me.labelDate.Location = New System.Drawing.Point(4, 10)
        Me.labelDate.Margin = New System.Windows.Forms.Padding(4, 10, 4, 0)
        Me.labelDate.Name = "labelDate"
        Me.labelDate.Size = New System.Drawing.Size(53, 17)
        Me.labelDate.TabIndex = 3
        Me.labelDate.Text = "Datum:"
        '
        'labelRespons
        '
        Me.labelRespons.AutoSize = True
        Me.labelRespons.Location = New System.Drawing.Point(4, 37)
        Me.labelRespons.Margin = New System.Windows.Forms.Padding(4, 10, 4, 0)
        Me.labelRespons.Name = "labelRespons"
        Me.labelRespons.Size = New System.Drawing.Size(75, 17)
        Me.labelRespons.TabIndex = 4
        Me.labelRespons.Text = "Zuständig:"
        '
        'labelAmpel
        '
        Me.labelAmpel.AutoSize = True
        Me.labelAmpel.Location = New System.Drawing.Point(4, 64)
        Me.labelAmpel.Margin = New System.Windows.Forms.Padding(4, 10, 4, 0)
        Me.labelAmpel.Name = "labelAmpel"
        Me.labelAmpel.Size = New System.Drawing.Size(51, 17)
        Me.labelAmpel.TabIndex = 5
        Me.labelAmpel.Text = "Ampel:"
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.AutoScrollMargin = New System.Drawing.Size(10, 10)
        Me.TableLayoutPanel1.BackColor = System.Drawing.Color.White
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 107.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel1.Controls.Add(Me.labelAmpel, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.labelRespons, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.labelDate, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.eleAmpelText, 0, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.labelDeliver, 0, 5)
        Me.TableLayoutPanel1.Controls.Add(Me.eleDatum, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.eleRespons, 1, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.eleDeliverables, 0, 6)
        Me.TableLayoutPanel1.Controls.Add(Me.Panel1, 1, 2)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(4, 53)
        Me.TableLayoutPanel1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 8
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 25.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 25.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(305, 583)
        Me.TableLayoutPanel1.TabIndex = 6
        '
        'eleAmpelText
        '
        Me.eleAmpelText.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.eleAmpelText.AutoSize = True
        Me.eleAmpelText.BackColor = System.Drawing.Color.Transparent
        Me.TableLayoutPanel1.SetColumnSpan(Me.eleAmpelText, 2)
        Me.eleAmpelText.Location = New System.Drawing.Point(4, 94)
        Me.eleAmpelText.Margin = New System.Windows.Forms.Padding(4, 10, 4, 0)
        Me.eleAmpelText.MaximumSize = New System.Drawing.Size(293, 0)
        Me.eleAmpelText.Name = "eleAmpelText"
        Me.eleAmpelText.Size = New System.Drawing.Size(293, 136)
        Me.eleAmpelText.TabIndex = 6
        Me.eleAmpelText.Text = resources.GetString("eleAmpelText.Text")
        '
        'labelDeliver
        '
        Me.labelDeliver.AutoSize = True
        Me.labelDeliver.Location = New System.Drawing.Point(4, 240)
        Me.labelDeliver.Margin = New System.Windows.Forms.Padding(4, 10, 4, 0)
        Me.labelDeliver.Name = "labelDeliver"
        Me.labelDeliver.Size = New System.Drawing.Size(90, 17)
        Me.labelDeliver.TabIndex = 8
        Me.labelDeliver.Text = "Deliverables:"
        '
        'eleDatum
        '
        Me.eleDatum.AutoSize = True
        Me.eleDatum.Location = New System.Drawing.Point(111, 10)
        Me.eleDatum.Margin = New System.Windows.Forms.Padding(4, 10, 4, 0)
        Me.eleDatum.Name = "eleDatum"
        Me.eleDatum.Size = New System.Drawing.Size(51, 17)
        Me.eleDatum.TabIndex = 9
        Me.eleDatum.Text = "Label8"
        '
        'eleRespons
        '
        Me.eleRespons.AutoSize = True
        Me.eleRespons.Location = New System.Drawing.Point(111, 37)
        Me.eleRespons.Margin = New System.Windows.Forms.Padding(4, 10, 4, 0)
        Me.eleRespons.Name = "eleRespons"
        Me.eleRespons.Size = New System.Drawing.Size(51, 17)
        Me.eleRespons.TabIndex = 10
        Me.eleRespons.Text = "Label9"
        '
        'eleDeliverables
        '
        Me.eleDeliverables.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.eleDeliverables.BackColor = System.Drawing.Color.White
        Me.eleDeliverables.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TableLayoutPanel1.SetColumnSpan(Me.eleDeliverables, 2)
        Me.eleDeliverables.Location = New System.Drawing.Point(4, 261)
        Me.eleDeliverables.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.eleDeliverables.MaximumSize = New System.Drawing.Size(0, 246)
        Me.eleDeliverables.Name = "eleDeliverables"
        Me.eleDeliverables.ReadOnly = True
        Me.eleDeliverables.Size = New System.Drawing.Size(298, 121)
        Me.eleDeliverables.TabIndex = 12
        Me.eleDeliverables.Text = resources.GetString("eleDeliverables.Text")
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.Controls.Add(Me.percentDone)
        Me.Panel1.Controls.Add(Me.eleAmpel)
        Me.Panel1.Location = New System.Drawing.Point(111, 58)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(191, 22)
        Me.Panel1.TabIndex = 13
        '
        'percentDone
        '
        Me.percentDone.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.percentDone.AutoSize = True
        Me.percentDone.Location = New System.Drawing.Point(103, 6)
        Me.percentDone.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.percentDone.Name = "percentDone"
        Me.percentDone.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.percentDone.Size = New System.Drawing.Size(56, 17)
        Me.percentDone.TabIndex = 12
        Me.percentDone.Text = "% done"
        Me.percentDone.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'eleAmpel
        '
        Me.eleAmpel.BackColor = System.Drawing.Color.DarkGray
        Me.eleAmpel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.eleAmpel.Location = New System.Drawing.Point(4, 0)
        Me.eleAmpel.Margin = New System.Windows.Forms.Padding(4, 4, 4, 0)
        Me.eleAmpel.Name = "eleAmpel"
        Me.eleAmpel.Size = New System.Drawing.Size(23, 22)
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
        Me.FlowLayoutPanel1.Location = New System.Drawing.Point(4, 2)
        Me.FlowLayoutPanel1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.FlowLayoutPanel1.Name = "FlowLayoutPanel1"
        Me.FlowLayoutPanel1.Size = New System.Drawing.Size(304, 48)
        Me.FlowLayoutPanel1.TabIndex = 7
        '
        'eleName
        '
        Me.eleName.AutoSize = True
        Me.eleName.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.eleName.Location = New System.Drawing.Point(4, 0)
        Me.eleName.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.eleName.Name = "eleName"
        Me.eleName.Size = New System.Drawing.Size(57, 20)
        Me.eleName.TabIndex = 0
        Me.eleName.Text = "Name"
        '
        'eleType
        '
        Me.eleType.AutoSize = True
        Me.eleType.Location = New System.Drawing.Point(4, 24)
        Me.eleType.Margin = New System.Windows.Forms.Padding(4, 4, 4, 0)
        Me.eleType.Name = "eleType"
        Me.eleType.Size = New System.Drawing.Size(255, 17)
        Me.eleType.TabIndex = 4
        Me.eleType.Text = "Objekttyp (zb. Meilenstein, Phase, etc.)"
        '
        'ucInfo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.FlowLayoutPanel1)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MinimumSize = New System.Drawing.Size(309, 0)
        Me.Name = "ucInfo"
        Me.Size = New System.Drawing.Size(309, 640)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
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
    Friend WithEvents FlowLayoutPanel1 As Windows.Forms.FlowLayoutPanel
    Friend WithEvents eleName As Windows.Forms.Label
    Friend WithEvents eleType As Windows.Forms.Label
    Friend WithEvents labelDeliver As Windows.Forms.Label
    Friend WithEvents eleDatum As Windows.Forms.Label
    Friend WithEvents eleRespons As Windows.Forms.Label
    Friend WithEvents eleAmpel As Windows.Forms.Label
    Friend WithEvents eleDeliverables As Windows.Forms.RichTextBox
    Friend WithEvents Panel1 As Windows.Forms.Panel
    Friend WithEvents percentDone As Windows.Forms.Label
End Class
