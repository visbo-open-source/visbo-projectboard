<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmOutputWindow
    Inherits System.Windows.Forms.Form

    'Das Formular überschreibt den Löschvorgang, um die Komponentenliste zu bereinigen.
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmOutputWindow))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.LinkLabelKontakt = New System.Windows.Forms.LinkLabel()
        Me.LabelHL = New System.Windows.Forms.Label()
        Me.lblOutput = New System.Windows.Forms.Label()
        Me.ListBoxOutput = New System.Windows.Forms.ListBox()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.Controls.Add(Me.LinkLabelKontakt)
        Me.Panel1.Controls.Add(Me.LabelHL)
        Me.Panel1.Controls.Add(Me.lblOutput)
        Me.Panel1.Controls.Add(Me.ListBoxOutput)
        Me.Panel1.Location = New System.Drawing.Point(4, 0)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(2)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(369, 378)
        Me.Panel1.TabIndex = 2
        '
        'LinkLabelKontakt
        '
        Me.LinkLabelKontakt.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LinkLabelKontakt.AutoSize = True
        Me.LinkLabelKontakt.LinkBehavior = System.Windows.Forms.LinkBehavior.AlwaysUnderline
        Me.LinkLabelKontakt.Location = New System.Drawing.Point(112, 348)
        Me.LinkLabelKontakt.Name = "LinkLabelKontakt"
        Me.LinkLabelKontakt.Size = New System.Drawing.Size(124, 13)
        Me.LinkLabelKontakt.TabIndex = 5
        Me.LinkLabelKontakt.TabStop = True
        Me.LinkLabelKontakt.Text = "https://visbo.de/kontakt"
        '
        'LabelHL
        '
        Me.LabelHL.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LabelHL.AutoSize = True
        Me.LabelHL.Location = New System.Drawing.Point(8, 348)
        Me.LabelHL.Name = "LabelHL"
        Me.LabelHL.Size = New System.Drawing.Size(98, 13)
        Me.LabelHL.TabIndex = 4
        Me.LabelHL.Text = "Please contact us: "
        '
        'lblOutput
        '
        Me.lblOutput.AutoSize = True
        Me.lblOutput.Location = New System.Drawing.Point(3, 15)
        Me.lblOutput.Name = "lblOutput"
        Me.lblOutput.Size = New System.Drawing.Size(39, 13)
        Me.lblOutput.TabIndex = 3
        Me.lblOutput.Text = "Label1"
        '
        'ListBoxOutput
        '
        Me.ListBoxOutput.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ListBoxOutput.FormattingEnabled = True
        Me.ListBoxOutput.HorizontalExtent = 1024
        Me.ListBoxOutput.HorizontalScrollbar = True
        Me.ListBoxOutput.Location = New System.Drawing.Point(0, 48)
        Me.ListBoxOutput.Name = "ListBoxOutput"
        Me.ListBoxOutput.Size = New System.Drawing.Size(368, 290)
        Me.ListBoxOutput.TabIndex = 2
        '
        'frmOutputWindow
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(382, 383)
        Me.Controls.Add(Me.Panel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmOutputWindow"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Meldungen"
        Me.TopMost = True
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Panel1 As Windows.Forms.Panel
    Public WithEvents lblOutput As Windows.Forms.Label
    Friend WithEvents ListBoxOutput As Windows.Forms.ListBox
    Public WithEvents LabelHL As Windows.Forms.Label
    Friend WithEvents LinkLabelKontakt As Windows.Forms.LinkLabel
End Class
