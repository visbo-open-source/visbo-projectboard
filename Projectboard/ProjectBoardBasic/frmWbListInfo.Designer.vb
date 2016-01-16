<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmWbListInfo
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
        Me.elementName = New System.Windows.Forms.Label()
        Me.headerText = New System.Windows.Forms.Label()
        Me.ergebnisListe = New System.Windows.Forms.ListBox()
        Me.deleteButton = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'elementName
        '
        Me.elementName.AutoSize = True
        Me.elementName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.elementName.Location = New System.Drawing.Point(9, 22)
        Me.elementName.Name = "elementName"
        Me.elementName.Size = New System.Drawing.Size(98, 16)
        Me.elementName.TabIndex = 0
        Me.elementName.Text = "Element-Name"
        '
        'headerText
        '
        Me.headerText.AutoSize = True
        Me.headerText.Location = New System.Drawing.Point(13, 59)
        Me.headerText.Name = "headerText"
        Me.headerText.Size = New System.Drawing.Size(192, 13)
        Me.headerText.TabIndex = 1
        Me.headerText.Text = "dieses Element hat folgende Synonyme"
        '
        'ergebnisListe
        '
        Me.ergebnisListe.FormattingEnabled = True
        Me.ergebnisListe.HorizontalScrollbar = True
        Me.ergebnisListe.Location = New System.Drawing.Point(12, 80)
        Me.ergebnisListe.Name = "ergebnisListe"
        Me.ergebnisListe.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.ergebnisListe.Size = New System.Drawing.Size(260, 173)
        Me.ergebnisListe.Sorted = True
        Me.ergebnisListe.TabIndex = 2
        '
        'deleteButton
        '
        Me.deleteButton.BackColor = System.Drawing.SystemColors.Control
        Me.deleteButton.Enabled = False
        Me.deleteButton.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.deleteButton.Location = New System.Drawing.Point(37, 274)
        Me.deleteButton.Name = "deleteButton"
        Me.deleteButton.Size = New System.Drawing.Size(196, 27)
        Me.deleteButton.TabIndex = 3
        Me.deleteButton.Text = "Regel aus Wörterbuch entfernen"
        Me.deleteButton.UseVisualStyleBackColor = False
        '
        'frmWbListInfo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 322)
        Me.Controls.Add(Me.deleteButton)
        Me.Controls.Add(Me.ergebnisListe)
        Me.Controls.Add(Me.headerText)
        Me.Controls.Add(Me.elementName)
        Me.Name = "frmWbListInfo"
        Me.Text = "Information"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents elementName As System.Windows.Forms.Label
    Friend WithEvents headerText As System.Windows.Forms.Label
    Friend WithEvents ergebnisListe As System.Windows.Forms.ListBox
    Friend WithEvents deleteButton As System.Windows.Forms.Button
End Class
