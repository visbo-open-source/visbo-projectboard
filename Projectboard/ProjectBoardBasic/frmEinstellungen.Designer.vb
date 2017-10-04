<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEinstellungen
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmEinstellungen))
        Me.chkboxMassEdit = New System.Windows.Forms.CheckBox()
        Me.chkboxPropAnpass = New System.Windows.Forms.CheckBox()
        Me.chkboxAmpel = New System.Windows.Forms.CheckBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SprachAusw = New System.Windows.Forms.ComboBox()
        Me.statusLabel = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'chkboxMassEdit
        '
        Me.chkboxMassEdit.AutoSize = True
        Me.chkboxMassEdit.Location = New System.Drawing.Point(27, 33)
        Me.chkboxMassEdit.Name = "chkboxMassEdit"
        Me.chkboxMassEdit.Size = New System.Drawing.Size(170, 17)
        Me.chkboxMassEdit.TabIndex = 0
        Me.chkboxMassEdit.Text = "Mass-Edit extended table view"
        Me.chkboxMassEdit.UseVisualStyleBackColor = True
        '
        'chkboxPropAnpass
        '
        Me.chkboxPropAnpass.AutoSize = True
        Me.chkboxPropAnpass.Location = New System.Drawing.Point(27, 70)
        Me.chkboxPropAnpass.Name = "chkboxPropAnpass"
        Me.chkboxPropAnpass.Size = New System.Drawing.Size(230, 17)
        Me.chkboxPropAnpass.TabIndex = 1
        Me.chkboxPropAnpass.Text = "Ressourcen-Bedarfe proportional anpassen"
        Me.chkboxPropAnpass.UseVisualStyleBackColor = True
        '
        'chkboxAmpel
        '
        Me.chkboxAmpel.AutoSize = True
        Me.chkboxAmpel.Location = New System.Drawing.Point(27, 105)
        Me.chkboxAmpel.Name = "chkboxAmpel"
        Me.chkboxAmpel.Size = New System.Drawing.Size(101, 17)
        Me.chkboxAmpel.TabIndex = 2
        Me.chkboxAmpel.Text = "Ampel anzeigen"
        Me.chkboxAmpel.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!)
        Me.Label1.Location = New System.Drawing.Point(24, 141)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(116, 15)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Sprache für Reports"
        '
        'SprachAusw
        '
        Me.SprachAusw.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.SprachAusw.FormattingEnabled = True
        Me.SprachAusw.Location = New System.Drawing.Point(146, 141)
        Me.SprachAusw.MaxDropDownItems = 4
        Me.SprachAusw.Name = "SprachAusw"
        Me.SprachAusw.Size = New System.Drawing.Size(158, 21)
        Me.SprachAusw.TabIndex = 5
        '
        'statusLabel
        '
        Me.statusLabel.AutoSize = True
        Me.statusLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.statusLabel.Location = New System.Drawing.Point(24, 176)
        Me.statusLabel.Name = "statusLabel"
        Me.statusLabel.Size = New System.Drawing.Size(45, 15)
        Me.statusLabel.TabIndex = 45
        Me.statusLabel.Text = "Label1"
        Me.statusLabel.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        Me.statusLabel.Visible = False
        '
        'frmEinstellungen
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(335, 198)
        Me.Controls.Add(Me.statusLabel)
        Me.Controls.Add(Me.SprachAusw)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.chkboxAmpel)
        Me.Controls.Add(Me.chkboxPropAnpass)
        Me.Controls.Add(Me.chkboxMassEdit)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmEinstellungen"
        Me.Text = "Einstellungen"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents chkboxMassEdit As System.Windows.Forms.CheckBox
    Friend WithEvents chkboxPropAnpass As System.Windows.Forms.CheckBox
    Friend WithEvents chkboxAmpel As System.Windows.Forms.CheckBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents SprachAusw As System.Windows.Forms.ComboBox
    Friend WithEvents statusLabel As System.Windows.Forms.Label
End Class
