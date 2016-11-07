<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmStoreReportProfil
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
        Me.NameReportProfil = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.AbbruchButton = New System.Windows.Forms.Button()
        Me.descLabel = New System.Windows.Forms.Label()
        Me.nameLabel = New System.Windows.Forms.Label()
        Me.profilDescription = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'NameReportProfil
        '
        Me.NameReportProfil.FormattingEnabled = True
        Me.NameReportProfil.Location = New System.Drawing.Point(15, 78)
        Me.NameReportProfil.Name = "NameReportProfil"
        Me.NameReportProfil.Size = New System.Drawing.Size(470, 21)
        Me.NameReportProfil.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 37)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(384, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Bitte geben Sie Name und Beschreibung für das zu speichernde ReportProfil ein"
        '
        'OKButton
        '
        Me.OKButton.Location = New System.Drawing.Point(15, 255)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(184, 23)
        Me.OKButton.TabIndex = 3
        Me.OKButton.Text = "Speichern"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'AbbruchButton
        '
        Me.AbbruchButton.Location = New System.Drawing.Point(301, 255)
        Me.AbbruchButton.Name = "AbbruchButton"
        Me.AbbruchButton.Size = New System.Drawing.Size(184, 23)
        Me.AbbruchButton.TabIndex = 4
        Me.AbbruchButton.Text = "Abbrechen"
        Me.AbbruchButton.UseVisualStyleBackColor = True
        '
        'descLabel
        '
        Me.descLabel.AutoSize = True
        Me.descLabel.Location = New System.Drawing.Point(12, 119)
        Me.descLabel.Name = "descLabel"
        Me.descLabel.Size = New System.Drawing.Size(72, 13)
        Me.descLabel.TabIndex = 5
        Me.descLabel.Text = "Beschreibung"
        '
        'nameLabel
        '
        Me.nameLabel.AutoSize = True
        Me.nameLabel.Location = New System.Drawing.Point(12, 62)
        Me.nameLabel.Name = "nameLabel"
        Me.nameLabel.Size = New System.Drawing.Size(35, 13)
        Me.nameLabel.TabIndex = 6
        Me.nameLabel.Text = "Name"
        '
        'profilDescription
        '
        Me.profilDescription.Location = New System.Drawing.Point(15, 135)
        Me.profilDescription.Multiline = True
        Me.profilDescription.Name = "profilDescription"
        Me.profilDescription.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.profilDescription.Size = New System.Drawing.Size(470, 93)
        Me.profilDescription.TabIndex = 2
        '
        'frmStoreReportProfil
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(497, 290)
        Me.Controls.Add(Me.profilDescription)
        Me.Controls.Add(Me.nameLabel)
        Me.Controls.Add(Me.descLabel)
        Me.Controls.Add(Me.AbbruchButton)
        Me.Controls.Add(Me.OKButton)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.NameReportProfil)
        Me.Name = "frmStoreReportProfil"
        Me.Text = "Aktuelles ReportProfil speichern"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents NameReportProfil As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents OKButton As System.Windows.Forms.Button
    Friend WithEvents AbbruchButton As System.Windows.Forms.Button
    Friend WithEvents descLabel As System.Windows.Forms.Label
    Friend WithEvents nameLabel As System.Windows.Forms.Label
    Friend WithEvents profilDescription As System.Windows.Forms.TextBox
End Class
