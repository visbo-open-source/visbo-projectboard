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
        Me.SuspendLayout()
        '
        'NameReportProfil
        '
        Me.NameReportProfil.FormattingEnabled = True
        Me.NameReportProfil.Location = New System.Drawing.Point(34, 77)
        Me.NameReportProfil.Name = "NameReportProfil"
        Me.NameReportProfil.Size = New System.Drawing.Size(428, 21)
        Me.NameReportProfil.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(31, 39)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(322, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Bitte geben Sie den Namen für das zu speichernde ReportProfil ein"
        '
        'OKButton
        '
        Me.OKButton.Location = New System.Drawing.Point(34, 133)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(184, 23)
        Me.OKButton.TabIndex = 3
        Me.OKButton.Text = "Speichern"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'AbbruchButton
        '
        Me.AbbruchButton.Location = New System.Drawing.Point(278, 133)
        Me.AbbruchButton.Name = "AbbruchButton"
        Me.AbbruchButton.Size = New System.Drawing.Size(184, 23)
        Me.AbbruchButton.TabIndex = 4
        Me.AbbruchButton.Text = "Abbrechen"
        Me.AbbruchButton.UseVisualStyleBackColor = True
        '
        'frmStoreReportProfil
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(497, 185)
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
End Class
