<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmdemoModusDate
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmdemoModusDate))
        Me.kennzeichnungDate = New System.Windows.Forms.Label()
        Me.DateTimeHistory = New System.Windows.Forms.DateTimePicker()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.AbbrButton = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'kennzeichnungDate
        '
        Me.kennzeichnungDate.AutoSize = True
        Me.kennzeichnungDate.Enabled = False
        Me.kennzeichnungDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.kennzeichnungDate.Location = New System.Drawing.Point(20, 22)
        Me.kennzeichnungDate.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.kennzeichnungDate.Name = "kennzeichnungDate"
        Me.kennzeichnungDate.Size = New System.Drawing.Size(101, 16)
        Me.kennzeichnungDate.TabIndex = 10
        Me.kennzeichnungDate.Text = "Datum f. History"
        '
        'DateTimeHistory
        '
        Me.DateTimeHistory.Location = New System.Drawing.Point(144, 22)
        Me.DateTimeHistory.Margin = New System.Windows.Forms.Padding(2)
        Me.DateTimeHistory.Name = "DateTimeHistory"
        Me.DateTimeHistory.Size = New System.Drawing.Size(208, 20)
        Me.DateTimeHistory.TabIndex = 27
        '
        'OKButton
        '
        Me.OKButton.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.OKButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OKButton.Location = New System.Drawing.Point(23, 67)
        Me.OKButton.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(130, 22)
        Me.OKButton.TabIndex = 28
        Me.OKButton.Text = "OK"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'AbbrButton
        '
        Me.AbbrButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.AbbrButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AbbrButton.Location = New System.Drawing.Point(218, 67)
        Me.AbbrButton.Margin = New System.Windows.Forms.Padding(2, 3, 2, 3)
        Me.AbbrButton.Name = "AbbrButton"
        Me.AbbrButton.Size = New System.Drawing.Size(133, 22)
        Me.AbbrButton.TabIndex = 29
        Me.AbbrButton.Text = "Abbrechen"
        Me.AbbrButton.UseVisualStyleBackColor = True
        '
        'frmdemoModusDate
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.ClientSize = New System.Drawing.Size(403, 107)
        Me.Controls.Add(Me.AbbrButton)
        Me.Controls.Add(Me.OKButton)
        Me.Controls.Add(Me.DateTimeHistory)
        Me.Controls.Add(Me.kennzeichnungDate)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "frmdemoModusDate"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Datum für die Historie (Demo)"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents kennzeichnungDate As System.Windows.Forms.Label
    Public WithEvents DateTimeHistory As System.Windows.Forms.DateTimePicker
    Public WithEvents OKButton As System.Windows.Forms.Button
    Public WithEvents AbbrButton As System.Windows.Forms.Button
End Class
