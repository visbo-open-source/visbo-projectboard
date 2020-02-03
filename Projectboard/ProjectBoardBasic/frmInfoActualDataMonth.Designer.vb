<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmInfoActualDataMonth
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmInfoActualDataMonth))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.MonatJahr = New System.Windows.Forms.DateTimePicker()
        Me.okBtn = New System.Windows.Forms.Button()
        Me.cancelBtn = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(191, 15)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Ist-Daten bis einschließlich Monat"
        '
        'MonatJahr
        '
        Me.MonatJahr.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.MonatJahr.Location = New System.Drawing.Point(209, 24)
        Me.MonatJahr.MaxDate = New Date(2020, 1, 29, 0, 0, 0, 0)
        Me.MonatJahr.Name = "MonatJahr"
        Me.MonatJahr.Size = New System.Drawing.Size(82, 20)
        Me.MonatJahr.TabIndex = 2
        Me.MonatJahr.Value = New Date(2020, 1, 29, 0, 0, 0, 0)
        '
        'okBtn
        '
        Me.okBtn.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.okBtn.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.okBtn.Location = New System.Drawing.Point(15, 55)
        Me.okBtn.Name = "okBtn"
        Me.okBtn.Size = New System.Drawing.Size(95, 23)
        Me.okBtn.TabIndex = 5
        Me.okBtn.Text = "Import Daten"
        Me.okBtn.UseVisualStyleBackColor = False
        '
        'cancelBtn
        '
        Me.cancelBtn.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.cancelBtn.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cancelBtn.Location = New System.Drawing.Point(196, 55)
        Me.cancelBtn.Name = "cancelBtn"
        Me.cancelBtn.Size = New System.Drawing.Size(95, 23)
        Me.cancelBtn.TabIndex = 6
        Me.cancelBtn.Text = "Cancel"
        Me.cancelBtn.UseVisualStyleBackColor = False
        '
        'frmInfoActualDataMonth
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.ClientSize = New System.Drawing.Size(303, 90)
        Me.Controls.Add(Me.cancelBtn)
        Me.Controls.Add(Me.okBtn)
        Me.Controls.Add(Me.MonatJahr)
        Me.Controls.Add(Me.Label1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmInfoActualDataMonth"
        Me.Text = "Import Istdaten"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents okBtn As Windows.Forms.Button
    Friend WithEvents cancelBtn As Windows.Forms.Button
    Public WithEvents MonatJahr As Windows.Forms.DateTimePicker
End Class
