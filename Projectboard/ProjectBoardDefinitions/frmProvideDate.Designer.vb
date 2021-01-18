<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmProvideDate
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
        Me.newDateValue = New System.Windows.Forms.DateTimePicker()
        Me.ok_btn = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'newDateValue
        '
        Me.newDateValue.Location = New System.Drawing.Point(25, 21)
        Me.newDateValue.Name = "newDateValue"
        Me.newDateValue.Size = New System.Drawing.Size(200, 20)
        Me.newDateValue.TabIndex = 0
        '
        'ok_btn
        '
        Me.ok_btn.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.ok_btn.Location = New System.Drawing.Point(88, 57)
        Me.ok_btn.Name = "ok_btn"
        Me.ok_btn.Size = New System.Drawing.Size(75, 23)
        Me.ok_btn.TabIndex = 1
        Me.ok_btn.Text = "OK"
        Me.ok_btn.UseVisualStyleBackColor = True
        '
        'clsFrmProvideDate
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(251, 97)
        Me.Controls.Add(Me.ok_btn)
        Me.Controls.Add(Me.newDateValue)
        Me.Name = "clsFrmProvideDate"
        Me.Text = "Provide Date"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents newDateValue As Windows.Forms.DateTimePicker
    Friend WithEvents ok_btn As Windows.Forms.Button
End Class
