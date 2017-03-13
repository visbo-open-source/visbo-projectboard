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
        Me.lblOutput = New System.Windows.Forms.Label()
        Me.ListBoxOutput = New System.Windows.Forms.ListBox()
        Me.SuspendLayout()
        '
        'lblOutput
        '
        Me.lblOutput.AutoSize = True
        Me.lblOutput.Location = New System.Drawing.Point(12, 10)
        Me.lblOutput.Name = "lblOutput"
        Me.lblOutput.Size = New System.Drawing.Size(39, 13)
        Me.lblOutput.TabIndex = 0
        Me.lblOutput.Text = "Label1"
        '
        'ListBoxOutput
        '
        Me.ListBoxOutput.FormattingEnabled = True
        Me.ListBoxOutput.HorizontalExtent = 1024
        Me.ListBoxOutput.HorizontalScrollbar = True
        Me.ListBoxOutput.Location = New System.Drawing.Point(12, 44)
        Me.ListBoxOutput.Name = "ListBoxOutput"
        Me.ListBoxOutput.Size = New System.Drawing.Size(371, 329)
        Me.ListBoxOutput.TabIndex = 1
        '
        'frmOutputWindow
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(395, 397)
        Me.Controls.Add(Me.ListBoxOutput)
        Me.Controls.Add(Me.lblOutput)
        Me.Name = "frmOutputWindow"
        Me.Text = "Meldungen"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ListBoxOutput As System.Windows.Forms.ListBox
    Public WithEvents lblOutput As System.Windows.Forms.Label
End Class
