<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmProvideActualDataMonth
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.valueMonth = New System.Windows.Forms.TextBox()
        Me.createUnknownProjects = New System.Windows.Forms.CheckBox()
        Me.readPastAndFutureData = New System.Windows.Forms.CheckBox()
        Me.okBtn = New System.Windows.Forms.Button()
        Me.cancelBtn = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(15, 29)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(165, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Ist-Daten bis einschließlich Monat"
        '
        'valueMonth
        '
        Me.valueMonth.Location = New System.Drawing.Point(187, 26)
        Me.valueMonth.Name = "valueMonth"
        Me.valueMonth.Size = New System.Drawing.Size(66, 20)
        Me.valueMonth.TabIndex = 1
        Me.valueMonth.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'createUnknownProjects
        '
        Me.createUnknownProjects.AutoSize = True
        Me.createUnknownProjects.Location = New System.Drawing.Point(18, 64)
        Me.createUnknownProjects.Name = "createUnknownProjects"
        Me.createUnknownProjects.Size = New System.Drawing.Size(166, 17)
        Me.createUnknownProjects.TabIndex = 2
        Me.createUnknownProjects.Text = "unbekannte Projekte anlegen"
        Me.createUnknownProjects.UseVisualStyleBackColor = True
        '
        'readPastAndFutureData
        '
        Me.readPastAndFutureData.AutoSize = True
        Me.readPastAndFutureData.Location = New System.Drawing.Point(18, 87)
        Me.readPastAndFutureData.Name = "readPastAndFutureData"
        Me.readPastAndFutureData.Size = New System.Drawing.Size(232, 17)
        Me.readPastAndFutureData.TabIndex = 3
        Me.readPastAndFutureData.Text = "Auch (zukünftige Zuweisungs-) Daten lesen"
        Me.readPastAndFutureData.UseVisualStyleBackColor = True
        '
        'okBtn
        '
        Me.okBtn.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.okBtn.Location = New System.Drawing.Point(18, 131)
        Me.okBtn.Name = "okBtn"
        Me.okBtn.Size = New System.Drawing.Size(75, 23)
        Me.okBtn.TabIndex = 4
        Me.okBtn.Text = "Import Daten"
        Me.okBtn.UseVisualStyleBackColor = True
        '
        'cancelBtn
        '
        Me.cancelBtn.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cancelBtn.Location = New System.Drawing.Point(178, 131)
        Me.cancelBtn.Name = "cancelBtn"
        Me.cancelBtn.Size = New System.Drawing.Size(75, 23)
        Me.cancelBtn.TabIndex = 5
        Me.cancelBtn.Text = "Cancel"
        Me.cancelBtn.UseVisualStyleBackColor = True
        '
        'frmProvideActualDataMonth
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 179)
        Me.Controls.Add(Me.cancelBtn)
        Me.Controls.Add(Me.okBtn)
        Me.Controls.Add(Me.readPastAndFutureData)
        Me.Controls.Add(Me.createUnknownProjects)
        Me.Controls.Add(Me.valueMonth)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmProvideActualDataMonth"
        Me.Text = "Import Ist-Daten"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents okBtn As Windows.Forms.Button
    Friend WithEvents cancelBtn As Windows.Forms.Button
    Public WithEvents valueMonth As Windows.Forms.TextBox
    Public WithEvents createUnknownProjects As Windows.Forms.CheckBox
    Public WithEvents readPastAndFutureData As Windows.Forms.CheckBox
End Class
