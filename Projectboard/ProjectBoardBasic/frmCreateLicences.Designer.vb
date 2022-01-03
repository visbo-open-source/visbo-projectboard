<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCreateLicences
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCreateLicences))
        Me.untilDate = New System.Windows.Forms.DateTimePicker()
        Me.labelUntilDate = New System.Windows.Forms.Label()
        Me.LabelUser = New System.Windows.Forms.Label()
        Me.LabelKomponente = New System.Windows.Forms.Label()
        Me.ListKomponenten = New System.Windows.Forms.ListBox()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.UserName = New System.Windows.Forms.TextBox()
        Me.AddLicences = New System.Windows.Forms.Button()
        Me.FileNameUserList = New System.Windows.Forms.TextBox()
        Me.LabelUserList = New System.Windows.Forms.Label()
        Me.SaveButton = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'untilDate
        '
        Me.untilDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.untilDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.untilDate.Location = New System.Drawing.Point(186, 292)
        Me.untilDate.Name = "untilDate"
        Me.untilDate.Size = New System.Drawing.Size(103, 22)
        Me.untilDate.TabIndex = 0
        '
        'labelUntilDate
        '
        Me.labelUntilDate.AutoSize = True
        Me.labelUntilDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelUntilDate.Location = New System.Drawing.Point(29, 298)
        Me.labelUntilDate.Name = "labelUntilDate"
        Me.labelUntilDate.Size = New System.Drawing.Size(67, 16)
        Me.labelUntilDate.TabIndex = 1
        Me.labelUntilDate.Text = "gültig bis: "
        '
        'LabelUser
        '
        Me.LabelUser.AutoSize = True
        Me.LabelUser.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelUser.Location = New System.Drawing.Point(22, 106)
        Me.LabelUser.Name = "LabelUser"
        Me.LabelUser.Size = New System.Drawing.Size(74, 16)
        Me.LabelUser.TabIndex = 2
        Me.LabelUser.Text = "Username:"
        '
        'LabelKomponente
        '
        Me.LabelKomponente.AutoSize = True
        Me.LabelKomponente.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelKomponente.Location = New System.Drawing.Point(22, 157)
        Me.LabelKomponente.Name = "LabelKomponente"
        Me.LabelKomponente.Size = New System.Drawing.Size(139, 16)
        Me.LabelKomponente.TabIndex = 3
        Me.LabelKomponente.Text = "SoftwareKomponente:"
        '
        'ListKomponenten
        '
        Me.ListKomponenten.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ListKomponenten.FormattingEnabled = True
        Me.ListKomponenten.HorizontalScrollbar = True
        Me.ListKomponenten.ItemHeight = 16
        Me.ListKomponenten.Location = New System.Drawing.Point(186, 157)
        Me.ListKomponenten.Name = "ListKomponenten"
        Me.ListKomponenten.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.ListKomponenten.Size = New System.Drawing.Size(282, 116)
        Me.ListKomponenten.Sorted = True
        Me.ListKomponenten.TabIndex = 4
        '
        'OKButton
        '
        Me.OKButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OKButton.Location = New System.Drawing.Point(32, 343)
        Me.OKButton.Name = "OKButton"
        Me.OKButton.Size = New System.Drawing.Size(140, 23)
        Me.OKButton.TabIndex = 6
        Me.OKButton.Text = "Create Licence-File"
        Me.OKButton.UseVisualStyleBackColor = True
        '
        'UserName
        '
        Me.UserName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UserName.Location = New System.Drawing.Point(186, 103)
        Me.UserName.Name = "UserName"
        Me.UserName.Size = New System.Drawing.Size(282, 22)
        Me.UserName.TabIndex = 7
        '
        'AddLicences
        '
        Me.AddLicences.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AddLicences.Location = New System.Drawing.Point(356, 343)
        Me.AddLicences.Name = "AddLicences"
        Me.AddLicences.Size = New System.Drawing.Size(155, 23)
        Me.AddLicences.TabIndex = 8
        Me.AddLicences.Text = "Add Licences"
        Me.AddLicences.UseVisualStyleBackColor = True
        '
        'FileNameUserList
        '
        Me.FileNameUserList.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FileNameUserList.Location = New System.Drawing.Point(186, 35)
        Me.FileNameUserList.Name = "FileNameUserList"
        Me.FileNameUserList.Size = New System.Drawing.Size(282, 22)
        Me.FileNameUserList.TabIndex = 9
        '
        'LabelUserList
        '
        Me.LabelUserList.AutoSize = True
        Me.LabelUserList.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelUserList.Location = New System.Drawing.Point(22, 35)
        Me.LabelUserList.Name = "LabelUserList"
        Me.LabelUserList.Size = New System.Drawing.Size(69, 16)
        Me.LabelUserList.TabIndex = 10
        Me.LabelUserList.Text = "User-Liste"
        '
        'SaveButton
        '
        Me.SaveButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SaveButton.Location = New System.Drawing.Point(198, 343)
        Me.SaveButton.Name = "SaveButton"
        Me.SaveButton.Size = New System.Drawing.Size(140, 23)
        Me.SaveButton.TabIndex = 11
        Me.SaveButton.Text = "Save Licence-File"
        Me.SaveButton.UseVisualStyleBackColor = True
        '
        'frmCreateLicences
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(540, 378)
        Me.Controls.Add(Me.SaveButton)
        Me.Controls.Add(Me.LabelUserList)
        Me.Controls.Add(Me.FileNameUserList)
        Me.Controls.Add(Me.AddLicences)
        Me.Controls.Add(Me.UserName)
        Me.Controls.Add(Me.OKButton)
        Me.Controls.Add(Me.ListKomponenten)
        Me.Controls.Add(Me.LabelKomponente)
        Me.Controls.Add(Me.LabelUser)
        Me.Controls.Add(Me.labelUntilDate)
        Me.Controls.Add(Me.untilDate)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmCreateLicences"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Erzeugen der Lizenzen"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents untilDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents labelUntilDate As System.Windows.Forms.Label
    Friend WithEvents LabelUser As System.Windows.Forms.Label
    Friend WithEvents LabelKomponente As System.Windows.Forms.Label
    Friend WithEvents ListKomponenten As System.Windows.Forms.ListBox
    Friend WithEvents OKButton As System.Windows.Forms.Button
    Friend WithEvents UserName As System.Windows.Forms.TextBox
    Friend WithEvents AddLicences As System.Windows.Forms.Button
    Friend WithEvents FileNameUserList As System.Windows.Forms.TextBox
    Friend WithEvents LabelUserList As System.Windows.Forms.Label
    Friend WithEvents SaveButton As System.Windows.Forms.Button
End Class
