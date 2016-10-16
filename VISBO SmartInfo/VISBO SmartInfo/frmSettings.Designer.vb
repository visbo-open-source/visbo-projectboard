<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSettings
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
        Me.protectShapes = New System.Windows.Forms.CheckBox()
        Me.extendedSearch = New System.Windows.Forms.CheckBox()
        Me.showInfoBC = New System.Windows.Forms.CheckBox()
        Me.dbLoginButton = New System.Windows.Forms.Button()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.abstandseinheit = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.dbURI = New System.Windows.Forms.TextBox()
        Me.dbName = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.schriftSize = New System.Windows.Forms.TextBox()
        Me.lbl_schrift = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'protectShapes
        '
        Me.protectShapes.AutoSize = True
        Me.protectShapes.Location = New System.Drawing.Point(16, 98)
        Me.protectShapes.Name = "protectShapes"
        Me.protectShapes.Size = New System.Drawing.Size(170, 17)
        Me.protectShapes.TabIndex = 31
        Me.protectShapes.Text = "Smart PPT Elemente schützen"
        Me.protectShapes.UseVisualStyleBackColor = True
        '
        'extendedSearch
        '
        Me.extendedSearch.AutoSize = True
        Me.extendedSearch.Location = New System.Drawing.Point(168, 74)
        Me.extendedSearch.Name = "extendedSearch"
        Me.extendedSearch.Size = New System.Drawing.Size(106, 17)
        Me.extendedSearch.TabIndex = 30
        Me.extendedSearch.Text = "erweiterte Suche"
        Me.extendedSearch.UseVisualStyleBackColor = True
        '
        'showInfoBC
        '
        Me.showInfoBC.AutoSize = True
        Me.showInfoBC.Location = New System.Drawing.Point(16, 74)
        Me.showInfoBC.Name = "showInfoBC"
        Me.showInfoBC.Size = New System.Drawing.Size(129, 17)
        Me.showInfoBC.TabIndex = 29
        Me.showInfoBC.Text = "Breadcrumb anzeigen"
        Me.showInfoBC.UseVisualStyleBackColor = True
        '
        'dbLoginButton
        '
        Me.dbLoginButton.Location = New System.Drawing.Point(135, 291)
        Me.dbLoginButton.Name = "dbLoginButton"
        Me.dbLoginButton.Size = New System.Drawing.Size(235, 23)
        Me.dbLoginButton.TabIndex = 28
        Me.dbLoginButton.Text = "Datenbank Login"
        Me.dbLoginButton.UseVisualStyleBackColor = True
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(135, 238)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.TextBox3.Size = New System.Drawing.Size(235, 20)
        Me.TextBox3.TabIndex = 27
        Me.TextBox3.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.TextBox3.UseSystemPasswordChar = True
        Me.TextBox3.WordWrap = False
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(135, 215)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(235, 20)
        Me.TextBox2.TabIndex = 26
        Me.TextBox2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'abstandseinheit
        '
        Me.abstandseinheit.FormattingEnabled = True
        Me.abstandseinheit.Items.AddRange(New Object() {"Tagen", "Wochen", "Monaten"})
        Me.abstandseinheit.Location = New System.Drawing.Point(287, 45)
        Me.abstandseinheit.Name = "abstandseinheit"
        Me.abstandseinheit.Size = New System.Drawing.Size(83, 21)
        Me.abstandseinheit.TabIndex = 25
        Me.abstandseinheit.Text = "Tagen"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(13, 50)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(106, 13)
        Me.Label5.TabIndex = 24
        Me.Label5.Text = "Abstand anzeigen in:"
        '
        'dbURI
        '
        Me.dbURI.Location = New System.Drawing.Point(135, 171)
        Me.dbURI.Name = "dbURI"
        Me.dbURI.Size = New System.Drawing.Size(235, 20)
        Me.dbURI.TabIndex = 23
        Me.dbURI.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'dbName
        '
        Me.dbName.Location = New System.Drawing.Point(135, 149)
        Me.dbName.Name = "dbName"
        Me.dbName.Size = New System.Drawing.Size(235, 20)
        Me.dbName.TabIndex = 22
        Me.dbName.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(13, 241)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(56, 13)
        Me.Label4.TabIndex = 21
        Me.Label4.Text = "Password:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(13, 218)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(83, 13)
        Me.Label3.TabIndex = 20
        Me.Label3.Text = "Benutzer-Name:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(13, 175)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(104, 13)
        Me.Label2.TabIndex = 19
        Me.Label2.Text = "Datenbank-Adresse:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 153)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(94, 13)
        Me.Label1.TabIndex = 18
        Me.Label1.Text = "Datenbank-Name:"
        '
        'schriftSize
        '
        Me.schriftSize.Location = New System.Drawing.Point(287, 22)
        Me.schriftSize.Name = "schriftSize"
        Me.schriftSize.Size = New System.Drawing.Size(83, 20)
        Me.schriftSize.TabIndex = 17
        Me.schriftSize.Text = "8"
        Me.schriftSize.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lbl_schrift
        '
        Me.lbl_schrift.AutoSize = True
        Me.lbl_schrift.Location = New System.Drawing.Point(13, 26)
        Me.lbl_schrift.Name = "lbl_schrift"
        Me.lbl_schrift.Size = New System.Drawing.Size(67, 13)
        Me.lbl_schrift.TabIndex = 16
        Me.lbl_schrift.Text = "Schriftgröße:"
        '
        'frmSettings
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(536, 336)
        Me.Controls.Add(Me.protectShapes)
        Me.Controls.Add(Me.extendedSearch)
        Me.Controls.Add(Me.showInfoBC)
        Me.Controls.Add(Me.dbLoginButton)
        Me.Controls.Add(Me.TextBox3)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.abstandseinheit)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.dbURI)
        Me.Controls.Add(Me.dbName)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.schriftSize)
        Me.Controls.Add(Me.lbl_schrift)
        Me.Name = "frmSettings"
        Me.Text = "Einstellungen"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents protectShapes As System.Windows.Forms.CheckBox
    Friend WithEvents extendedSearch As System.Windows.Forms.CheckBox
    Friend WithEvents showInfoBC As System.Windows.Forms.CheckBox
    Friend WithEvents dbLoginButton As System.Windows.Forms.Button
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents abstandseinheit As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents dbURI As System.Windows.Forms.TextBox
    Friend WithEvents dbName As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents schriftSize As System.Windows.Forms.TextBox
    Friend WithEvents lbl_schrift As System.Windows.Forms.Label
End Class
